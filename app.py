"""
DocxFill API 服务 v3 - 精简版
"""
from fastapi import FastAPI, HTTPException, UploadFile, File, Form, Request
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import requests as req
import json
import re
import io
import os
import uuid
from docx import Document

app = FastAPI(title="DocxFill API", description="Word模板占位符替换服务")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

OUTPUT_DIR = "/tmp/docxfill_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


class FillRequest(BaseModel):
    template_url: str
    excel_data: str


PERCENT_FIELDS = {
    "商业用途占比", "住宅用途占比", "城镇住宅用途占比", "商业用地占比", "城镇住宅用地占比",
    "建筑密度", "绿地率", "附加税率", "开发利润率"
}

THOUSAND_SEP_FIELDS = {
    "总建筑面积", "计容建筑面积", "地下建筑面积", "建筑基底面积",
    "商业建筑面积", "一层商业建筑面积", "二层商业建筑面积", "住宅建筑面积",
    "房屋建筑安装工程费", "勘察设计和前期工程费", "宗地内基础设施建设费",
    "其他费用", "房屋建造成本", "管理费用", "销售费用",
    "投资利息1", "销项税额1", "建安进项", "前期进项",
    "其他费用进项", "销售进项", "抵扣合计", "增值税",
    "增值税附加1", "印花税", "合计税费1", "开发成本1",
    "开发利润1", "总地价"
}

FIELD_MAPPING = {
    "其他进项": "其他费用进项",
    "住宅用途占比": "城镇住宅用途占比",
}


def format_value(key, value):
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return "资料受限"
    if key in PERCENT_FIELDS:
        try:
            num = float(value)
            return f"{num}%" if num > 1 else f"{round(num * 100, 2)}%"
        except:
            return str(value)
    if key in THOUSAND_SEP_FIELDS:
        try:
            num = float(value)
            return f"{int(num):,}" if num == int(num) else f"{num:,.2f}"
        except:
            return str(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value != int(value):
            return f"{value:.2f}"
        return str(int(value)) if isinstance(value, float) else str(value)
    return str(value)


def resolve_key(key, data):
    key_stripped = key.strip()
    if key_stripped in data:
        return key_stripped, data[key_stripped]
    if key_stripped in FIELD_MAPPING:
        mapped = FIELD_MAPPING[key_stripped]
        if mapped in data:
            return mapped, data[mapped]
    for data_key in data:
        if data_key.strip() == key_stripped:
            return data_key, data[data_key]
    return None, None


def replace_in_doc(doc, data):
    count = 0
    pattern = re.compile(r'\{\{(.+?)\}\}')

    for para in doc.paragraphs:
        count += replace_in_para(para, data, pattern)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    count += replace_in_para(para, data, pattern)

    for section in doc.sections:
        for header in [section.header, section.first_page_header]:
            try:
                if header and not header.is_linked_to_previous:
                    for para in header.paragraphs:
                        count += replace_in_para(para, data, pattern)
            except:
                pass
        for footer in [section.footer, section.first_page_footer]:
            try:
                if footer and not footer.is_linked_to_previous:
                    for para in footer.paragraphs:
                        count += replace_in_para(para, data, pattern)
            except:
                pass
    return count


def replace_in_para(para, data, pattern):
    count = 0
    if not pattern.search(para.text):
        return 0

    for run in para.runs:
        if not run.text:
            continue
        for match in pattern.finditer(run.text):
            key = match.group(1)
            resolved_key, value = resolve_key(key, data)
            if resolved_key is not None:
                formatted = format_value(resolved_key, value)
                run.text = run.text.replace('{{' + key + '}}', formatted)
                count += 1

    # 跨run处理
    remaining = pattern.search(para.text)
    if remaining:
        count += handle_cross_run(para, data, pattern)
    return count


def handle_cross_run(para, data, pattern):
    count = 0
    runs = para.runs
    if not runs:
        return 0

    full_text = ''.join(r.text for r in runs)
    matches = list(pattern.finditer(full_text))
    if not matches:
        return 0

    char_map = []
    for i, r in enumerate(runs):
        for _ in r.text:
            char_map.append(i)

    for match in reversed(matches):
        s, e = match.start(), match.end()
        if s >= len(char_map) or e - 1 >= len(char_map):
            continue
        key = match.group(1)
        resolved_key, value = resolve_key(key, data)
        if resolved_key is None:
            continue
        formatted = format_value(resolved_key, value)

        fi = char_map[s]
        li = char_map[e - 1]

        if fi == li:
            run = runs[fi]
            rs = sum(len(runs[i].text) for i in range(fi))
            ps = s - rs
            pe = e - rs
            run.text = run.text[:ps] + formatted + run.text[pe:]
            count += 1
        else:
            fr = runs[fi]
            rs = sum(len(runs[i].text) for i in range(fi))
            ps = s - rs
            fr.text = fr.text[:ps] + formatted
            for i in range(fi + 1, li):
                runs[i].text = ''
            lr = runs[li]
            ls = sum(len(runs[i].text) for i in range(li))
            pe = e - ls
            if 0 <= pe <= len(lr.text):
                lr.text = lr.text[pe:]
            else:
                lr.text = ''
            count += 1
    return count


@app.post("/fill")
async def fill_template(request: FillRequest):
    try:
        resp = req.get(request.template_url, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        return {"success": False, "message": f"下载模板失败: {str(e)}", "replaced_count": 0, "file_id": ""}

    try:
        excel_rows = json.loads(request.excel_data)
    except Exception as e:
        return {"success": False, "message": f"JSON解析失败: {str(e)}", "replaced_count": 0, "file_id": ""}

    if not excel_rows or not isinstance(excel_rows, list):
        return {"success": False, "message": "excel_data必须是非空数组", "replaced_count": 0, "file_id": ""}

    data = excel_rows[0]
    for row in excel_rows[1:]:
        if isinstance(row, dict):
            for k, v in row.items():
                if k not in data:
                    data[k] = v

    doc = Document(io.BytesIO(resp.content))
    count = replace_in_doc(doc, data)

    file_id = str(uuid.uuid4())
    path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")
    doc.save(path)

    return {"success": True, "message": f"替换完成，共替换{count}个占位符", "replaced_count": count, "file_id": file_id}


@app.get("/download/{file_id}")
async def download_file(file_id: str):
    if "/" in file_id or ".." in file_id:
        raise HTTPException(status_code=400)
    path = os.path.join(OUTPUT_DIR, f"{file_id}.docx")
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename=f"filled.docx")


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/coze-openapi.json")
async def coze_openapi():
    """返回Coze兼容的OpenAPI 3.0.0文档"""
    return {
      "openapi": "3.0.0",
      "info": {
        "title": "DocxFill API",
        "description": "Word模板占位符替换服务，保留原格式",
        "version": "1.0.0"
      },
      "servers": [
        {"url": "https://docxfill-api-production.up.railway.app"}
      ],
      "paths": {
        "/fill": {
          "post": {
            "summary": "替换Word模板占位符",
            "description": "接收Word模板URL和Excel JSON数据，替换占位符后返回文件ID",
            "operationId": "fill_template",
            "requestBody": {
              "required": True,
              "content": {
                "application/json": {
                  "schema": {
                    "type": "object",
                    "required": ["template_url", "excel_data"],
                    "properties": {
                      "template_url": {
                        "type": "string",
                        "description": "Word模板文件的下载链接"
                      },
                      "excel_data": {
                        "type": "string",
                        "description": "Excel数据的JSON字符串"
                      }
                    }
                  }
                }
              }
            },
            "responses": {
              "200": {
                "description": "成功",
                "content": {
                  "application/json": {
                    "schema": {
                      "type": "object",
                      "properties": {
                        "success": {"type": "boolean", "description": "是否成功"},
                        "message": {"type": "string", "description": "结果消息"},
                        "replaced_count": {"type": "integer", "description": "替换的占位符数量"},
                        "file_id": {"type": "string", "description": "生成文件的ID，用于下载"}
                      }
                    }
                  }
                }
              }
            }
          }
        },
        "/download/{file_id}": {
          "get": {
            "summary": "下载替换后的Word文件",
            "description": "根据fill返回的file_id下载生成的docx文件",
            "operationId": "download_file",
            "parameters": [
              {
                "name": "file_id",
                "in": "path",
                "required": True,
                "schema": {"type": "string"},
                "description": "fill接口返回的文件ID"
              }
            ],
            "responses": {
              "200": {
                "description": "Word文件",
                "content": {
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {
                    "schema": {"type": "string", "format": "binary"}
                  }
                }
              }
            }
          }
        }
      }
    }