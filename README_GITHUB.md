# 中航光电线束物料推荐与接线表生成服务

## 运行链路

1. 用户在前端上传 PDF 图纸或 DOCX 技术要求文件。
2. PDF 通过 OCR 服务抽取图纸文字和表格；DOCX 直接解析正文和表格。
3. 后端抽取图纸需求，可选择启用 LLM 辅助抽取。
4. 后端读取运行时治理知识库 `.xlsx`，匹配连接器、导线、防护辅材、扎带卡扣等物料。
5. 后端复制接线表模板，只保留模板格式，回填 `接线表` 和 `连接器及其所用辅材`，生成新 Excel。

## GitHub 副本包含

- `jonhon_web_app.py`：FastAPI 后端，负责上传、任务状态、OCR 调用、治理文件管理和下载。
- `jonhon_material_pipeline_no_fallback.py`：图纸需求抽取、知识库匹配、Excel 回填主逻辑。
- `paddle_service.py`：OCR 服务调用示例。
- `web_frontend/`：前端页面。
- `requirements.txt`：后端 Python 依赖。
- `.env.example`：环境变量示例。
- `docs/物料知识治理过程说明.md`：治理过程说明。
- `data/README.md`：运行时数据文件放置说明。

## 运行时必须准备

项目根目录需要放置：

- `中航光电线束物料知识治理结果_final.xlsx`
- `线缆组件产品接线表格式模板20260420.xls`

服务依赖外部/本地能力：

- OCR 服务：默认由 `JONHON_OCR_URL` 指定。
- 可选 LLM：DashScope OpenAI-compatible API，使用 `DASHSCOPE_API_KEY`、`DASHSCOPE_BASE_URL`、`JONHON_WEB_LLM_MODEL` 配置。
- LibreOffice：用于转换 `.xls` 模板。
- 本地文件系统：保存上传文件、OCR 结果、任务输出和上传的治理文件。

## 启动

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python jonhon_web_app.py
```

默认监听 `0.0.0.0:8050`。可用环境变量修改：

```bash
export JONHON_WEB_PORT=18052
export JONHON_OCR_URL=http://127.0.0.1:8030/ocr
export JONHON_WEB_LLM_MODEL=qwen3.6-max-preview
```
