

https://github.com/user-attachments/assets/3e51090d-53e9-4ee6-9a9a-02f03abbc51c

# 中航光电线束物料推荐、接线表以及对应CAXA绘图服务
本项目面向线束制造场景，提供从客户 PDF 图纸到内部接线表、物料清单和 CAXA 装配示意图的自动化转换能力，帮助生产人员快速、准确地理解装配方法。
## 背景与动机

在线束及连接器装配场景中，客户提供的原始图纸通常包含连接器型号、引脚关系、线缆规格、物料明细、节点距离、标签位置、扎带位置以及技术要求等信息。这些图纸面向客户侧设计表达，信息密度高，图面结构复杂，且不同客户在物料命名、图纸格式、标注习惯和接线关系表达方式上存在差异。

对于中航光电内部生产和装配人员而言，客户图纸中的物料信息往往不能直接对应到企业内部物料体系。工人如果直接依据客户图纸进行装配，需要人工理解复杂走线关系，手动识别连接器、线缆、扎带、标签、波纹管等部件，并将客户物料信息映射到本厂物料编码和工艺表达方式。这一过程依赖经验，效率较低，容易出现理解偏差、物料选型错误和装配指导不清晰等问题。

本项目旨在构建一套面向线束图纸的智能转换系统，将客户图纸自动转化为中航光电内部可直接使用的制造交付物。系统以客户 PDF 图纸为输入，通过 OCR 和多模态大模型识别图纸中的文本、表格、物料明细、接线关系和装配结构信息；随后结合企业规则库和物料知识库，完成客户物料到内部物料的匹配与转换，自动生成接线表和物料清单。

在此基础上，系统进一步理解图纸中的装配意图，根据接线表、物料清单和节点距离等结构化信息，生成可在 CAXA 中执行的 C++ 宏命令，自动绘制面向制造装配的简化示意图。该示意图重点表达连接器、节点、标签、扎带、距离和装配顺序等关键信息，用更直观的方式指导内部员工完成线束装配。

通过本项目，企业可以将“客户图纸理解、物料映射、接线表填写、物料清单生成、装配示意图绘制”等原本高度依赖人工经验的环节自动化、标准化和结构化，降低图纸理解门槛，减少人工转换错误，提高工艺准备效率，并为后续知识库沉淀、规则维护和智能制造流程集成提供数据基础。

## 运行链路
### 阶段一
1. 用户在前端上传 PDF 图纸或 DOCX 技术要求文件。
2. PDF 通过 OCR 服务抽取图纸文字和表格；DOCX 直接解析正文和表格。
3. 后端抽取图纸需求，可选择启用 LLM 辅助抽取。
4. 后端读取运行时治理知识库 `.xlsx`，匹配连接器、导线、防护辅材、扎带卡扣等物料。
5. 后端复制接线表模板，只保留模板格式，回填 `接线表` 和 `连接器及其所用辅材`，生成新 Excel。
### 阶段二
1. 上传阶段一生成的`接线表` 和 `连接器及其所用辅材`,以及PDF图纸给多模态大模型，理解绘制意图。
2. 大模型将提取到的物料信息已经模块的位置信息填入到semantic.json中，在json中用户可以检验大模型提取信息完整性和物料位置准确性。
3. 通过规则算法将semantic.json转为确定性布局的layout.json，便于后期定位到故障点。
4. layout.json转为.cpp，CAXA绘图宏命令，经过vs2010编译后导入CAXA软件，实现一键绘制。

## GitHub 副本包含
### 阶段一：
- `jonhon_web_app.py`：FastAPI 后端，负责上传、任务状态、OCR 调用、治理文件管理和下载。
- `jonhon_material_pipeline_no_fallback.py`：图纸需求抽取、知识库匹配、Excel 回填主逻辑。
- `paddle_service.py`：OCR 服务调用示例。
- `web_frontend/`：前端页面。
- `requirements.txt`：后端 Python 依赖。
- `.env.example`：环境变量示例。
- `docs/物料知识治理过程说明.md`：治理过程说明。
- `data/README.md`：运行时数据文件放置说明。
#### 运行时必须准备

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
### 阶段二：

```text
pdf_to_caxa_wiring/
├─ extract_pdf_pages.py             # PDF 转 PNG 页面
├─ recognize_wiring_semantics.py    # 调用阿里百炼识别接线语义
├─ generate_layout.py               # semantic.json 转 layout.json
├─ generate_caxa_cpp.py             # layout.json 转 CAXA C++ 宏代码
├─ block_catalog.example.json       # CAXA 图块名称、尺寸、布局规则示例
└─ output/                          # 中间结果和生成代码
```

#### 输入输出

主要输入：

- PDF 图纸
- CAXA 中已导入的图块
- `block_catalog.example.json` 中的图块映射与布局规则

主要输出：

- `semantic.json`：大模型识别出的连接器、导线、波纹管、辅材、尺寸标注等语义信息
- `layout.json`：CAD 示意图实体布局，包含插入点、缩放、绘制顺序和尺寸线位置
- `AutoDrawWiring.cpp`：可复制到 CAXA ObjectCRX 工程中的 C++ 绘图代码

## 使用方法


### 1. 渲染 PDF

```powershell
python .\pdf_to_caxa_wiring\extract_pdf_pages.py ".\用户提供图纸1.pdf" -o ".\pdf_to_caxa_wiring\output\pages" --dpi 200
```

输出：

```text
pdf_to_caxa_wiring/output/pages/manifest.json
pdf_to_caxa_wiring/output/pages/page_0001.png
```

### 2. 调用阿里百炼识别语义

先设置 API Key：

```powershell
$env:DASHSCOPE_API_KEY="你的阿里百炼API Key"
```

再运行：

```powershell
python .\pdf_to_caxa_wiring\recognize_wiring_semantics.py ".\pdf_to_caxa_wiring\output\pages\manifest.json" -o ".\pdf_to_caxa_wiring\output\semantic.json"
```

### 3. 生成 CAD 布局

```powershell
python .\pdf_to_caxa_wiring\generate_layout.py ".\pdf_to_caxa_wiring\output\semantic.json" -o ".\pdf_to_caxa_wiring\output\layout.json" --catalog ".\pdf_to_caxa_wiring\block_catalog.example.json"
```

### 4. 生成 CAXA C++ 宏代码

```powershell
python .\pdf_to_caxa_wiring\generate_caxa_cpp.py ".\pdf_to_caxa_wiring\output\layout.json" -o ".\pdf_to_caxa_wiring\output\AutoDrawWiring.cpp"
```

如果 layout 仍是草稿状态，可联调用：

```powershell
python .\pdf_to_caxa_wiring\generate_caxa_cpp.py ".\pdf_to_caxa_wiring\output\layout.json" -o ".\pdf_to_caxa_wiring\output\AutoDrawWiring.cpp" --allow-draft
```

## CAXA 使用说明

1. 确保 CAXA 当前图纸已导入所需图块，例如：
   - `WIRE`
   - `CORRUGATED_TUBE`
   - 连接器块
   - 扎带、端子、标签等辅材块
2. 将生成的 `AutoDrawWiring.cpp` 中代码合并到 ObjectCRX 工程的 `CrxEntryPoint.cpp`。
3. 在 `On_kInitAppMsg` 中注册命令：

```cpp
crxedRegCmds->addCommand(
    _T("HelloApp"),
    _T("GAutoDrawWiring"),
    _T("AutoDrawWiring"),
    ACRX_CMD_MODAL,
    &cmdAutoDrawWiring
);
```

4. 编译、加载 CRX 后，在 CAXA 命令行执行：

```text
AutoDrawWiring
```


## 阶段一示例


https://github.com/user-attachments/assets/f7aa779c-62fd-4470-a38e-b21ad03a411b

## 阶段二示例


https://github.com/user-attachments/assets/2b7c44a6-3c9c-43ab-a011-1b217d142a97

示例中所用到的图纸样本：
[用户提供图纸1.pdf](https://github.com/user-attachments/files/29748137/1.pdf)




