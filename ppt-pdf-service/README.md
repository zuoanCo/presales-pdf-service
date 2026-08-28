# PPT 模板 PDF 生成服务

接收参数 → 填充 PPT 模板中的 `{{占位符}}` → 导出与 PPT 排版完全一致的 PDF。

## 工作原理

1. **填充**：用 python-pptx 在 XML 级别替换文本，字体、字号、颜色、加粗、
   表格样式、图片、版式全部原样保留（占位符只是文本替换，不触碰布局）。
2. **转换**：两种后端自动选择（也可用环境变量 `PDF_CONVERTER` 强制指定）：

   | 后端 | 适用环境 | 保真度 |
   |------|---------|--------|
   | `powerpoint` | Windows + 已安装 Office | 100%（PowerPoint 自身渲染） |
   | `libreoffice` | 任意 Linux / Docker / Windows / macOS | 接近 100%，**需安装模板所用字体** |

## 快速开始

```bash
pip install -r requirements.txt
# Windows + Office 环境再装：pip install pywin32

# 把模板放进 templates/ 目录，用 {{参数名}} 标记可替换内容
python app.py          # 或 uvicorn app:app --port 8000
```

### 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/templates` | 列出可用模板 |
| GET | `/placeholders/{模板名}` | 扫描模板里有哪些占位符 |
| POST | `/render` | JSON 渲染：`{"template": "汇报模板.pptx", "params": {"姓名": "张三"}, "filename": "输出名"}` |
| POST | `/render-upload` | multipart 上传模板渲染：`file=xxx.pptx`，`params={"姓名":"张三"}` |

示例：

```bash
curl -X POST http://localhost:8000/render \
  -H "Content-Type: application/json" \
  -d '{"template":"汇报模板.pptx","params":{"姓名":"张三","日期":"2026-08-28"}}' \
  -o result.pdf
```

## 部署

### 方式一：Windows + Office（保真度 100%，推荐用于正式汇报）

直接 `python app.py` 即可，自动使用 PowerPoint 后端。
注意：以服务模式（如 Windows Service / 容器）运行时 PowerPoint COM 可能受限，
建议以普通用户进程方式运行。

### 方式二：任意 Linux 服务器 / Docker

```bash
docker build -t ppt-pdf-service .
docker run -p 8000:8000 -v $(pwd)/templates:/srv/templates ppt-pdf-service
```

**字体是保真度的关键**：模板里用了什么字体，服务器就必须装什么字体。
Dockerfile 已内置 Noto CJK（思源黑体/宋体）。若模板使用「微软雅黑」等商业字体，
需自行把 `.ttf/.ttc` 拷入镜像 `/usr/share/fonts/` 并执行 `fc-cache -f`，
否则 LibreOffice 会做字体替换，排版可能出现轻微偏移。

## 已接入的模板：售前模板.pptx

由 `prepare_template.py` 从原始文件生成（红框标记 → 占位符，红框已删除，原文件不动）。
共 16 个参数：

| 参数 | 位置 | 参数 | 位置 |
|------|------|------|------|
| 核心诊断 | P27 右侧总结 | 战略打法 | P29 |
| 战法 / 市场阶段 / 用户心智状态 / 企业资源要求 / 核心战略目标 | P27 表格黄行 | 产品名称 | P30 标题 |
| 行业趋势 | P28 | 原点用户 / 价值载体 / 价值心智 / 视觉锤语言钉 / 价值方案 | P30 五个内容框 |
| 用户原点 | P28 | 品类切口 | P29 |

参数值支持 `\n` 换行（自动拆成多个段落，继承原段落格式）。
如模板红框有变动，修改 `prepare_template.py` 中的 `TARGETS` 后重新运行即可。

## 约定与注意

- 占位符格式：`{{参数名}}`，可以出现在正文、标题、表格单元格、组合形状内。
- 未提供的参数会保留 `{{参数名}}` 原样并在服务端日志告警，不会阻断生成。
- 占位符在模板编辑时尽量一次性输入完整；若被 PowerPoint 拆成多个样式片段
  （run），服务也能合并替换，但被波及片段的格式会统一为第一个片段的格式。
- 每次转换在独立子进程中完成，PowerPoint/LibreOffice 异常不会拖垮服务。
