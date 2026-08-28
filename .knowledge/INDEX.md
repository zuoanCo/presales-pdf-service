# 知识库索引

## PPT/PDF 服务
- [PPT 转 PDF 的保真度只有两条可靠路线](ppt-pdf-service.md#ppt-转-pdf-的保真度只有两条可靠路线) — PowerPoint COM（100%）或 LibreOffice（接近100%），不要自行渲染
- [LibreOffice 转换保真度的关键是服务器字体](ppt-pdf-service.md#libreoffice-转换保真度的关键是服务器字体) — 字体缺失导致静默替换和排版偏移
- [python-pptx 替换占位符要处理 run 拆分](ppt-pdf-service.md#python-pptx-替换占位符要处理-run-拆分) — 先逐 run、后段落级合并的两级策略
- [python-pptx 访问 shape.line 会"无中生有"创建边框](ppt-pdf-service.md#python-pptx-访问-shapeline-会无中生有创建边框) — 只读属性也会改文档，扫描样式必须走 XML
- [装了 WPS 的机器上"PowerPoint 自动化"实际是 WPS 在响应](ppt-pdf-service.md#装了-wps-的机器上powerpoint-自动化实际是-wps-在响应) — RETRYLATER 报错先杀 wpp 僵尸进程
- [固定模板 + 无 Office 环境：预渲染背景图 + 文字叠加是最务实的路线](ppt-pdf-service.md#固定模板--无-office-环境预渲染背景图--文字叠加是最务实的路线) — PHP/Linux 轻量部署的 PPT→PDF 方案
- [PowerShell 调中文 JSON 接口要显式转 UTF-8 字节](ppt-pdf-service.md#powershell-调中文-json-接口要显式转-utf-8-字节) — Invoke-WebRequest 中文乱码解法
