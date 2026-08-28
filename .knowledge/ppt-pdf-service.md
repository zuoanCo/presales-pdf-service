# PPT 模板转 PDF 服务

## PPT 转 PDF 的保真度只有两条可靠路线

- **类别**：决策
- **日期**：2026-08-28
- **背景**：需要一个服务把填好参数的 PPT 模板导出为"完全还原排版"的 PDF，且可能部署在任意服务器。
- **内容**：自己用代码渲染 PPT 版式（如 python-pptx + reportlab）不可行，永远追不上真实渲染器。可靠路线只有：① Windows + Office 时用 PowerPoint COM 自动化导出（`Presentation.SaveAs(path, 32)`，32=ppSaveAsPDF），保真度 100%，因为就是 PowerPoint 自己在渲染；② 任意平台用 LibreOffice 无头模式（`soffice --headless --convert-to pdf`），保真度接近 100%。架构上做成可插拔后端，按环境自动探测选择。
- **如何复用**：任何"PPT/Word/Excel 转 PDF 且要求保真"的需求，直接走原生应用自动化或 LibreOffice，不要尝试自行渲染。

## LibreOffice 转换保真度的关键是服务器字体

- **类别**：经验
- **日期**：2026-08-28
- **背景**：LibreOffice 后端转换中文 PPT 模板。
- **内容**：LibreOffice 遇到服务器上没装的字体时会静默替换成其他字体，导致行宽、换行、位置整体偏移——这是"排版不一致"的最常见原因，不是转换器 bug。Docker 镜像里装 `fonts-noto-cjk` 可覆盖一般中文需求；模板若用微软雅黑等商业字体，必须自行拷贝字体文件到 `/usr/share/fonts/` 并 `fc-cache -f`。
- **如何复用**：交付基于 LibreOffice 的文档转换服务时，把"模板字体清单 → 服务器字体安装"作为部署 checklist 的必查项。

## python-pptx 替换占位符要处理 run 拆分

- **类别**：经验
- **日期**：2026-08-28
- **背景**：实现 `{{姓名}}` 这类占位符填充。
- **内容**：PowerPoint 会把一段文本拆成多个 run（样式片段），占位符可能被从中间拆开（比如 `{{姓` 和 `名}}`），逐 run 替换会漏掉。稳妥做法：先逐 run 尝试（最常见情况，格式零损失）；失败则在段落级全文替换，结果写回第一个含占位符的 run、清空其余 run（代价是被波及 run 的格式统一为第一个 run）。另外要递归处理组合形状（group，shape_type==6）和表格单元格。
- **如何复用**：凡是基于 python-pptx 做模板填充/批量改文案，都用这个两级策略；同时提醒模板制作者占位符尽量一次性输入完整。

## python-pptx 访问 shape.line 会"无中生有"创建边框

- **类别**：经验
- **日期**：2026-08-28
- **背景**：给售前模板做占位符化时，用 `shape.line.color` 扫描红框标记后保存文件，生成的 PPT 里多个文本框莫名出现黑色边框。
- **内容**：python-pptx 的 `shape.line` 属性内部调用 `get_or_add_ln()`——只要**读取**它，就会在 XML 里物化一个空的 `<a:ln/>` 元素；空边框没有显式 noFill，按主题默认线条样式渲染成黑色实线框。即"只读操作修改了文档"。凡是扫描后要保存文件的场景，判断边框/填充属性必须直接读 XML：`shape._element.find(qn('p:spPr'))` → 找 `a:ln` → 找 `a:srgbClr`，不经过 python-pptx 的属性访问器。同类坑还包括 `shape.fill`、`shape.shadow` 等惰性创建属性。
- **如何复用**：用 python-pptx 批量检查形状样式且随后要 save 时，一律走 XML 只读路径；或检查完对比 XML diff 确认没有意外物化元素。（已验证的事实）

## 装了 WPS 的机器上"PowerPoint 自动化"实际是 WPS 在响应

- **类别**：知识
- **日期**：2026-08-28
- **背景**：PPT 转 PDF 服务用 `DispatchEx("PowerPoint.Application")` 调 COM 转换，运行中突然持续报 -2147417846「消息筛选器显示应用程序正在使用中」。
- **内容**：用户机器同时装了 MS Office 和 WPS Office，WPS 会接管 `PowerPoint.Application` 的 COM 注册（注册表里留下 `.ksobak` 备份值是 WPS 改过的标志），COM 调用实际由 WPS 演示（wpp.exe）响应——连 `Version` 都伪装成 12.0。WPS 是单实例模型：一旦有一个 wpp.exe 实例卡死（比如自动化进程被强杀后留下的僵尸实例），后续所有 COM 调用都返回 RETRYLATER。解法：`Get-Process wpp` 找到残留实例杀掉即可恢复，用户的交互窗口在 wps.exe 里，不受影响。另外 COM 转换放在独立子进程里跑比在主进程内联调用更抗这种环境故障。
- **如何复用**：在 Windows 上做 Office COM 自动化，先确认 ProgID 背后是 MS Office 还是 WPS；遇到 RETRYLATER 先查 wpp/POWERPNT 僵尸进程，不要盲目重试超过 1 分钟。（已验证的事实）

## 固定模板 + 无 Office 环境：预渲染背景图 + 文字叠加是最务实的路线

- **类别**：决策
- **日期**：2026-08-28
- **背景**：用户要求 PHP 版 PPT→PDF 服务，部署到 Linux，明确拒绝 Office/WPS/LibreOffice 这类重依赖。
- **内容**：PHP 生态没有能渲染 PPT 的库（PhpPresentation 只读写不渲染，TCPDF/mPDF 只生成 PDF），云端转换 API 又有费用和保密问题（售前方案涉密）。既然模板固定、只有 16 个文本参数变，就把渲染拆成两个阶段：构建期用 Office/WPS 一次性把模板渲染成每页背景图并提取参数框坐标样式（存 JSON）；运行期 PHP 只做"背景图 + 按坐标叠文字 + TCPDF 拼 PDF"，服务器零重依赖。版式保真度 100%（背景图），文字保真度取决于打包的字体。要点：① TCPDF 的 MultiCell 第三个参数是每行高度、maxh 会裁剪溢出文字——顶对齐文本框不要设 maxh（PowerPoint 默认就是溢出而非裁剪）；② 可分发的中文字体选 OFL 许可的（Noto Sans SC），微软雅黑不能分发；③ Noto Sans SC 官方只发可变字体，用 fontTools 的 varLib.instancer 在构建期实例化出静态 Regular/Bold TTF（TCPDF 不认可变字体和 CFF OTF）。
- **如何复用**：凡是"固定模板 + 少量动态文字 + 轻量部署"的文档生成需求，都用这个预渲染叠加模式；模板变更时重跑一次预处理即可。

## PowerShell 调中文 JSON 接口要显式转 UTF-8 字节

- **类别**：经验
- **日期**：2026-08-28
- **背景**：在 Windows 上测试 FastAPI 接口，请求体含中文参数。
- **内容**：`Invoke-WebRequest -Body` 直接传含中文的 JSON 字符串会按错误编码发送，服务端收到乱码。要先 `[System.Text.Encoding]::UTF8.GetBytes($json)` 转成字节数组再作为 Body 传入，并带 `Content-Type: application/json; charset=utf-8`。
- **如何复用**：在 PowerShell 里测试任何含中文的 HTTP 接口都这么做；或者干脆写个小 Python 脚本调接口，避开编码坑。
