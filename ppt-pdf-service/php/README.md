# 售前诊断建议书 PDF 生成（PHP 轻量版）

纯 PHP 实现，**运行时不需要 Office / WPS / LibreOffice**，可部署到任意 Linux 服务器。

## 原理

模板是固定的 31 页，预处理时已把每页渲染成背景图（`assets/slides/`），
参数框的坐标、字号、颜色、行距提取到 `assets/config.json`。
PHP 只需把参数文字按配置叠加到背景图上，用 TCPDF 合成 PDF——
服务器上除了 PHP 什么都不用装。

## 目录

```
php/
├── api.php              HTTP 接口（生产用这个）
├── generate.php         命令行生成（本地/脚本用）
├── params.example.json  参数示例（16 个参数齐全）
├── src/pdfgen.php       核心逻辑
├── tcpdf/               TCPDF 库（已随包自带，无需 composer）
└── assets/
    ├── config.json      参数框坐标与样式
    ├── slides/          31 页背景图（1920x1080）
    └── fonts/           Noto Sans SC Regular/Bold（OFL 免费商用）
```

## 部署（Linux）

```bash
# 只需要 PHP 7.4+（8.x 推荐）
# 必需的 PHP 扩展：mbstring、curl、gd
# （TCPDF 内部会引用 curl 常量，没装 php-curl 会直接 Fatal error；
#   Debian/Ubuntu: apt install php-mbstring php-curl php-gd）
# 把整个 php/ 目录拷到服务器，然后二选一：

# 1) 挂到现有 nginx/apache + PHP-FPM：把 api.php 放进站点目录即可
# 2) 快速验证：
php -S 0.0.0.0:8080 api.php
```

**注意**：确保 `tcpdf/fonts/` 目录对 PHP 进程可写（首次运行会生成字体缓存）。

## 使用

### HTTP 接口

**Swagger 在线文档**：`http://服务器:8080/docs/`（可查看参数说明、直接在线试调并下载 PDF，
Swagger UI 已打包在 `docs/` 目录，无需外网）。OpenAPI 规范在 `/openapi.json`。

```bash
curl -X POST http://服务器:8080/ \
  -H "Content-Type: application/json" \
  -d @params.example.json \
  -o result.pdf
```

请求体：`{"params": {"产品名称": "...", ...}, "filename": "可选文件名"}`
- `GET /` 返回服务说明和 16 个参数名列表（可用于健康检查）；
- 未填的参数对应区域留空，响应头 `X-Missing-Params` 会列出；
- 参数值用 `\n` 换行。

### 命令行

```bash
php generate.php params.example.json output/售前诊断建议书.pdf
```

## 16 个参数

核心诊断、战法、市场阶段、用户心智状态、企业资源要求、核心战略目标、
行业趋势、用户原点、品类切口、战略打法、
产品名称、原点用户、价值载体、价值心智、视觉锤语言钉、价值方案

## 保真度说明

- **版式**：背景图由 PowerPoint/WPS 渲染（一次性预处理），像素级还原；
- **参数文字**：PHP 叠加，字体用 Noto Sans SC（模板原为微软雅黑/阿里普惠体，
  因授权不能随意分发，视觉风格接近）。行内混排（如加粗标签）会统一字重；
- 如需像素级一致的文字，把模板原字体的 TTF 放入 `assets/fonts/` 并修改
  `src/pdfgen.php` 中的字体路径（授权自行评估）。

## 模板变更怎么办

参数区位置或模板内容变了，需要在有 Office/WPS 的机器上重新跑一次预处理
（ppt-pdf-service 下的 `prep_php.py` + `pdf_convert.py` + `rasterize.py`），
重新生成 `assets/config.json` 和 `assets/slides/`。
