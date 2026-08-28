<?php
/**
 * HTTP 接口（部署到 Linux 用）：
 *   开发测试: php -S 0.0.0.0:8080 api.php
 *   生产:     挂到 nginx/apache 的 PHP-FPM 下即可
 *
 * 请求：POST /  Content-Type: application/json
 *   {
 *     "params": { "产品名称": "亲密陪伴机器人", "核心诊断": "...", ... },
 *     "filename": "售前诊断建议书"   // 可选，下载文件名（不含 .pdf）
 *   }
 *
 * 响应：application/pdf 二进制流；参数错误返回 4xx JSON。
 */

require __DIR__ . '/src/pdfgen.php';

header('X-Content-Type-Options: nosniff');

$method = $_SERVER['REQUEST_METHOD'] ?? 'GET';
$uri    = parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH);

// Swagger 文档页和 OpenAPI 规范按静态文件处理
// （PHP 内置服务器：return false 回退到静态文件；nginx/apache 天然支持）
if ($method === 'GET' && $uri !== '/') {
    return false;
}

if ($method === 'GET') {
    // 健康检查/说明
    header('Content-Type: application/json; charset=utf-8');
    $config = json_decode(file_get_contents(__DIR__ . '/assets/config.json'), true);
    echo json_encode([
        'service' => '售前诊断建议书 PDF 生成',
        'usage'   => 'POST JSON: {"params": {"参数名": "值"}, "filename": "可选"}',
        'docs'    => '/docs/ （Swagger UI，可在线试调）',
        'params'  => array_column($config['params'], 'name'),
    ], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    exit;
}

$body = json_decode(file_get_contents('php://input'), true);
if (!is_array($body) || !isset($body['params']) || !is_array($body['params'])) {
    http_response_code(400);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['error' => '请求体必须是 JSON，且包含 params 对象'], JSON_UNESCAPED_UNICODE);
    exit;
}

$filename = preg_replace('/[\\\\\/\:\*\?\"\<\>\|]/u', '', $body['filename'] ?? '售前诊断建议书');
$outFile  = sys_get_temp_dir() . '/presales_' . uniqid() . '.pdf';

try {
    $missing = build_pdf($body['params'], $outFile);
} catch (Throwable $e) {
    http_response_code(500);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode(['error' => $e->getMessage()], JSON_UNESCAPED_UNICODE);
    exit;
}

header('Content-Type: application/pdf');
header('Content-Disposition: attachment; filename*=UTF-8\'\'' . rawurlencode($filename) . '.pdf');
if ($missing) {
    header('X-Missing-Params: ' . implode(',', $missing));
}
header('Content-Length: ' . filesize($outFile));
readfile($outFile);
@unlink($outFile);
