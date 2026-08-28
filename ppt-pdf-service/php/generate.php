<?php
/**
 * 命令行用法：
 *   php generate.php [参数.json] [输出.pdf]
 *
 * 参数文件格式见 params.example.json；不传参数时用 params.example.json。
 */

require __DIR__ . '/src/pdfgen.php';

$paramsFile = $argv[1] ?? __DIR__ . '/params.example.json';
$outFile    = $argv[2] ?? __DIR__ . '/output/售前诊断建议书.pdf';

// TCPDF 写文件要求绝对路径
if (!preg_match('/^([A-Za-z]:[\\\\\/]|\/|\\\\\\\\)/', $outFile)) {
    $outFile = getcwd() . DIRECTORY_SEPARATOR . $outFile;
}

if (!file_exists($paramsFile)) {
    fwrite(STDERR, "参数文件不存在: $paramsFile\n");
    exit(1);
}
$params = json_decode(file_get_contents($paramsFile), true);
if (!is_array($params)) {
    fwrite(STDERR, "参数文件不是合法 JSON\n");
    exit(1);
}

@mkdir(dirname($outFile), 0777, true);
$missing = build_pdf($params, $outFile);

if ($missing) {
    echo "[提示] 以下参数未填，对应区域留空: " . implode(', ', $missing) . "\n";
}
echo "[完成] $outFile\n";
