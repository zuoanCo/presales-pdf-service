<?php
/**
 * 售前诊断建议书 PDF 生成核心（纯 PHP，无 Office/WPS 依赖）
 *
 * 原理：模板 31 页已预先渲染成背景图（assets/slides/），
 * 本文件按 assets/config.json 里记录的坐标和样式，把参数文字叠加到对应位置。
 *
 * 依赖：TCPDF（已随包放在 tcpdf/ 目录），PHP 7.4+ 即可，无需任何扩展之外的组件。
 */

require_once __DIR__ . '/../tcpdf/tcpdf.php';

/** 个别参数的样式覆盖（模板里是主题色/特殊字重，提取时拿不到准确值） */
const PARAM_OVERRIDES = [
    '核心诊断' => ['color' => 'FFFFFF', 'bold' => false], // 深蓝底上的白色大字
];

/**
 * 生成 PDF。
 *
 * @param array  $params  参数键值对（key 为参数名，value 支持 \n 换行）
 * @param string $outFile 输出 PDF 路径
 * @return string[] 未填充（留空）的参数名
 */
function build_pdf(array $params, string $outFile): array
{
    $assets = __DIR__ . '/../assets';
    $config = json_decode(file_get_contents($assets . '/config.json'), true);

    $pageW = $config['page_w_pt'];
    $pageH = $config['page_h_pt'];

    $pdf = new TCPDF('L', 'pt', [$pageW, $pageH], true, 'UTF-8', false);
    $pdf->SetCreator('售前PDF生成服务');
    $pdf->SetTitle('售前诊断建议书');
    $pdf->setPrintHeader(false);
    $pdf->setPrintFooter(false);
    $pdf->SetMargins(0, 0, 0);
    $pdf->SetAutoPageBreak(false, 0);
    $pdf->setCellMargins(0, 0, 0, 0);
    $pdf->setCellPaddings(0, 0, 0, 0);

    // 字体在 addTTFfont 时会生成缓存文件到 tcpdf/fonts/，需保证该目录可写
    $fontRegular = TCPDF_FONTS::addTTFfont($assets . '/fonts/NotoSansSC-Regular.ttf', 'TrueTypeUnicode', '', 96);
    $fontBold    = TCPDF_FONTS::addTTFfont($assets . '/fonts/NotoSansSC-Bold.ttf', 'TrueTypeUnicode', '', 96);

    // 参数按页分组
    $byPage = [];
    foreach ($config['params'] as $p) {
        $byPage[$p['page']][] = $p;
    }

    $missing = [];

    for ($page = 1; $page <= $config['pages']; $page++) {
        $pdf->AddPage();
        // 整页背景图（模板原版式，像素级还原）
        $pdf->Image(
            $assets . '/slides/slide-' . $page . '.jpg',
            0, 0, $pageW, $pageH, 'JPG', '', '', false, 300, '', false, false, 0
        );

        foreach ($byPage[$page] ?? [] as $p) {
            $name = $p['name'];
            $val = isset($params[$name]) ? trim((string)$params[$name]) : '';
            if ($val === '') {
                $missing[] = $name;
                continue;
            }
            if (!empty($p['suffix'])) {
                $val .= "\n" . $p['suffix'];
            }

            // 样式：先取提取值，再应用覆盖
            $style = array_merge($p, PARAM_OVERRIDES[$name] ?? []);
            $bold  = !empty($style['bold']);
            $color = $style['color'] ?? '000000';
            $size  = $style['size_pt'];

            // 对齐 / 垂直锚点映射（OOXML -> TCPDF）
            $align  = ['l' => 'L', 'ctr' => 'C', 'r' => 'R', 'just' => 'J'][$style['align']] ?? 'L';
            $valign = ['t' => 'T', 'ctr' => 'M', 'b' => 'B'][$style['anchor']] ?? 'T';

            // 行高：PPT 单倍行距≈1.2 倍字号，spcPct 在此基础上再乘
            $spacing = is_numeric($style['line_spacing']) ? (float)$style['line_spacing'] : 1.0;
            $lineH   = $size * 1.2 * $spacing;

            // 文字区 = 形状框 - 内边距
            $ins = $style['insets_pt'];
            $x = $style['x_pt'] + $ins['l'];
            $y = $style['y_pt'] + $ins['t'];
            $w = $style['w_pt'] - $ins['l'] - $ins['r'];
            $h = $style['h_pt'] - $ins['t'] - $ins['b'];

            $pdf->SetFont($bold ? $fontBold : $fontRegular, '', $size);
            $pdf->SetTextColor(
                hexdec(substr($color, 0, 2)),
                hexdec(substr($color, 2, 2)),
                hexdec(substr($color, 4, 2))
            );
            // MultiCell 的 $h 是"每行高度"。maxh 只在需要垂直居中时启用（用于计算居中偏移）；
            // 顶部对齐的框不限制高度——与 PowerPoint 默认行为一致（文字超出框时自然溢出而非裁剪）
            $maxh     = ($valign === 'M') ? $h : 0;
            $useAlign = ($valign === 'M') ? 'M' : 'T';
            $pdf->MultiCell($w, $lineH, $val, 0, $align, false, 2, $x, $y, true, 0, false, true, $maxh, $useAlign, false);
        }
    }

    $pdf->Output($outFile, 'F');
    return array_values(array_unique($missing));
}
