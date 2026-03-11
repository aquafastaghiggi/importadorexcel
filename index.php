<?php
header('Content-Type: text/html; charset=utf-8');

$resultado = null;
$erro = null;
$jsonPath = null;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (!isset($_FILES['arquivo']) || $_FILES['arquivo']['error'] !== UPLOAD_ERR_OK) {
        $erro = 'Falha no upload do arquivo.';
    } else {
        $ext = strtolower(pathinfo($_FILES['arquivo']['name'], PATHINFO_EXTENSION));

        if (!in_array($ext, ['xlsx', 'xlsm', 'xltx', 'xltm'])) {
            $erro = 'Envie um arquivo Excel válido (.xlsx, .xlsm, .xltx, .xltm).';
        } else {
            $uploadDir = __DIR__ . DIRECTORY_SEPARATOR . 'uploads';
            $resultDir = __DIR__ . DIRECTORY_SEPARATOR . 'resultados';

            if (!is_dir($uploadDir)) {
                mkdir($uploadDir, 0777, true);
            }

            if (!is_dir($resultDir)) {
                mkdir($resultDir, 0777, true);
            }

            $nomeSeguro = date('Ymd_His') . '_' . preg_replace('/[^a-zA-Z0-9._-]/', '_', $_FILES['arquivo']['name']);
            $arquivoDestino = $uploadDir . DIRECTORY_SEPARATOR . $nomeSeguro;

            if (!move_uploaded_file($_FILES['arquivo']['tmp_name'], $arquivoDestino)) {
                $erro = 'Não foi possível salvar o arquivo enviado.';
            } else {
                $python = 'python';
                $script = __DIR__ . DIRECTORY_SEPARATOR . 'python' . DIRECTORY_SEPARATOR . 'process_excel.py';

                $comando = $python . ' ' . escapeshellarg($script) . ' ' . escapeshellarg($arquivoDestino);
                $saida = shell_exec($comando);

                if ($saida === null) {
                    $erro = 'Não foi possível executar o script Python.';
                } else {
                    $dados = json_decode(trim($saida), true);

                    if (json_last_error() !== JSON_ERROR_NONE) {
                        $erro = 'Erro ao interpretar JSON retornado pelo Python: ' . json_last_error_msg()
                              . '<br><pre>' . htmlspecialchars($saida, ENT_QUOTES, 'UTF-8') . '</pre>';
                    } elseif (empty($dados['success'])) {
                        $erro = htmlspecialchars($dados['error'] ?? 'Erro desconhecido no processamento.', ENT_QUOTES, 'UTF-8');
                    } else {
                        $resultado = $dados;

                        $jsonName = pathinfo($nomeSeguro, PATHINFO_FILENAME) . '.json';
                        $jsonPath = $resultDir . DIRECTORY_SEPARATOR . $jsonName;

                        file_put_contents(
                            $jsonPath,
                            json_encode($resultado, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE)
                        );
                    }
                }
            }
        }
    }
}

function e($valor): string
{
    return htmlspecialchars((string)$valor, ENT_QUOTES, 'UTF-8');
}

function get_hidden_columns(): array
{
    return [
        'start_row',
        'end_row',
        'total_rows_raw',
        'region_start_col',
        'region_end_col',
        'region_width',
        'row_excel',
        'linha_ordem',
    ];
}

function render_assoc_table(array $rows, string $titulo): void
{
    if (empty($rows)) {
        return;
    }

    $hiddenColumns = get_hidden_columns();

    $allKeys = [];
    foreach ($rows as $row) {
        if (!is_array($row)) {
            continue;
        }

        foreach ($row as $key => $value) {
            if (in_array($key, $hiddenColumns, true)) {
                continue;
            }

            if (!in_array($key, $allKeys, true)) {
                $allKeys[] = $key;
            }
        }
    }

    if (empty($allKeys)) {
        return;
    }

    echo '<div class="bloco-card">';
    echo '<h3>' . e($titulo) . '</h3>';
    echo '<div class="table-wrap">';
    echo '<table>';
    echo '<thead><tr>';

    foreach ($allKeys as $key) {
        echo '<th>' . e($key) . '</th>';
    }

    echo '</tr></thead>';
    echo '<tbody>';

    foreach ($rows as $row) {
        echo '<tr>';

        foreach ($allKeys as $key) {
            $value = $row[$key] ?? '';

            if (is_array($value)) {
                $value = json_encode($value, JSON_UNESCAPED_UNICODE);
            }

            echo '<td>' . e((string)$value) . '</td>';
        }

        echo '</tr>';
    }

    echo '</tbody></table>';
    echo '</div>';
    echo '</div>';
}

function render_key_value_table(array $data, string $titulo): void
{
    if (empty($data)) {
        return;
    }

    $hiddenColumns = get_hidden_columns();

    echo '<div class="bloco-card">';
    echo '<h3>' . e($titulo) . '</h3>';
    echo '<div class="table-wrap">';
    echo '<table>';
    echo '<thead><tr><th>Campo</th><th>Valor</th></tr></thead>';
    echo '<tbody>';

    foreach ($data as $key => $value) {
        if (in_array($key, $hiddenColumns, true)) {
            continue;
        }

        if (is_array($value)) {
            $value = json_encode($value, JSON_UNESCAPED_UNICODE);
        }

        echo '<tr>';
        echo '<td>' . e((string)$key) . '</td>';
        echo '<td>' . e((string)$value) . '</td>';
        echo '</tr>';
    }

    echo '</tbody></table>';
    echo '</div>';
    echo '</div>';
}

function render_block_summary(array $rawBlocks): void
{
    if (empty($rawBlocks)) {
        return;
    }

    $rows = [];
    foreach ($rawBlocks as $block) {
        $rows[] = [
            'block' => $block['block'] ?? '',
            'label_found' => $block['label_found'] ?? '',
        ];
    }

    render_assoc_table($rows, 'Blocos detectados');
}

function render_ignored_regions(array $ignoredRegions): void
{
    if (empty($ignoredRegions)) {
        return;
    }

    $rows = [];
    foreach ($ignoredRegions as $region) {
        $rows[] = [
            'motivo' => $region['motivo'] ?? '',
        ];
    }

    render_assoc_table($rows, 'Regiões ignoradas');
}

function block_meta(string $blockKey): array
{
    $map = [
        'header' => ['label' => 'Cabeçalho principal', 'icon' => '📌'],
        'plano_negocios_rows' => ['label' => 'Plano de negócios', 'icon' => '📋'],
        'historico_rows' => ['label' => 'Histórico', 'icon' => '📈'],
        'objetivos_rows' => ['label' => 'Objetivos', 'icon' => '🎯'],
        'descricao_investimento_rows' => ['label' => 'Descrição do investimento', 'icon' => '💰'],
        'contrapartidas_rows' => ['label' => 'Contrapartidas', 'icon' => '🧩'],
        'contrapartidas_itens_foco_rows' => ['label' => 'Contrapartidas - Itens foco', 'icon' => '📦'],
        'encartes_obrigatorios_rows' => ['label' => 'Encartes obrigatórios / Sugestão de encartes', 'icon' => '🗓️'],
        'cadastros_vinculados_rows' => ['label' => 'Cadastros vinculados', 'icon' => '🧾'],
        'investimentos_extras_rows' => ['label' => 'Investimentos extras', 'icon' => '➕'],
        'raw_blocks' => ['label' => 'Blocos detectados', 'icon' => '🔎'],
        'resumo_analitico' => ['label' => 'Resumo analítico', 'icon' => '📊'],
    ];

    return $map[$blockKey] ?? ['label' => $blockKey, 'icon' => '•'];
}

function has_block_content(array $form, string $blockKey): bool
{
    if ($blockKey === 'resumo_analitico') {
        return true;
    }

    if (!array_key_exists($blockKey, $form)) {
        return false;
    }

    if ($blockKey === 'header') {
        foreach (($form['header'] ?? []) as $value) {
            if ($value !== null && $value !== '') {
                return true;
            }
        }
        return false;
    }

    return !empty($form[$blockKey]);
}

function get_available_blocks(array $form): array
{
    $ordered = [
        'header',
        'plano_negocios_rows',
        'historico_rows',
        'objetivos_rows',
        'descricao_investimento_rows',
        'contrapartidas_rows',
        'contrapartidas_itens_foco_rows',
        'encartes_obrigatorios_rows',
        'cadastros_vinculados_rows',
        'investimentos_extras_rows',
        'resumo_analitico',
        'raw_blocks',
    ];

    $available = [];
    foreach ($ordered as $key) {
        if (has_block_content($form, $key)) {
            $available[] = $key;
        }
    }

    return $available;
}

function normalize_cmp(string $text): string
{
    $text = trim($text);
    if ($text === '') {
        return '';
    }

    $map = [
        'Á'=>'A','À'=>'A','Â'=>'A','Ã'=>'A','Ä'=>'A',
        'É'=>'E','È'=>'E','Ê'=>'E','Ë'=>'E',
        'Í'=>'I','Ì'=>'I','Î'=>'I','Ï'=>'I',
        'Ó'=>'O','Ò'=>'O','Ô'=>'O','Õ'=>'O','Ö'=>'O',
        'Ú'=>'U','Ù'=>'U','Û'=>'U','Ü'=>'U',
        'Ç'=>'C',
        'á'=>'A','à'=>'A','â'=>'A','ã'=>'A','ä'=>'A',
        'é'=>'E','è'=>'E','ê'=>'E','ë'=>'E',
        'í'=>'I','ì'=>'I','î'=>'I','ï'=>'I',
        'ó'=>'O','ò'=>'O','ô'=>'O','õ'=>'O','ö'=>'O',
        'ú'=>'U','ù'=>'U','û'=>'U','ü'=>'U',
        'ç'=>'C'
    ];

    $text = strtr($text, $map);
    $text = strtoupper($text);
    $text = preg_replace('/\s+/', ' ', $text);
    return trim($text);
}

function contains_all(string $text, array $terms): bool
{
    foreach ($terms as $term) {
        if (strpos($text, normalize_cmp($term)) === false) {
            return false;
        }
    }
    return true;
}

function contains_any(string $text, array $terms): bool
{
    foreach ($terms as $term) {
        if (strpos($text, normalize_cmp($term)) !== false) {
            return true;
        }
    }
    return false;
}

function first_non_empty_year(array $rows): ?int
{
    foreach ($rows as $row) {
        if (isset($row['ano_bloco']) && $row['ano_bloco'] !== null && $row['ano_bloco'] !== '') {
            return (int)$row['ano_bloco'];
        }
    }
    return null;
}

function format_number_br($value, int $decimals = 2): string
{
    if ($value === null || $value === '') {
        return '';
    }

    if (!is_numeric($value)) {
        return (string)$value;
    }

    return number_format((float)$value, $decimals, ',', '.');
}

function build_itens_foco_summary(array $itensFocoRows): array
{
    $produtos = [];
    $volumeMensalTotal = 0.0;
    $volumeMensalTemValor = false;
    $volumePeriodoTotal = 0.0;
    $volumePeriodoTemValor = false;
    $unidades = [];

    foreach ($itensFocoRows as $row) {
        $produto = $row['produto'] ?? '';
        if ($produto !== '') {
            $produtos[] = (string)$produto;
        }

        if (isset($row['volume_mensal_numerica']) && is_numeric($row['volume_mensal_numerica'])) {
            $volumeMensalTotal += (float)$row['volume_mensal_numerica'];
            $volumeMensalTemValor = true;
        } elseif (isset($row['volume_mensal']) && is_numeric(str_replace(',', '.', (string)$row['volume_mensal']))) {
            $volumeMensalTotal += (float)str_replace(',', '.', (string)$row['volume_mensal']);
            $volumeMensalTemValor = true;
        }

        if (isset($row['volume_periodo_numerica']) && is_numeric($row['volume_periodo_numerica'])) {
            $volumePeriodoTotal += (float)$row['volume_periodo_numerica'];
            $volumePeriodoTemValor = true;
        } elseif (isset($row['volume_periodo']) && is_numeric(str_replace(',', '.', (string)$row['volume_periodo']))) {
            $volumePeriodoTotal += (float)str_replace(',', '.', (string)$row['volume_periodo']);
            $volumePeriodoTemValor = true;
        }

        if (!empty($row['unidade'])) {
            $unidades[] = (string)$row['unidade'];
        }
    }

    $produtos = array_values(array_unique(array_filter($produtos)));
    $unidades = array_values(array_unique(array_filter($unidades)));

    return [
        'total_produtos_itens_foco' => count($produtos),
        'produtos_itens_foco' => implode(' | ', $produtos),
        'volume_mensal_total_itens_foco' => $volumeMensalTemValor ? $volumeMensalTotal : null,
        'volume_periodo_total_itens_foco' => $volumePeriodoTemValor ? $volumePeriodoTotal : null,
        'unidades_itens_foco' => implode(' | ', $unidades),
    ];
}

function build_client_summary(array $form): array
{
    $header = $form['header'] ?? [];
    $historico = $form['historico_rows'] ?? [];
    $objetivos = $form['objetivos_rows'] ?? [];
    $descricaoInvestimento = $form['descricao_investimento_rows'] ?? [];
    $investimentosExtras = $form['investimentos_extras_rows'] ?? [];
    $itensFocoRows = $form['contrapartidas_itens_foco_rows'] ?? [];

    $anoBase = first_non_empty_year($historico) ?? first_non_empty_year($objetivos);
    $anoSeguinte = $anoBase ? $anoBase + 1 : null;

    $summary = [
        'cliente' => $header['cliente'] ?? '',
        'periodo_plano' => $header['periodo_original'] ?? '',
        'numero_acordo' => $header['numero_acordo'] ?? '',
        'ano_base_historico' => $anoBase,
        'ano_objetivo_seguinte' => $anoSeguinte,
        'valor_total_periodo' => null,
        'valor_total_periodo_origem' => '',
        'caixas_total_periodo' => null,
        'caixas_total_periodo_origem' => '',
        'investimento_periodo' => null,
        'investimento_periodo_origem' => '',
        'objetivo_ano_seguinte_valor' => null,
        'objetivo_ano_seguinte_valor_origem' => '',
        'objetivo_ano_seguinte_caixas' => null,
        'objetivo_ano_seguinte_caixas_origem' => '',
    ];

    foreach ($historico as $row) {
        $desc = normalize_cmp((string)($row['descricao'] ?? ''));
        $valor = $row['valor_numerico'] ?? null;

        if ($valor === null || $desc === '') {
            continue;
        }

        if ($summary['valor_total_periodo'] === null) {
            $isValorHistorico =
                contains_any($desc, ['COMPRA']) &&
                contains_any($desc, ['IMP', 'IMPOSTO', 'IMPOSTOS', 'VALOR']);

            if ($isValorHistorico && !contains_any($desc, ['OBJETIVO'])) {
                $summary['valor_total_periodo'] = $valor;
                $summary['valor_total_periodo_origem'] = (string)($row['descricao'] ?? '');
            }
        }

        if ($summary['caixas_total_periodo'] === null) {
            $isCaixaHistorico =
                contains_any($desc, ['VOLUME', 'CAIXAS', 'CX']) &&
                !contains_any($desc, ['OBJETIVO']);

            if ($isCaixaHistorico) {
                $summary['caixas_total_periodo'] = $valor;
                $summary['caixas_total_periodo_origem'] = (string)($row['descricao'] ?? '');
            }
        }
    }

    foreach ($descricaoInvestimento as $row) {
        $desc = normalize_cmp((string)($row['descricao'] ?? ''));
        $valor = $row['valor_numerico'] ?? null;

        if ($valor === null || $desc === '') {
            continue;
        }

        if (
            $summary['investimento_periodo'] === null &&
            contains_any($desc, ['INVESTIMENTO']) &&
            (
                contains_any($desc, ['TOTAL']) ||
                contains_any($desc, ['ANUAL']) ||
                contains_any($desc, ['PLANO DE TRADE'])
            )
        ) {
            $summary['investimento_periodo'] = $valor;
            $summary['investimento_periodo_origem'] = (string)($row['descricao'] ?? '');
        }
    }

    if ($summary['investimento_periodo'] === null) {
        foreach ($investimentosExtras as $row) {
            $desc = normalize_cmp((string)($row['descricao'] ?? ''));
            $valor = $row['valor_numerico'] ?? null;

            if ($valor === null || $desc === '') {
                continue;
            }

            if (contains_any($desc, ['RETORNO FINANCEIRO', 'INVESTIMENTO', 'SCANNTECH'])) {
                $summary['investimento_periodo'] = $valor;
                $summary['investimento_periodo_origem'] = (string)($row['descricao'] ?? '');
                break;
            }
        }
    }

    foreach ($objetivos as $row) {
        $descOriginal = (string)($row['descricao'] ?? '');
        $desc = normalize_cmp($descOriginal);
        $valor = $row['valor_numerico'] ?? null;
        $anoLinha = isset($row['ano_bloco']) && $row['ano_bloco'] !== '' ? (int)$row['ano_bloco'] : null;

        if ($valor === null || $desc === '') {
            continue;
        }

        $ehAnoSeguinte = false;
        if ($anoSeguinte !== null) {
            $ehAnoSeguinte = ($anoLinha === $anoSeguinte) || strpos($desc, (string)$anoSeguinte) !== false;
        } else {
            $ehAnoSeguinte = true;
        }

        if (!$ehAnoSeguinte) {
            continue;
        }

        if ($summary['objetivo_ano_seguinte_valor'] === null) {
            $isValorObjetivo =
                contains_any($desc, ['OBJETIVO']) &&
                (
                    contains_any($desc, ['VALOR']) ||
                    contains_all($desc, ['COMPRA', 'IMP']) ||
                    contains_any($desc, ['COMPRA C/ IMP', 'COMPRA C IMP', 'COMPRA COM IMPOSTO'])
                );

            if ($isValorObjetivo) {
                $summary['objetivo_ano_seguinte_valor'] = $valor;
                $summary['objetivo_ano_seguinte_valor_origem'] = $descOriginal;
            }
        }

        if ($summary['objetivo_ano_seguinte_caixas'] === null) {
            $isCaixaObjetivo =
                contains_any($desc, ['OBJETIVO']) &&
                contains_any($desc, ['VOLUME', 'CAIXAS', 'CX']);

            if ($isCaixaObjetivo) {
                $summary['objetivo_ano_seguinte_caixas'] = $valor;
                $summary['objetivo_ano_seguinte_caixas_origem'] = $descOriginal;
            }
        }
    }

    $summary = array_merge(
        $summary,
        build_itens_foco_summary($itensFocoRows)
    );

    return $summary;
}

function build_contrapartidas_detail(array $contrapartidasRows): array
{
    $out = [];

    foreach ($contrapartidasRows as $row) {
        $out[] = [
            'quantidade' => $row['quantidade'] ?? '',
            'quantidade_numerica' => isset($row['quantidade_numerica']) ? format_number_br($row['quantidade_numerica']) : '',
            'tipo_acao' => $row['tipo_acao'] ?? '',
            'periodo' => $row['periodo'] ?? '',
            'lojas' => $row['lojas'] ?? '',
            'obs' => $row['obs'] ?? '',
        ];
    }

    return $out;
}

function render_client_summary(array $form): void
{
    $summary = build_client_summary($form);

    $rows = [[
        'cliente' => $summary['cliente'],
        'periodo_plano' => $summary['periodo_plano'],
        'numero_acordo' => $summary['numero_acordo'],
        'ano_base_historico' => $summary['ano_base_historico'],
        'ano_objetivo_seguinte' => $summary['ano_objetivo_seguinte'],
        'valor_total_periodo' => format_number_br($summary['valor_total_periodo']),
        'valor_total_periodo_origem' => $summary['valor_total_periodo_origem'],
        'caixas_total_periodo' => format_number_br($summary['caixas_total_periodo']),
        'caixas_total_periodo_origem' => $summary['caixas_total_periodo_origem'],
        'investimento_periodo' => format_number_br($summary['investimento_periodo']),
        'investimento_periodo_origem' => $summary['investimento_periodo_origem'],
        'objetivo_ano_seguinte_valor' => format_number_br($summary['objetivo_ano_seguinte_valor']),
        'objetivo_ano_seguinte_valor_origem' => $summary['objetivo_ano_seguinte_valor_origem'],
        'objetivo_ano_seguinte_caixas' => format_number_br($summary['objetivo_ano_seguinte_caixas']),
        'objetivo_ano_seguinte_caixas_origem' => $summary['objetivo_ano_seguinte_caixas_origem'],
    ]];

    render_assoc_table($rows, 'Resumo consolidado do cliente');

    $contrapartidasDetalhe = build_contrapartidas_detail($form['contrapartidas_rows'] ?? []);
    if (!empty($contrapartidasDetalhe)) {
        render_assoc_table($contrapartidasDetalhe, 'Contrapartidas detalhadas (desagrupadas)');
    }

    $rowsItensFoco = [[
        'total_produtos_itens_foco' => $summary['total_produtos_itens_foco'] ?? '',
        'produtos_itens_foco' => $summary['produtos_itens_foco'] ?? '',
        'volume_mensal_total_itens_foco' => format_number_br($summary['volume_mensal_total_itens_foco'] ?? null),
        'volume_periodo_total_itens_foco' => format_number_br($summary['volume_periodo_total_itens_foco'] ?? null),
        'unidades_itens_foco' => $summary['unidades_itens_foco'] ?? '',
    ]];

    render_assoc_table($rowsItensFoco, 'Resumo de contrapartidas - itens foco');

    echo '<div class="bloco-card">';
    echo '<h3>Leitura analítica inicial</h3>';
    echo '<ul class="summary-list">';
    echo '<li><strong>Valor total e investimento no período:</strong> busca principalmente em histórico, descrição do investimento e investimentos extras.</li>';
    echo '<li><strong>Objetivo do ano seguinte:</strong> busca principalmente em objetivos, usando o ano do bloco como referência quando disponível.</li>';
    echo '<li><strong>Caixas:</strong> tenta identificar descrições com volume, caixas ou CX.</li>';
    echo '<li><strong>Contrapartidas:</strong> são apresentadas de forma desagrupada, linha a linha, preservando a quantidade por ação.</li>';
    echo '<li><strong>Encartes obrigatórios / Sugestão de encartes:</strong> agora o parser lê múltiplas grades mensais dentro do mesmo bloco, como janeiro-junho e julho-dezembro.</li>';
    echo '<li><strong>Itens foco:</strong> lista os produtos encontrados e tenta consolidar volume mensal e volume do período.</li>';
    echo '<li><strong>Origem:</strong> ao lado de cada valor principal foi mantida a descrição de origem encontrada pelo parser.</li>';
    echo '</ul>';
    echo '</div>';
}
?>
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Importador de Excel - Navegação por Formulários e Blocos</title>
    <style>
        * {
            box-sizing: border-box;
        }

        body {
            font-family: Arial, sans-serif;
            background: #f4f6f8;
            margin: 0;
            padding: 24px;
            color: #1f2937;
        }

        .container {
            max-width: 1680px;
            margin: 0 auto;
            background: #ffffff;
            border-radius: 14px;
            padding: 24px;
            box-shadow: 0 2px 14px rgba(0, 0, 0, 0.08);
        }

        h1, h2, h3, h4 {
            margin-top: 0;
        }

        .muted {
            color: #5b6470;
            font-size: 14px;
        }

        .erro {
            background: #ffe8e8;
            color: #a10000;
            padding: 14px;
            border-radius: 8px;
            margin-bottom: 20px;
        }

        .ok {
            background: #e9f8ee;
            color: #0b6a2b;
            padding: 14px;
            border-radius: 8px;
            margin-bottom: 20px;
        }

        .summary-card {
            background: #f7f9fc;
            border: 1px solid #dfe7f1;
            border-radius: 10px;
            padding: 16px;
            margin-bottom: 20px;
        }

        .summary-grid {
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
        }

        .summary-item {
            background: #eef3f9;
            border: 1px solid #d7dee8;
            border-radius: 8px;
            padding: 10px 12px;
            min-width: 220px;
        }

        .summary-list {
            margin: 0;
            padding-left: 18px;
        }

        .summary-list li {
            margin-bottom: 8px;
        }

        form {
            margin-bottom: 20px;
        }

        input[type="file"] {
            padding: 8px;
        }

        button {
            padding: 10px 18px;
            cursor: pointer;
            border: none;
            border-radius: 8px;
            background: #0f3b66;
            color: #fff;
            font-weight: bold;
        }

        button:hover {
            background: #0c3256;
        }

        .sheet-card {
            border: 1px solid #d9e3ee;
            border-radius: 12px;
            padding: 18px;
            margin-top: 24px;
            background: #fcfdff;
        }

        .sheet-title {
            margin-bottom: 12px;
        }

        .sheet-info {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 16px;
        }

        .chip {
            display: inline-block;
            padding: 6px 10px;
            background: #e8eef8;
            border-radius: 20px;
            font-size: 13px;
            border: 1px solid #d5ddeb;
        }

        .form-card {
            border: 1px solid #d8e2ed;
            border-radius: 12px;
            padding: 18px;
            margin-top: 20px;
            background: #ffffff;
        }

        .form-title {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            flex-wrap: wrap;
            margin-bottom: 14px;
        }

        .form-badges {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }

        .badge {
            display: inline-block;
            padding: 6px 10px;
            background: #edf4fb;
            border: 1px solid #d7e4f2;
            border-radius: 18px;
            font-size: 12px;
            color: #26415f;
        }

        .sheet-layout {
            display: flex;
            gap: 18px;
            align-items: flex-start;
        }

        .sidebar {
            width: 300px;
            min-width: 300px;
            background: #f7f9fc;
            border: 1px solid #dbe4ef;
            border-radius: 12px;
            padding: 14px;
            position: sticky;
            top: 18px;
        }

        .sidebar h3 {
            font-size: 18px;
            margin-bottom: 12px;
        }

        .nav-blocks {
            display: flex;
            flex-direction: column;
            gap: 8px;
        }

        .nav-link {
            display: block;
            text-decoration: none;
            color: #102a43;
            background: #ffffff;
            border: 1px solid #dbe4ef;
            border-radius: 10px;
            padding: 10px 12px;
            font-size: 14px;
            transition: all 0.2s ease;
        }

        .nav-link:hover {
            background: #eef4fb;
            border-color: #b7c7db;
        }

        .nav-link small {
            display: block;
            color: #5b6470;
            margin-top: 4px;
            font-size: 12px;
        }

        .content-area {
            flex: 1;
            min-width: 0;
        }

        .section-block {
            margin-bottom: 22px;
            scroll-margin-top: 24px;
        }

        .section-header {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 10px;
            padding-bottom: 8px;
            border-bottom: 2px solid #e6edf5;
        }

        .section-header .icon {
            font-size: 20px;
        }

        .bloco-card {
            background: #f7f9fc;
            border: 1px solid #dfe7f1;
            border-radius: 10px;
            padding: 16px;
            margin-bottom: 16px;
        }

        .table-wrap {
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 8px;
            background: #fff;
        }

        table th,
        table td {
            border: 1px solid #d7dee8;
            padding: 8px;
            text-align: left;
            vertical-align: top;
            font-size: 14px;
            white-space: nowrap;
        }

        table th {
            background: #eef3f9;
        }

        .back-top {
            display: inline-block;
            margin-top: 10px;
            text-decoration: none;
            font-size: 13px;
            color: #0f3b66;
        }

        .sheet-anchor,
        .form-anchor {
            scroll-margin-top: 24px;
        }

        @media (max-width: 1150px) {
            .sheet-layout {
                flex-direction: column;
            }

            .sidebar {
                width: 100%;
                min-width: 100%;
                position: static;
            }
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Importador de Excel - Navegação por Formulários e Blocos</h1>
    <p class="muted">
        Esta versão revisa o bloco de encartes para capturar múltiplas grades mensais dentro do mesmo formulário, mantendo o restante da navegação e da apresentação.
    </p>

    <form method="post" enctype="multipart/form-data">
        <input type="file" name="arquivo" accept=".xlsx,.xlsm,.xltx,.xltm" required>
        <button type="submit">Processar arquivo</button>
    </form>

    <?php if ($erro): ?>
        <div class="erro"><?= $erro ?></div>
    <?php endif; ?>

    <?php if ($resultado): ?>
        <div class="ok">
            Arquivo processado com sucesso.
            <?php if ($jsonPath): ?>
                <br>JSON técnico salvo em: <strong><?= e($jsonPath) ?></strong>
            <?php endif; ?>
        </div>

        <div class="summary-card" id="topo">
            <h2>Resumo do processamento</h2>
            <div class="summary-grid">
                <div class="summary-item"><strong>Arquivo:</strong><br><?= e($resultado['file_name'] ?? '') ?></div>
                <div class="summary-item"><strong>Total de abas:</strong><br><?= e($resultado['total_sheets'] ?? 0) ?></div>
                <div class="summary-item"><strong>Data de processamento:</strong><br><?= e($resultado['processed_at'] ?? '') ?></div>
            </div>
        </div>

        <?php if (!empty($resultado['sheets'])): ?>
            <?php foreach ($resultado['sheets'] as $sheetIndex => $sheet): ?>
                <?php $sheetSlug = 'sheet_' . $sheetIndex; ?>
                <div class="sheet-card sheet-anchor" id="<?= e($sheetSlug) ?>">
                    <div class="sheet-title">
                        <h2>Aba: <?= e($sheet['sheet_name'] ?? '') ?></h2>
                    </div>

                    <?php
                    $forms = $sheet['forms'] ?? [];
                    $ignoredRegions = $sheet['ignored_regions'] ?? [];
                    ?>

                    <div class="sheet-info">
                        <span class="chip">formulários detectados: <?= e(count($forms)) ?></span>
                        <?php if (!empty($ignoredRegions)): ?>
                            <span class="chip">regiões ignoradas: <?= e(count($ignoredRegions)) ?></span>
                        <?php endif; ?>
                    </div>

                    <?php if (!empty($ignoredRegions)): ?>
                        <?php render_ignored_regions($ignoredRegions); ?>
                    <?php endif; ?>

                    <?php if (!empty($forms)): ?>
                        <?php foreach ($forms as $formIndex => $form): ?>
                            <?php
                            $availableBlocks = get_available_blocks($form);
                            $formSlug = $sheetSlug . '_form_' . ($form['formulario_index'] ?? ($formIndex + 1));
                            ?>
                            <div class="form-card form-anchor" id="<?= e($formSlug) ?>">
                                <div class="form-title">
                                    <h3>Formulário <?= e($form['formulario_index'] ?? ($formIndex + 1)) ?></h3>
                                    <div class="form-badges">
                                        <?php if (!empty($form['header']['cliente'])): ?>
                                            <span class="badge">cliente: <?= e($form['header']['cliente']) ?></span>
                                        <?php endif; ?>
                                        <?php if (!empty($form['header']['titulo_plano'])): ?>
                                            <span class="badge">tipo: <?= e($form['header']['titulo_plano']) ?></span>
                                        <?php endif; ?>
                                        <?php if (!empty($form['identified_blocks'])): ?>
                                            <span class="badge">blocos: <?= e(count($form['identified_blocks'])) ?></span>
                                        <?php endif; ?>
                                    </div>
                                </div>

                                <?php if (!empty($form['identified_blocks'])): ?>
                                    <div class="sheet-info">
                                        <?php foreach ($form['identified_blocks'] as $identified): ?>
                                            <span class="chip"><?= e($identified) ?></span>
                                        <?php endforeach; ?>
                                    </div>
                                <?php endif; ?>

                                <div class="sheet-layout">
                                    <aside class="sidebar">
                                        <h3>Navegação do formulário</h3>
                                        <div class="nav-blocks">
                                            <?php foreach ($availableBlocks as $blockKey): ?>
                                                <?php
                                                $meta = block_meta($blockKey);
                                                $targetId = $formSlug . '_' . $blockKey;
                                                $countLabel = '';

                                                if ($blockKey !== 'header' && $blockKey !== 'resumo_analitico') {
                                                    $count = is_array($form[$blockKey] ?? null) ? count($form[$blockKey]) : 0;
                                                    $countLabel = $count . ' registro(s)';
                                                } elseif ($blockKey === 'resumo_analitico') {
                                                    $countLabel = 'visão consolidada';
                                                } else {
                                                    $countLabel = 'campos do cabeçalho';
                                                }
                                                ?>
                                                <a class="nav-link" href="#<?= e($targetId) ?>">
                                                    <?= e($meta['icon']) ?> <?= e($meta['label']) ?>
                                                    <small><?= e($countLabel) ?></small>
                                                </a>
                                            <?php endforeach; ?>

                                            <a class="nav-link" href="#<?= e($sheetSlug) ?>">
                                                ⬆ Voltar à aba
                                                <small><?= e($sheet['sheet_name'] ?? '') ?></small>
                                            </a>

                                            <a class="nav-link" href="#topo">
                                                ⬆ Voltar ao topo
                                                <small>resumo geral</small>
                                            </a>
                                        </div>
                                    </aside>

                                    <div class="content-area">
                                        <?php foreach ($availableBlocks as $blockKey): ?>
                                            <?php
                                            $meta = block_meta($blockKey);
                                            $targetId = $formSlug . '_' . $blockKey;
                                            ?>
                                            <section class="section-block" id="<?= e($targetId) ?>">
                                                <div class="section-header">
                                                    <span class="icon"><?= e($meta['icon']) ?></span>
                                                    <h3><?= e($meta['label']) ?></h3>
                                                </div>

                                                <?php if ($blockKey === 'header'): ?>
                                                    <?php render_key_value_table($form['header'], 'Campos do cabeçalho'); ?>

                                                <?php elseif ($blockKey === 'plano_negocios_rows'): ?>
                                                    <?php render_assoc_table($form['plano_negocios_rows'], 'Tabela simulada: plano_negocios'); ?>

                                                <?php elseif ($blockKey === 'historico_rows'): ?>
                                                    <?php render_assoc_table($form['historico_rows'], 'Tabela simulada: historico'); ?>

                                                <?php elseif ($blockKey === 'objetivos_rows'): ?>
                                                    <?php render_assoc_table($form['objetivos_rows'], 'Tabela simulada: objetivos'); ?>

                                                <?php elseif ($blockKey === 'descricao_investimento_rows'): ?>
                                                    <?php render_assoc_table($form['descricao_investimento_rows'], 'Tabela simulada: descricao_investimento'); ?>

                                                <?php elseif ($blockKey === 'contrapartidas_rows'): ?>
                                                    <?php render_assoc_table($form['contrapartidas_rows'], 'Tabela simulada: contrapartidas'); ?>

                                                <?php elseif ($blockKey === 'contrapartidas_itens_foco_rows'): ?>
                                                    <?php render_assoc_table($form['contrapartidas_itens_foco_rows'], 'Tabela simulada: contrapartidas_itens_foco'); ?>

                                                <?php elseif ($blockKey === 'encartes_obrigatorios_rows'): ?>
                                                    <?php render_assoc_table($form['encartes_obrigatorios_rows'], 'Tabela simulada: encartes_obrigatorios / sugestao_encartes'); ?>

                                                <?php elseif ($blockKey === 'cadastros_vinculados_rows'): ?>
                                                    <?php render_assoc_table($form['cadastros_vinculados_rows'], 'Tabela simulada: cadastros_vinculados'); ?>

                                                <?php elseif ($blockKey === 'investimentos_extras_rows'): ?>
                                                    <?php render_assoc_table($form['investimentos_extras_rows'], 'Tabela simulada: investimentos_extras'); ?>

                                                <?php elseif ($blockKey === 'resumo_analitico'): ?>
                                                    <?php render_client_summary($form); ?>

                                                <?php elseif ($blockKey === 'raw_blocks'): ?>
                                                    <?php render_block_summary($form['raw_blocks']); ?>
                                                <?php endif; ?>

                                                <a class="back-top" href="#<?= e($formSlug) ?>">↑ voltar ao topo do formulário</a>
                                            </section>
                                        <?php endforeach; ?>
                                    </div>
                                </div>
                            </div>
                        <?php endforeach; ?>
                    <?php else: ?>
                        <div class="bloco-card">
                            <h3>Nenhum formulário detectado nesta aba</h3>
                            <p class="muted">
                                Esta aba não apresentou blocos principais suficientes para ser considerada um formulário válido.
                            </p>
                        </div>
                    <?php endif; ?>
                </div>
            <?php endforeach; ?>
        <?php endif; ?>
    <?php endif; ?>
</div>
</body>
</html>