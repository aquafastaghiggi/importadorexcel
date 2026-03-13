<?php

declare(strict_types=1);

if ($argc < 2) {
    fwrite(STDERR, "Uso: php scripts/import_json_mysql_analitico.php caminho_do_json [database]\n");
    exit(1);
}

$jsonPath = $argv[1];
$database = $argv[2] ?? 'importador_excel_analitico';

if (!is_file($jsonPath)) {
    fwrite(STDERR, "Arquivo JSON nao encontrado: {$jsonPath}\n");
    exit(1);
}

$jsonRaw = (string) file_get_contents($jsonPath);
$jsonRaw = iconv('UTF-8', 'UTF-8//IGNORE', $jsonRaw);
$jsonRaw = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/u', '', $jsonRaw);
$data = json_decode($jsonRaw, true, 4096);
if (!is_array($data) || empty($data['success'])) {
    fwrite(STDERR, "JSON invalido ou processamento sem sucesso.\n");
    exit(1);
}

$pdo = new PDO(
    "mysql:host=localhost;dbname={$database};charset=utf8mb4",
    'root',
    '',
    [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
    ]
);

const TABLE_COLUMNS = [
    'historico' => ['linha_ordem', 'tipo_registro', 'ano_bloco', 'descricao', 'valor_original', 'valor_numerico', 'unidade', 'row_excel', 'linha_original'],
    'objetivos' => ['linha_ordem', 'tipo_registro', 'ano_bloco', 'descricao', 'valor_original', 'valor_numerico', 'unidade', 'gatilho_anual', 'faturamento', 'lava_roupas_liquido', 'amaciante_concentrado', 'observacao', 'lado_layout', 'row_excel', 'linha_original'],
    'descricao_investimento' => ['linha_ordem', 'tipo_registro', 'descricao', 'valor_original', 'valor_numerico', 'unidade', 'observacao', 'lado_layout', 'row_excel', 'linha_original'],
    'premissas_gerais' => ['linha_ordem', 'descricao', 'valor_original', 'row_excel', 'linha_original'],
    'contrapartidas' => ['bloco_origem', 'linha_ordem', 'quantidade', 'quantidade_numerica', 'tipo_acao', 'periodo', 'lojas', 'obs', 'row_excel', 'linha_original'],
    'contrapartidas_mensais' => ['bloco_origem', 'linha_ordem', 'mes', 'mes_numero', 'produto', 'row_excel', 'linha_original'],
    'contrapartidas_itens_foco' => ['linha_ordem', 'produto', 'volume_mensal', 'volume_periodo', 'volume_mensal_numerica', 'volume_periodo_numerica', 'unidade', 'row_excel', 'linha_original'],
    'objetivo_compra' => ['linha_ordem', 'produto', 'embalagem', 'objetivo_trimestral', 'objetivo_trimestral_numerico', 'row_excel', 'linha_original'],
    'stok_objetivo_compra' => ['linha_ordem', 'produto', 'embalagem', 'janeiro_junho', 'julho_dezembro', 'periodo', 'atingimento', 'investimento_percentual', 'investimento_percentual_numerico', 'row_excel', 'linha_original'],
    'cadastros_vinculados' => ['bloco_origem', 'linha_ordem', 'secao_interna', 'categoria_secao', 'produto', 'litragem', 'versao', 'abrangencia', 'row_excel', 'linha_original'],
    'investimentos_extras' => ['linha_ordem', 'descricao', 'valor_original', 'valor_numerico', 'unidade', 'row_excel', 'linha_original'],
];

function insertRows(PDO $pdo, string $table, int $formularioId, array $rows, array $extra = []): void
{
    $allowed = TABLE_COLUMNS[$table] ?? [];
    if ($allowed === []) {
        return;
    }

    foreach ($rows as $row) {
        if (!is_array($row) || $row === []) {
            continue;
        }

        $payload = ['formulario_id' => $formularioId];
        foreach ($allowed as $column) {
            if (array_key_exists($column, $extra)) {
                $payload[$column] = $extra[$column];
                continue;
            }
            if (array_key_exists($column, $row)) {
                $payload[$column] = $row[$column];
            }
        }

        if (count($payload) <= 1) {
            continue;
        }

        $columns = array_keys($payload);
        $placeholders = array_map(static fn(string $col): string => ':' . $col, $columns);
        $sql = sprintf(
            'INSERT INTO %s (%s) VALUES (%s)',
            $table,
            implode(', ', $columns),
            implode(', ', $placeholders)
        );

        $stmt = $pdo->prepare($sql);
        foreach ($payload as $key => $value) {
            if (is_bool($value)) {
                $value = $value ? 1 : 0;
            }
            $stmt->bindValue(':' . $key, $value);
        }
        $stmt->execute();
    }
}

$importStmt = $pdo->prepare(
    'INSERT INTO importacoes (arquivo_nome, processado_em, total_abas, status_importacao) VALUES (:arquivo_nome, :processado_em, :total_abas, :status_importacao)'
);
$formStmt = $pdo->prepare(
    'INSERT INTO formularios (
        importacao_id, aba_origem, formulario_index, cliente, titulo_plano, numero_acordo,
        periodo_original, periodo_inicio, periodo_fim, periodo_normalizado, periodo_status,
        periodo_acoes_original, periodo_acoes_inicio, periodo_acoes_fim, periodo_acoes_normalizado, periodo_acoes_status
    ) VALUES (
        :importacao_id, :aba_origem, :formulario_index, :cliente, :titulo_plano, :numero_acordo,
        :periodo_original, :periodo_inicio, :periodo_fim, :periodo_normalizado, :periodo_status,
        :periodo_acoes_original, :periodo_acoes_inicio, :periodo_acoes_fim, :periodo_acoes_normalizado, :periodo_acoes_status
    )'
);

$pdo->beginTransaction();
try {
    $importStmt->execute([
        ':arquivo_nome' => $data['file_name'] ?? basename($jsonPath),
        ':processado_em' => $data['processed_at'] ?? date('Y-m-d H:i:s'),
        ':total_abas' => $data['total_sheets'] ?? null,
        ':status_importacao' => 'processado',
    ]);
    $importacaoId = (int) $pdo->lastInsertId();

    $formulariosInseridos = 0;
    foreach (($data['sheets'] ?? []) as $sheet) {
        foreach (($sheet['forms'] ?? []) as $form) {
            $header = $form['header'] ?? [];
            $formStmt->execute([
                ':importacao_id' => $importacaoId,
                ':aba_origem' => $form['sheet_name'] ?? $sheet['sheet_name'] ?? '',
                ':formulario_index' => $form['formulario_index'] ?? 1,
                ':cliente' => $header['cliente'] ?? null,
                ':titulo_plano' => $header['titulo_plano'] ?? null,
                ':numero_acordo' => $header['numero_acordo'] ?? null,
                ':periodo_original' => $header['periodo_original'] ?? null,
                ':periodo_inicio' => $header['periodo_inicio'] ?? null,
                ':periodo_fim' => $header['periodo_fim'] ?? null,
                ':periodo_normalizado' => $header['periodo_normalizado'] ?? null,
                ':periodo_status' => $header['periodo_status'] ?? null,
                ':periodo_acoes_original' => $header['periodo_acoes_original'] ?? null,
                ':periodo_acoes_inicio' => $header['periodo_acoes_inicio'] ?? null,
                ':periodo_acoes_fim' => $header['periodo_acoes_fim'] ?? null,
                ':periodo_acoes_normalizado' => $header['periodo_acoes_normalizado'] ?? null,
                ':periodo_acoes_status' => $header['periodo_acoes_status'] ?? null,
            ]);
            $formularioId = (int) $pdo->lastInsertId();
            $formulariosInseridos++;

            insertRows($pdo, 'historico', $formularioId, $form['historico_rows'] ?? []);
            insertRows($pdo, 'objetivos', $formularioId, $form['objetivos_rows'] ?? []);
            insertRows($pdo, 'objetivos', $formularioId, $form['objetivos_compactos_rows'] ?? []);
            insertRows($pdo, 'descricao_investimento', $formularioId, $form['descricao_investimento_rows'] ?? []);
            insertRows($pdo, 'descricao_investimento', $formularioId, $form['descricao_investimento_compacto_rows'] ?? []);
            insertRows($pdo, 'premissas_gerais', $formularioId, $form['premissas_gerais_rows'] ?? []);
            insertRows($pdo, 'contrapartidas', $formularioId, $form['contrapartidas_rows'] ?? [], ['bloco_origem' => 'contrapartidas']);
            insertRows($pdo, 'contrapartidas', $formularioId, $form['contrapartidas_acoes_rows'] ?? [], ['bloco_origem' => 'contrapartidas_acoes']);
            insertRows($pdo, 'contrapartidas', $formularioId, $form['contrapartidas_encartes_mensal_rows'] ?? [], ['bloco_origem' => 'contrapartidas_encartes_mensal']);
            insertRows($pdo, 'contrapartidas_itens_foco', $formularioId, $form['contrapartidas_itens_foco_rows'] ?? []);
            insertRows($pdo, 'contrapartidas_mensais', $formularioId, $form['encartes_obrigatorios_rows'] ?? [], ['bloco_origem' => 'encartes_obrigatorios']);
            insertRows($pdo, 'contrapartidas_mensais', $formularioId, $form['encartes_sugestao_rows'] ?? [], ['bloco_origem' => 'encartes_sugestao']);
            insertRows($pdo, 'contrapartidas_mensais', $formularioId, $form['contrapartidas_itens_acao_obrigatorios_rows'] ?? [], ['bloco_origem' => 'contrapartidas_itens_acao_obrigatorios']);
            insertRows($pdo, 'contrapartidas_mensais', $formularioId, $form['contrapartidas_itens_ponta_gondola_obrigatorios_rows'] ?? [], ['bloco_origem' => 'contrapartidas_itens_ponta_gondola_obrigatorios']);
            insertRows($pdo, 'objetivo_compra', $formularioId, $form['objetivo_compra_rows'] ?? []);
            insertRows($pdo, 'stok_objetivo_compra', $formularioId, $form['stok_objetivo_compra_rows'] ?? []);
            insertRows($pdo, 'cadastros_vinculados', $formularioId, $form['cadastros_vinculados_rows'] ?? [], ['bloco_origem' => 'cadastros_vinculados']);
            insertRows($pdo, 'cadastros_vinculados', $formularioId, $form['situacao_liberacao_rows'] ?? [], ['bloco_origem' => 'situacao_liberacao']);
            insertRows($pdo, 'cadastros_vinculados', $formularioId, $form['oportunidades_cadastros_rows'] ?? [], ['bloco_origem' => 'oportunidades_cadastros']);
            insertRows($pdo, 'cadastros_vinculados', $formularioId, $form['sugestao_liberacao_rows'] ?? [], ['bloco_origem' => 'sugestao_liberacao']);
            insertRows($pdo, 'investimentos_extras', $formularioId, $form['investimentos_extras_rows'] ?? []);
        }
    }

    $pdo->commit();
    echo json_encode([
        'success' => true,
        'database' => $database,
        'importacao_id' => $importacaoId,
        'formularios_inseridos' => $formulariosInseridos,
    ], JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) . PHP_EOL;
} catch (Throwable $e) {
    if ($pdo->inTransaction()) {
        $pdo->rollBack();
    }
    fwrite(STDERR, 'Erro ao importar: ' . $e->getMessage() . PHP_EOL);
    exit(1);
}