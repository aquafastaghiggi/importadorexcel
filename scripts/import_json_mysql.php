<?php

declare(strict_types=1);

if ($argc < 2) {
    fwrite(STDERR, "Uso: php scripts/import_json_mysql.php caminho_do_json [database]\n");
    exit(1);
}

$jsonPath = $argv[1];
$database = $argv[2] ?? 'importador_excel';

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

function json_encode_mysql(mixed $value): string
{
    return json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
}

function normalize_bool_for_db(mixed $value): mixed
{
    if (is_bool($value)) {
        return $value ? 1 : 0;
    }
    return $value;
}

function insertColumnarRow(PDO $pdo, string $table, int $formularioId, array $payload): void
{
    if (empty($payload)) {
        return;
    }

    $payload['formulario_id'] = $formularioId;
    if (isset($payload['campos_nao_mapeados'])) {
        $payload['campos_nao_mapeados_json'] = json_encode_mysql($payload['campos_nao_mapeados']);
        unset($payload['campos_nao_mapeados']);
    }

    foreach ($payload as $key => $value) {
        $payload[$key] = normalize_bool_for_db($value);
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
        $stmt->bindValue(':' . $key, $value);
    }
    $stmt->execute();
}

$importStmt = $pdo->prepare(
    'INSERT INTO importacoes (arquivo_nome, processado_em, total_abas, status_importacao) VALUES (:arquivo_nome, :processado_em, :total_abas, :status_importacao)'
);
$formStmt = $pdo->prepare(
    'INSERT INTO formularios (
        importacao_id, aba_origem, formulario_index, cliente, titulo_plano, numero_acordo,
        periodo_original, periodo_inicio, periodo_fim, periodo_normalizado, periodo_status,
        periodo_acoes_original, periodo_acoes_inicio, periodo_acoes_fim, periodo_acoes_normalizado, periodo_acoes_status,
        nome_ajuste_colunar
    ) VALUES (
        :importacao_id, :aba_origem, :formulario_index, :cliente, :titulo_plano, :numero_acordo,
        :periodo_original, :periodo_inicio, :periodo_fim, :periodo_normalizado, :periodo_status,
        :periodo_acoes_original, :periodo_acoes_inicio, :periodo_acoes_fim, :periodo_acoes_normalizado, :periodo_acoes_status,
        :nome_ajuste_colunar
    )'
);
$formJsonStmt = $pdo->prepare('INSERT INTO formulario_json (formulario_id, dados_json) VALUES (:formulario_id, :dados_json)');
$blockStmt = $pdo->prepare('INSERT INTO formulario_blocos (formulario_id, bloco_nome, total_registros, dados_json) VALUES (:formulario_id, :bloco_nome, :total_registros, :dados_json)');

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
            $ajusteColunar = $form['ajuste_colunar'] ?? [];

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
                ':nome_ajuste_colunar' => $ajusteColunar['nome_ajuste'] ?? 'ajuste colunar',
            ]);
            $formularioId = (int) $pdo->lastInsertId();
            $formulariosInseridos++;

            $formJsonStmt->execute([
                ':formulario_id' => $formularioId,
                ':dados_json' => json_encode_mysql($form),
            ]);

            foreach ($form as $blockName => $blockValue) {
                if (!str_ends_with((string) $blockName, '_rows')) {
                    continue;
                }
                if (!is_array($blockValue) || $blockValue === []) {
                    continue;
                }
                $blockStmt->execute([
                    ':formulario_id' => $formularioId,
                    ':bloco_nome' => $blockName,
                    ':total_registros' => count($blockValue),
                    ':dados_json' => json_encode_mysql($blockValue),
                ]);
            }

            if (!empty($ajusteColunar['descricao_investimento'])) {
                insertColumnarRow($pdo, 'formulario_descricao_investimento_colunar', $formularioId, $ajusteColunar['descricao_investimento']);
            }
            if (!empty($ajusteColunar['investimentos_extras'])) {
                insertColumnarRow($pdo, 'formulario_investimentos_extras_colunar', $formularioId, $ajusteColunar['investimentos_extras']);
            }
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