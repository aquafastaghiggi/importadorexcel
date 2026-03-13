CREATE DATABASE IF NOT EXISTS importador_excel
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE importador_excel;

CREATE TABLE IF NOT EXISTS importacoes (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    arquivo_nome VARCHAR(255) NOT NULL,
    processado_em DATETIME NOT NULL,
    total_abas INT NULL,
    status_importacao VARCHAR(40) NOT NULL DEFAULT 'processado',
    observacao TEXT NULL,
    PRIMARY KEY (id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS formularios (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    importacao_id BIGINT UNSIGNED NOT NULL,
    aba_origem VARCHAR(150) NOT NULL,
    formulario_index INT NOT NULL,
    cliente VARCHAR(255) NULL,
    titulo_plano VARCHAR(120) NULL,
    numero_acordo VARCHAR(80) NULL,
    periodo_original VARCHAR(120) NULL,
    periodo_inicio VARCHAR(20) NULL,
    periodo_fim VARCHAR(20) NULL,
    periodo_normalizado VARCHAR(60) NULL,
    periodo_status VARCHAR(40) NULL,
    periodo_acoes_original VARCHAR(120) NULL,
    periodo_acoes_inicio VARCHAR(20) NULL,
    periodo_acoes_fim VARCHAR(20) NULL,
    periodo_acoes_normalizado VARCHAR(60) NULL,
    periodo_acoes_status VARCHAR(40) NULL,
    nome_ajuste_colunar VARCHAR(80) NULL DEFAULT 'ajuste colunar',
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    KEY idx_formularios_importacao (importacao_id),
    KEY idx_formularios_cliente (cliente),
    CONSTRAINT fk_formularios_importacoes
        FOREIGN KEY (importacao_id) REFERENCES importacoes(id)
) ENGINE=InnoDB;
CREATE TABLE IF NOT EXISTS formulario_json (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    dados_json LONGTEXT NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_formulario_json (formulario_id),
    CONSTRAINT fk_formulario_json_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS formulario_blocos (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    bloco_nome VARCHAR(120) NOT NULL,
    total_registros INT NOT NULL DEFAULT 0,
    dados_json LONGTEXT NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    KEY idx_formulario_blocos_formulario (formulario_id),
    KEY idx_formulario_blocos_nome (bloco_nome),
    CONSTRAINT fk_formulario_blocos_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS formulario_descricao_investimento_colunar (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    origem_bloco VARCHAR(80) NULL,
    ajuste_colunar_ativo TINYINT(1) NOT NULL DEFAULT 1,

    forma_pagamento VARCHAR(255) NULL,
    forma_pagamento_texto VARCHAR(255) NULL,
    forma_pagamento_label VARCHAR(255) NULL,

    valor_investimento VARCHAR(255) NULL,
    valor_investimento_texto VARCHAR(255) NULL,
    valor_investimento_label VARCHAR(255) NULL,
    valor_investimento_numero DECIMAL(18,4) NULL,

    valor_investimento_total VARCHAR(255) NULL,
    valor_investimento_total_texto VARCHAR(255) NULL,
    valor_investimento_total_label VARCHAR(255) NULL,
    valor_investimento_total_numero DECIMAL(18,4) NULL,

    valor_investimento_mensal VARCHAR(255) NULL,
    valor_investimento_mensal_texto VARCHAR(255) NULL,
    valor_investimento_mensal_label VARCHAR(255) NULL,
    valor_investimento_mensal_numero DECIMAL(18,4) NULL,

    investimento_mensal_total VARCHAR(255) NULL,
    investimento_mensal_total_texto VARCHAR(255) NULL,
    investimento_mensal_total_label VARCHAR(255) NULL,
    investimento_mensal_total_numero DECIMAL(18,4) NULL,

    percentual_investimento_total VARCHAR(255) NULL,
    percentual_investimento_total_texto VARCHAR(255) NULL,
    percentual_investimento_total_label VARCHAR(255) NULL,
    percentual_investimento_total_percentual DECIMAL(10,4) NULL,

    percentual_investimento_mensal_total VARCHAR(255) NULL,
    percentual_investimento_mensal_total_texto VARCHAR(255) NULL,
    percentual_investimento_mensal_total_label VARCHAR(255) NULL,
    percentual_investimento_mensal_total_percentual DECIMAL(10,4) NULL,

    percentual_investimento_objetivo VARCHAR(255) NULL,
    percentual_investimento_objetivo_texto VARCHAR(255) NULL,
    percentual_investimento_objetivo_label VARCHAR(255) NULL,
    percentual_investimento_objetivo_percentual DECIMAL(10,4) NULL,

    percentual_investimento_volume_global VARCHAR(255) NULL,
    percentual_investimento_volume_global_texto VARCHAR(255) NULL,
    percentual_investimento_volume_global_label VARCHAR(255) NULL,
    percentual_investimento_volume_global_percentual DECIMAL(10,4) NULL,

    percentual_investimento_acoes VARCHAR(255) NULL,
    percentual_investimento_acoes_texto VARCHAR(255) NULL,
    percentual_investimento_acoes_label VARCHAR(255) NULL,
    percentual_investimento_acoes_percentual DECIMAL(10,4) NULL,

    percentual_investimento_volume_global_acoes VARCHAR(255) NULL,
    percentual_investimento_volume_global_acoes_texto VARCHAR(255) NULL,
    percentual_investimento_volume_global_acoes_label VARCHAR(255) NULL,
    percentual_investimento_volume_global_acoes_percentual DECIMAL(10,4) NULL,

    percentual_investimento_calendario_acoes VARCHAR(255) NULL,
    percentual_investimento_calendario_acoes_texto VARCHAR(255) NULL,
    percentual_investimento_calendario_acoes_label VARCHAR(255) NULL,
    percentual_investimento_calendario_acoes_percentual DECIMAL(10,4) NULL,

    percentual_investimento_pontas VARCHAR(255) NULL,
    percentual_investimento_pontas_texto VARCHAR(255) NULL,
    percentual_investimento_pontas_label VARCHAR(255) NULL,
    percentual_investimento_pontas_percentual DECIMAL(10,4) NULL,

    percentual_investimento_pedidos VARCHAR(255) NULL,
    percentual_investimento_pedidos_texto VARCHAR(255) NULL,
    percentual_investimento_pedidos_label VARCHAR(255) NULL,
    percentual_investimento_pedidos_percentual DECIMAL(10,4) NULL,

    percentual_investimento_jbp VARCHAR(255) NULL,
    percentual_investimento_jbp_texto VARCHAR(255) NULL,
    percentual_investimento_jbp_label VARCHAR(255) NULL,
    percentual_investimento_jbp_percentual DECIMAL(10,4) NULL,

    gatilho_adicional_volume_total VARCHAR(255) NULL,
    gatilho_adicional_volume_total_texto VARCHAR(255) NULL,
    gatilho_adicional_volume_total_label VARCHAR(255) NULL,
    gatilho_adicional_volume_total_numero DECIMAL(18,4) NULL,

    crescimento_valor VARCHAR(255) NULL,
    crescimento_valor_texto VARCHAR(255) NULL,
    crescimento_valor_label VARCHAR(255) NULL,
    crescimento_valor_percentual DECIMAL(10,4) NULL,

    campos_nao_mapeados_json LONGTEXT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_desc_inv_formulario (formulario_id),
    CONSTRAINT fk_desc_inv_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS formulario_investimentos_extras_colunar (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    origem_bloco VARCHAR(80) NULL,
    ajuste_colunar_ativo TINYINT(1) NOT NULL DEFAULT 1,

    promotor TINYINT(1) NULL,
    promotor_texto VARCHAR(255) NULL,
    promotor_label VARCHAR(255) NULL,

    retorno_financeiro TINYINT(1) NULL,
    retorno_financeiro_texto VARCHAR(255) NULL,
    retorno_financeiro_label VARCHAR(255) NULL,

    liberacao_dados_scanntech_2025 VARCHAR(255) NULL,
    liberacao_dados_scanntech_2025_texto VARCHAR(255) NULL,
    liberacao_dados_scanntech_2025_label VARCHAR(255) NULL,
    liberacao_dados_scanntech_2025_numero DECIMAL(18,4) NULL,

    adicional_crescimento_categorias VARCHAR(255) NULL,
    adicional_crescimento_categorias_texto VARCHAR(255) NULL,
    adicional_crescimento_categorias_label VARCHAR(255) NULL,

    campos_nao_mapeados_json LONGTEXT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (id),
    UNIQUE KEY uq_inv_extras_formulario (formulario_id),
    CONSTRAINT fk_inv_extras_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

