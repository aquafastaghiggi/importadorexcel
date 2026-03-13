CREATE DATABASE IF NOT EXISTS importador_excel_analitico
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE importador_excel_analitico;

CREATE TABLE IF NOT EXISTS importacoes (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    arquivo_nome VARCHAR(255) NOT NULL,
    processado_em DATETIME NOT NULL,
    total_abas INT NULL,
    status_importacao VARCHAR(40) NOT NULL DEFAULT 'processado',
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
    PRIMARY KEY (id),
    KEY idx_formularios_importacao (importacao_id),
    KEY idx_formularios_cliente (cliente),
    CONSTRAINT fk_analitico_formularios_importacoes
        FOREIGN KEY (importacao_id) REFERENCES importacoes(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS historico (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    ano_bloco INT NULL,
    descricao VARCHAR(500) NULL,
    valor_original VARCHAR(255) NULL,
    valor_numerico DECIMAL(18,4) NULL,
    unidade VARCHAR(30) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_historico_formulario (formulario_id),
    CONSTRAINT fk_historico_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS objetivos (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    tipo_registro VARCHAR(80) NULL,
    ano_bloco INT NULL,
    descricao VARCHAR(500) NULL,
    valor_original VARCHAR(255) NULL,
    valor_numerico DECIMAL(18,4) NULL,
    unidade VARCHAR(30) NULL,
    gatilho_anual VARCHAR(255) NULL,
    faturamento VARCHAR(255) NULL,
    lava_roupas_liquido VARCHAR(255) NULL,
    amaciante_concentrado VARCHAR(255) NULL,
    observacao TEXT NULL,
    lado_layout VARCHAR(40) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_objetivos_formulario (formulario_id),
    CONSTRAINT fk_objetivos_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS descricao_investimento (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    tipo_registro VARCHAR(80) NULL,
    descricao VARCHAR(500) NULL,
    valor_original VARCHAR(255) NULL,
    valor_numerico DECIMAL(18,4) NULL,
    unidade VARCHAR(30) NULL,
    observacao TEXT NULL,
    lado_layout VARCHAR(40) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_desc_inv_formulario (formulario_id),
    CONSTRAINT fk_desc_inv_formulario_analitico
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS premissas_gerais (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    descricao TEXT NULL,
    valor_original TEXT NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_premissas_formulario (formulario_id),
    CONSTRAINT fk_premissas_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS contrapartidas (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    bloco_origem VARCHAR(120) NOT NULL,
    linha_ordem INT NULL,
    quantidade VARCHAR(100) NULL,
    quantidade_numerica DECIMAL(18,4) NULL,
    tipo_acao VARCHAR(255) NULL,
    periodo VARCHAR(255) NULL,
    lojas VARCHAR(255) NULL,
    obs VARCHAR(255) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_contrapartidas_formulario (formulario_id),
    KEY idx_contrapartidas_bloco (bloco_origem),
    CONSTRAINT fk_contrapartidas_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS contrapartidas_mensais (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    bloco_origem VARCHAR(120) NOT NULL,
    linha_ordem INT NULL,
    mes VARCHAR(40) NULL,
    mes_numero INT NULL,
    produto VARCHAR(255) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_contrapartidas_mensais_formulario (formulario_id),
    KEY idx_contrapartidas_mensais_bloco (bloco_origem),
    CONSTRAINT fk_contrapartidas_mensais_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS contrapartidas_itens_foco (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    produto VARCHAR(255) NULL,
    volume_mensal VARCHAR(255) NULL,
    volume_periodo VARCHAR(255) NULL,
    volume_mensal_numerica DECIMAL(18,4) NULL,
    volume_periodo_numerica DECIMAL(18,4) NULL,
    unidade VARCHAR(30) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_itens_foco_formulario (formulario_id),
    CONSTRAINT fk_itens_foco_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS objetivo_compra (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    produto VARCHAR(255) NULL,
    embalagem VARCHAR(255) NULL,
    objetivo_trimestral VARCHAR(255) NULL,
    objetivo_trimestral_numerico DECIMAL(18,4) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_objetivo_compra_formulario (formulario_id),
    CONSTRAINT fk_objetivo_compra_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS cadastros_vinculados (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    bloco_origem VARCHAR(120) NOT NULL,
    linha_ordem INT NULL,
    secao_interna VARCHAR(255) NULL,
    categoria_secao VARCHAR(80) NULL,
    produto VARCHAR(255) NULL,
    litragem VARCHAR(120) NULL,
    versao VARCHAR(255) NULL,
    abrangencia VARCHAR(255) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_cadastros_formulario (formulario_id),
    KEY idx_cadastros_bloco (bloco_origem),
    CONSTRAINT fk_cadastros_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS investimentos_extras (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    descricao VARCHAR(500) NULL,
    valor_original TEXT NULL,
    valor_numerico DECIMAL(18,4) NULL,
    unidade VARCHAR(30) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_inv_extras_formulario (formulario_id),
    CONSTRAINT fk_inv_extras_formulario_analitico
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;
CREATE TABLE IF NOT EXISTS stok_objetivo_compra (
    id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    formulario_id BIGINT UNSIGNED NOT NULL,
    linha_ordem INT NULL,
    produto VARCHAR(255) NULL,
    embalagem VARCHAR(255) NULL,
    janeiro_junho VARCHAR(255) NULL,
    julho_dezembro VARCHAR(255) NULL,
    periodo VARCHAR(255) NULL,
    atingimento VARCHAR(255) NULL,
    investimento_percentual VARCHAR(255) NULL,
    investimento_percentual_numerico DECIMAL(18,4) NULL,
    row_excel INT NULL,
    linha_original TEXT NULL,
    PRIMARY KEY (id),
    KEY idx_stok_objetivo_compra_formulario (formulario_id),
    CONSTRAINT fk_stok_objetivo_compra_formulario
        FOREIGN KEY (formulario_id) REFERENCES formularios(id)
) ENGINE=InnoDB;