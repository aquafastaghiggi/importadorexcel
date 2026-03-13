USE importador_excel_analitico;

DROP VIEW IF EXISTS vw_formularios;
DROP VIEW IF EXISTS vw_historico;
DROP VIEW IF EXISTS vw_objetivos;
DROP VIEW IF EXISTS vw_descricao_investimento;
DROP VIEW IF EXISTS vw_premissas_gerais;
DROP VIEW IF EXISTS vw_contrapartidas;
DROP VIEW IF EXISTS vw_contrapartidas_mensais;
DROP VIEW IF EXISTS vw_contrapartidas_itens_foco;
DROP VIEW IF EXISTS vw_objetivo_compra;
DROP VIEW IF EXISTS vw_stok_objetivo_compra;
DROP VIEW IF EXISTS vw_cadastros_vinculados;
DROP VIEW IF EXISTS vw_investimentos_extras;

CREATE VIEW vw_formularios AS
SELECT
    f.id AS formulario_id,
    f.importacao_id,
    i.arquivo_nome,
    i.processado_em,
    i.total_abas,
    f.aba_origem,
    f.formulario_index,
    f.cliente,
    f.titulo_plano,
    f.numero_acordo,
    f.periodo_original,
    f.periodo_inicio,
    f.periodo_fim,
    f.periodo_normalizado,
    f.periodo_status,
    f.periodo_acoes_original,
    f.periodo_acoes_inicio,
    f.periodo_acoes_fim,
    f.periodo_acoes_normalizado,
    f.periodo_acoes_status
FROM formularios f
JOIN importacoes i ON i.id = f.importacao_id;

CREATE VIEW vw_historico AS
SELECT
    vf.*,
    h.id AS historico_id,
    h.linha_ordem,
    h.tipo_registro,
    h.ano_bloco,
    h.descricao,
    h.valor_original,
    h.valor_numerico,
    h.unidade,
    h.row_excel,
    h.linha_original
FROM historico h
JOIN vw_formularios vf ON vf.formulario_id = h.formulario_id;

CREATE VIEW vw_objetivos AS
SELECT
    vf.*,
    o.id AS objetivo_id,
    o.linha_ordem,
    o.tipo_registro,
    o.ano_bloco,
    o.descricao,
    o.valor_original,
    o.valor_numerico,
    o.unidade,
    o.gatilho_anual,
    o.faturamento,
    o.lava_roupas_liquido,
    o.amaciante_concentrado,
    o.observacao,
    o.lado_layout,
    o.row_excel,
    o.linha_original
FROM objetivos o
JOIN vw_formularios vf ON vf.formulario_id = o.formulario_id;

CREATE VIEW vw_descricao_investimento AS
SELECT
    vf.*,
    d.id AS descricao_investimento_id,
    d.linha_ordem,
    d.tipo_registro,
    d.descricao,
    d.valor_original,
    d.valor_numerico,
    d.unidade,
    d.observacao,
    d.lado_layout,
    d.row_excel,
    d.linha_original
FROM descricao_investimento d
JOIN vw_formularios vf ON vf.formulario_id = d.formulario_id;

CREATE VIEW vw_premissas_gerais AS
SELECT
    vf.*,
    p.id AS premissa_id,
    p.linha_ordem,
    p.descricao,
    p.valor_original,
    p.row_excel,
    p.linha_original
FROM premissas_gerais p
JOIN vw_formularios vf ON vf.formulario_id = p.formulario_id;

CREATE VIEW vw_contrapartidas AS
SELECT
    vf.*,
    c.id AS contrapartida_id,
    c.bloco_origem,
    c.linha_ordem,
    c.quantidade,
    c.quantidade_numerica,
    c.tipo_acao,
    c.periodo,
    c.lojas,
    c.obs,
    c.row_excel,
    c.linha_original
FROM contrapartidas c
JOIN vw_formularios vf ON vf.formulario_id = c.formulario_id;

CREATE VIEW vw_contrapartidas_mensais AS
SELECT
    vf.*,
    cm.id AS contrapartida_mensal_id,
    cm.bloco_origem,
    cm.linha_ordem,
    cm.mes,
    cm.mes_numero,
    cm.produto,
    cm.row_excel,
    cm.linha_original
FROM contrapartidas_mensais cm
JOIN vw_formularios vf ON vf.formulario_id = cm.formulario_id;

CREATE VIEW vw_contrapartidas_itens_foco AS
SELECT
    vf.*,
    cif.id AS contrapartida_item_foco_id,
    cif.linha_ordem,
    cif.produto,
    cif.volume_mensal,
    cif.volume_periodo,
    cif.volume_mensal_numerica,
    cif.volume_periodo_numerica,
    cif.unidade,
    cif.row_excel,
    cif.linha_original
FROM contrapartidas_itens_foco cif
JOIN vw_formularios vf ON vf.formulario_id = cif.formulario_id;

CREATE VIEW vw_objetivo_compra AS
SELECT
    vf.*,
    oc.id AS objetivo_compra_id,
    oc.linha_ordem,
    oc.produto,
    oc.embalagem,
    oc.objetivo_trimestral,
    oc.objetivo_trimestral_numerico,
    oc.row_excel,
    oc.linha_original
FROM objetivo_compra oc
JOIN vw_formularios vf ON vf.formulario_id = oc.formulario_id;

CREATE VIEW vw_stok_objetivo_compra AS
SELECT
    vf.*,
    soc.id AS stok_objetivo_compra_id,
    soc.linha_ordem,
    soc.produto,
    soc.embalagem,
    soc.janeiro_junho,
    soc.julho_dezembro,
    soc.periodo,
    soc.atingimento,
    soc.investimento_percentual,
    soc.investimento_percentual_numerico,
    soc.row_excel,
    soc.linha_original
FROM stok_objetivo_compra soc
JOIN vw_formularios vf ON vf.formulario_id = soc.formulario_id;

CREATE VIEW vw_cadastros_vinculados AS
SELECT
    vf.*,
    cv.id AS cadastro_vinculado_id,
    cv.bloco_origem,
    cv.linha_ordem,
    cv.secao_interna,
    cv.categoria_secao,
    cv.produto,
    cv.litragem,
    cv.versao,
    cv.abrangencia,
    cv.row_excel,
    cv.linha_original
FROM cadastros_vinculados cv
JOIN vw_formularios vf ON vf.formulario_id = cv.formulario_id;

CREATE VIEW vw_investimentos_extras AS
SELECT
    vf.*,
    ie.id AS investimento_extra_id,
    ie.linha_ordem,
    ie.descricao,
    ie.valor_original,
    ie.valor_numerico,
    ie.unidade,
    ie.row_excel,
    ie.linha_original
FROM investimentos_extras ie
JOIN vw_formularios vf ON vf.formulario_id = ie.formulario_id;