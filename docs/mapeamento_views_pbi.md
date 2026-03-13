# Mapeamento de Views para Power BI

Este documento ajuda quem vai consumir o banco `importador_excel_analitico` no Power BI.

A ideia geral e simples:
- usar `vw_formularios` como base de contexto do formulario
- usar as demais views conforme o bloco que aparece no PHP
- cada view ja traz os dados do formulario junto, evitando join manual na maior parte dos casos

## Banco e schema

- Banco: `importador_excel_analitico`
- Views criadas em: `database/mysql_analitico_views.sql`

## View principal

### `vw_formularios`

Use quando precisar do contexto geral de cada formulario.

Campos mais importantes:
- `formulario_id`
- `importacao_id`
- `arquivo_nome`
- `processado_em`
- `aba_origem`
- `formulario_index`
- `cliente`
- `titulo_plano`
- `numero_acordo`
- `periodo_original`
- `periodo_inicio`
- `periodo_fim`
- `periodo_normalizado`
- `periodo_acoes_original`
- `periodo_acoes_inicio`
- `periodo_acoes_fim`

## Mapeamento PHP -> View

### Historico
- Bloco no PHP: `historico_rows`
- View: `vw_historico`
- Campos de negocio:
  - `descricao`
  - `valor_original`
  - `valor_numerico`
  - `unidade`
  - `ano_bloco`

### Objetivos
- Bloco no PHP: `objetivos_rows`
- View: `vw_objetivos`
- Campos de negocio mais comuns:
  - `descricao`
  - `valor_original`
  - `valor_numerico`
  - `unidade`
- Campos especiais de tabela:
  - `gatilho_anual`
  - `faturamento`
  - `lava_roupas_liquido`
  - `amaciante_concentrado`
- Campos auxiliares:
  - `tipo_registro`
  - `observacao`
  - `lado_layout`

### Descricao do investimento
- Bloco no PHP: `descricao_investimento_rows`
- View: `vw_descricao_investimento`
- Campos de negocio:
  - `descricao`
  - `valor_original`
  - `valor_numerico`
  - `unidade`
  - `observacao`
  - `lado_layout`

### Premissas gerais
- Bloco no PHP: `premissas_gerais_rows`
- View: `vw_premissas_gerais`
- Campos de negocio:
  - `descricao`
  - `valor_original`

### Contrapartidas
- Bloco no PHP: `contrapartidas_rows`
- View: `vw_contrapartidas`
- Campos de negocio:
  - `bloco_origem`
  - `quantidade`
  - `quantidade_numerica`
  - `tipo_acao`
  - `periodo`
  - `lojas`
  - `obs`

### Contrapartidas mensais
Usada para blocos de grade por mes.

- Blocos no PHP:
  - `encartes_obrigatorios_rows`
  - `encartes_sugestao_rows`
  - `contrapartidas_itens_acao_obrigatorios_rows`
  - `contrapartidas_itens_ponta_gondola_obrigatorios_rows`
- View: `vw_contrapartidas_mensais`
- Campos de negocio:
  - `bloco_origem`
  - `mes`
  - `mes_numero`
  - `produto`

### Contrapartidas - itens foco
- Bloco no PHP: `contrapartidas_itens_foco_rows`
- View: `vw_contrapartidas_itens_foco`
- Campos de negocio:
  - `produto`
  - `volume_mensal`
  - `volume_periodo`
  - `volume_mensal_numerica`
  - `volume_periodo_numerica`
  - `unidade`

### Objetivo compra
- Bloco no PHP: `objetivo_compra_rows`
- View: `vw_objetivo_compra`
- Campos de negocio:
  - `produto`
  - `embalagem`
  - `objetivo_trimestral`
  - `objetivo_trimestral_numerico`

### STOK - objetivo compra
- Bloco no PHP: `stok_objetivo_compra_rows`
- View: `vw_stok_objetivo_compra`
- Campos de negocio:
  - `produto`
  - `embalagem`
  - `janeiro_junho`
  - `julho_dezembro`
  - `periodo`
  - `atingimento`
  - `investimento_percentual`
  - `investimento_percentual_numerico`

### Cadastros vinculados
- Blocos no PHP:
  - `cadastros_vinculados_rows`
  - `situacao_liberacao_rows`
  - `oportunidades_cadastros_rows`
  - `sugestao_liberacao_rows`
- View: `vw_cadastros_vinculados`
- Campos de negocio:
  - `bloco_origem`
  - `secao_interna`
  - `categoria_secao`
  - `produto`
  - `litragem`
  - `versao`
  - `abrangencia`

### Investimentos extras
- Bloco no PHP: `investimentos_extras_rows`
- View: `vw_investimentos_extras`
- Campos de negocio:
  - `descricao`
  - `valor_original`
  - `valor_numerico`
  - `unidade`

## Campos que aparecem em quase todas as views

As views por bloco trazem tambem o contexto do formulario:
- `cliente`
- `aba_origem`
- `formulario_index`
- `titulo_plano`
- `numero_acordo`
- `periodo_original`
- `periodo_inicio`
- `periodo_fim`
- `periodo_normalizado`
- `periodo_acoes_original`
- `periodo_acoes_inicio`
- `periodo_acoes_fim`
- `arquivo_nome`
- `processado_em`

## Sugestao de uso no Power BI

### Modelo simples
- carregar `vw_formularios`
- carregar apenas as views dos blocos que interessarem para o painel
- relacionar tudo por `formulario_id`

### Casos comuns
- analise de historico: `vw_historico`
- metas e objetivos: `vw_objetivos`
- investimento: `vw_descricao_investimento` e `vw_investimentos_extras`
- acoes comerciais: `vw_contrapartidas` e `vw_contrapartidas_mensais`
- cadastro e liberacao: `vw_cadastros_vinculados`

## Observacoes

- Algumas strings podem aparecer com perda de acentuacao no console do MySQL, mas isso nao impede o uso no Power BI.
- Quando um bloco tiver mais de um layout entre clientes, a view mantém colunas opcionais e deixa `NULL` no que nao se aplica.
- `bloco_origem` ajuda a separar subtipos dentro da mesma view.