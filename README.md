# Importador de Excel JBP

Aplicacao web em PHP que recebe planilhas Excel de acordos/JBP, delega a leitura para um parser Python e apresenta o resultado em uma interface HTML, alem de salvar o JSON tecnico gerado em disco.

## Visao geral

O projeto trabalha em duas camadas:

- `index.php` recebe o upload, valida a extensao, salva o arquivo em `uploads/`, executa o parser Python e renderiza os dados processados.
- `python/process_excel.py` abre a planilha com `openpyxl`, identifica formularios e blocos de negocio, normaliza o conteudo e devolve um JSON no `stdout`.

Quando o processamento termina com sucesso, o PHP tambem grava uma copia formatada do JSON em `resultados/`.

## Fluxo de processamento

1. O usuario envia um arquivo Excel pela interface.
2. O PHP aceita apenas extensoes `.xlsx`, `.xlsm`, `.xltx` e `.xltm`.
3. O arquivo e salvo com timestamp em `uploads/`.
4. O PHP executa `python/process_excel.py <arquivo>`.
5. O parser le todas as abas, identifica regioes com formularios e extrai os blocos reconhecidos.
6. O JSON retornado e exibido na tela e salvo em `resultados/`.

## Estrutura do projeto

- `index.php`: interface principal, upload, chamada do parser e renderizacao dos resumos e tabelas.
- `index-ok.php`: variante da interface principal usada como referencia alternativa.
- `python/process_excel.py`: parser principal em producao.
- `python/process_excel-old.py`: versao anterior do parser.
- `uploads/`: arquivos Excel enviados pela interface.
- `resultados/`: JSONs gerados apos o processamento.

## O que o parser identifica

O parser usa normalizacao de texto, aliases e regras de segmentacao para localizar formularios e blocos mesmo com pequenas variacoes de escrita. Entre os blocos tratados estao:

- plano de negocios / JBP / plano de introducao
- historico
- objetivos
- descricao do investimento
- contrapartidas
- contrapartidas - itens foco
- encartes obrigatorios / sugestao de encartes
- cadastros vinculados
- situacao da liberacao
- investimentos extras

Ele tambem extrai campos de cabecalho como cliente, numero do acordo, periodo do plano, periodo de acoes e titulo do plano.

## Formato da saida

Em alto nivel, o JSON retornado possui:

- metadados do processamento, como `success`, `file_name`, `processed_at` e `total_sheets`
- lista de `sheets`
- em cada aba, uma lista de `forms`
- em cada formulario, `header`, `identified_blocks`, tabelas por bloco e `raw_blocks`
- regioes descartadas em `ignored_regions` quando a aba nao contem blocos principais validos

Os blocos extraidos sao retornados em colecoes como:

- `plano_negocios_rows`
- `historico_rows`
- `objetivos_rows`
- `descricao_investimento_rows`
- `contrapartidas_rows`
- `contrapartidas_itens_foco_rows`
- `encartes_obrigatorios_rows`
- `cadastros_vinculados_rows`
- `situacao_liberacao_rows`
- `investimentos_extras_rows`

## Interface web

A interface PHP nao apenas mostra o JSON bruto. Ela gera uma leitura mais amigavel com:

- resumo consolidado por formulario
- tabelas detalhadas por bloco
- formatacao de percentuais
- blocos detectados e regioes ignoradas
- gravacao do JSON processado para auditoria e analise posterior

## Requisitos

- PHP com `shell_exec` habilitado
- Python 3
- biblioteca `openpyxl`

Instalacao da dependencia Python:

```bash
python -m pip install openpyxl
```

## Como executar

### Opcao 1: servidor embutido do PHP

```bash
php -S 0.0.0.0:8000
```

Depois acesse [http://localhost:8000](http://localhost:8000).

### Opcao 2: ambiente local com XAMPP

Como o projeto esta em `C:\xampp\htdocs\projeto`, ele tambem pode ser aberto pelo Apache do XAMPP em uma URL como:

`http://localhost/projeto`

## Observacoes tecnicas

- O parser percorre todas as abas do workbook.
- Formularios podem ser detectados por ancora de titulo ou por segmentacao de colunas.
- Campos vazios sao removidos da saida com `remove_empty_fields`, reduzindo ruido no JSON final.
- Parte da consolidacao visual e dos resumos exibidos acontece no PHP, nao apenas no parser Python.
- Se o Python falhar ou retornar JSON invalido, a interface exibe a mensagem de erro recebida.

## Validacao rapida

Comandos uteis para conferir a sintaxe dos arquivos principais:

```bash
php -l index.php
php -l index-ok.php
python -m py_compile python/process_excel.py
```
