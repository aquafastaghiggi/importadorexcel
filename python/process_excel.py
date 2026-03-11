import sys
import os
import json
import re
import unicodedata
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

try:
    from openpyxl import load_workbook
except ImportError:
    print(json.dumps({
        "success": False,
        "error": "Biblioteca openpyxl não instalada. Execute: python -m pip install openpyxl"
    }, ensure_ascii=False))
    sys.exit(1)


MAIN_BLOCK_ALIASES = {
    "plano_negocios": [
        "PLANO DE NEGOCIOS",
        "PLANO DE INTRODUCAO",
        "JBP",
    ],
    "historico": [
        "HISTORICO",
    ],
    "objetivos": [
        "OBJETIVOS",
    ],
    "descricao_investimento": [
        "DESCRICAO DO INVESTIMENTO",
        "DESCRIÇÃO DO INVESTIMENTO",
    ],
    "contrapartidas_itens_foco": [
        "CONTRAPARTIDAS - ITENS FOCO",
        "CONTRAPARTIDAS ITENS FOCO",
        "ITENS FOCO",
    ],
    "contrapartidas": [
        "CONTRAPARTIDAS",
    ],
    "encartes_obrigatorios": [
        "ENCARTES OBRIGATORIOS",
        "ENCARTES OBRIGATÓRIOS",
        "SUGESTAO DE ENCARTES",
        "SUGESTÃO DE ENCARTES",
    ],
    "cadastros_vinculados": [
        "CADASTROS VINCULADOS",
        "CADASTROS VINCULADOS | LIBERACOES",
        "CADASTROS VINCULADOS | LIBERAÇÕES",
    ],
    "investimentos_extras": [
        "INVESTIMENTOS EXTRAS",
    ],
}

FORM_TITLE_ALIASES = [
    "JBP",
    "PLANO DE NEGOCIOS",
    "PLANO DE INTRODUCAO",
]

HEADER_LABEL_ALIASES = {
    "cliente": [
        "CLIENTE",
        "CLIENTE:",
    ],
    "periodo_plano": [
        "PERIODO DO PLANO",
        "PERÍODO DO PLANO",
        "PERIODO DE PAGAMENTO PLANO",
        "PERÍODO DE PAGAMENTO PLANO",
    ],
    "periodo_acoes_plano": [
        "PERIODO DE ACOES PLANO",
        "PERÍODO DE AÇÕES PLANO",
    ],
    "numero_acordo": [
        "N DO ACORDO",
        "Nº DO ACORDO",
        "NUMERO DO ACORDO",
        "NÚMERO DO ACORDO",
        "N DO ACORDO:",
        "Nº DO ACORDO:",
    ],
}

CONTRAPARTIDAS_HEADERS = {
    "QUANTIDADE": "quantidade",
    "TIPO DE ACAO": "tipo_acao",
    "TIPO DE AÇÃO": "tipo_acao",
    "PERIODO": "periodo",
    "PERÍODO": "periodo",
    "LOJAS": "lojas",
    "OBS": "obs",
}

ITENS_FOCO_HEADERS = {
    "PRODUTO": "produto",
    "VOLUME MENSAL": "volume_mensal",
    "VOLUMEMENSAL": "volume_mensal",
    "VOLUME PERIODO": "volume_periodo",
    "VOLUME PERÍODO": "volume_periodo",
}

CADASTROS_HEADERS = {
    "PRODUTO": "produto",
    "LITRAGEM": "litragem",
    "VERSAO": "versao",
    "VERSÃO": "versao",
    "LOJAS": "abrangencia",
    "ABRANGENCIA": "abrangencia",
    "ABRANGÊNCIA": "abrangencia",
    "PRESENCA ATUAL": "abrangencia",
    "PRESENÇA ATUAL": "abrangencia",
}

MONTH_ALIASES = {
    "JANEIRO": 1,
    "FEVEREIRO": 2,
    "MARCO": 3,
    "MARÇO": 3,
    "ABRIL": 4,
    "MAIO": 5,
    "JUNHO": 6,
    "JULHO": 7,
    "AGOSTO": 8,
    "SETEMBRO": 9,
    "OUTUBRO": 10,
    "NOVEMBRO": 11,
    "DEZEMBRO": 12,
}


def normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"\s+", " ", text)
    return text.upper().strip()


def value_to_str(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def excel_col_letter(col_idx_zero_based):
    col_num = col_idx_zero_based + 1
    result = ""
    while col_num:
        col_num, rem = divmod(col_num - 1, 26)
        result = chr(65 + rem) + result
    return result


def parse_number(text):
    if text is None:
        return None

    raw = str(text).strip()
    if not raw:
        return None

    s = raw.upper()
    s = s.replace("R$", "")
    s = s.replace("%", "")
    s = s.replace("CAIXAS", "")
    s = s.replace("CAIXA", "")
    s = s.replace("CXS", "")
    s = s.replace("CX", "")
    s = s.replace(" ", "")

    if not s:
        return None

    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(".") > 1:
            s = s.replace(".", "")
        elif s.count(".") == 1:
            right = s.split(".")[-1]
            if len(right) == 3:
                s = s.replace(".", "")
        s = s.replace(",", ".")

    try:
        return float(s)
    except Exception:
        return None


def detect_unit(text):
    if not text:
        return None
    raw = str(text).upper()
    norm = normalize_text(text)

    if "R$" in raw:
        return "BRL"
    if "%" in raw:
        return "%"
    if "CAIXAS" in norm or "CAIXA" in norm or "CXS" in norm or "CX" in norm:
        return "CX"
    return None


def month_name_to_number(token):
    return MONTH_ALIASES.get(normalize_text(token))


def split_products_from_cell(text):
    if not text:
        return []

    s = str(text).replace("\r", "\n")
    parts = re.split(r"\n|[|]", s)
    out = []
    for p in parts:
        cleaned = p.strip()
        if cleaned:
            out.append(cleaned)
    return out


def normalize_period(raw_text):
    if not raw_text:
        return {
            "periodo_original": None,
            "periodo_inicio": None,
            "periodo_fim": None,
            "periodo_normalizado": None,
            "periodo_status": "ausente"
        }

    raw = str(raw_text).strip()
    t = normalize_text(raw)

    m = re.match(r"^(\d{1,2})\s*/\s*(\d{2,4})$", raw)
    if m:
        month = int(m.group(1))
        year = int(m.group(2))
        if year < 100:
            year += 2000
        p = f"{year:04d}-{month:02d}"
        return {
            "periodo_original": raw,
            "periodo_inicio": p,
            "periodo_fim": p,
            "periodo_normalizado": p,
            "periodo_status": "normalizado"
        }

    m = re.match(r"^([A-ZÇÃÕÉÊÁÍÓÚ]+)\s*/\s*(\d{2,4})$", t)
    if m:
        month = month_name_to_number(m.group(1))
        year = int(m.group(2))
        if year < 100:
            year += 2000
        if month:
            p = f"{year:04d}-{month:02d}"
            return {
                "periodo_original": raw,
                "periodo_inicio": p,
                "periodo_fim": p,
                "periodo_normalizado": p,
                "periodo_status": "normalizado"
            }

    m = re.match(
        r"^([A-ZÇÃÕÉÊÁÍÓÚ]+)\s*/\s*(\d{2,4})\s*(A|ATE|À)\s*([A-ZÇÃÕÉÊÁÍÓÚ]+)\s*/\s*(\d{2,4})$",
        t
    )
    if m:
        m1 = month_name_to_number(m.group(1))
        y1 = int(m.group(2))
        m2 = month_name_to_number(m.group(4))
        y2 = int(m.group(5))
        if y1 < 100:
            y1 += 2000
        if y2 < 100:
            y2 += 2000
        if m1 and m2:
            start = f"{y1:04d}-{m1:02d}"
            end = f"{y2:04d}-{m2:02d}"
            return {
                "periodo_original": raw,
                "periodo_inicio": start,
                "periodo_fim": end,
                "periodo_normalizado": f"{start} a {end}",
                "periodo_status": "normalizado"
            }

    m = re.match(
        r"^([A-ZÇÃÕÉÊÁÍÓÚ]+)\s*(A|ATE|À)\s*([A-ZÇÃÕÉÊÁÍÓÚ]+)\s*/\s*(\d{2,4})$",
        t
    )
    if m:
        m1 = month_name_to_number(m.group(1))
        m2 = month_name_to_number(m.group(3))
        y = int(m.group(4))
        if y < 100:
            y += 2000
        if m1 and m2:
            start = f"{y:04d}-{m1:02d}"
            end = f"{y:04d}-{m2:02d}"
            return {
                "periodo_original": raw,
                "periodo_inicio": start,
                "periodo_fim": end,
                "periodo_normalizado": f"{start} a {end}",
                "periodo_status": "normalizado"
            }

    return {
        "periodo_original": raw,
        "periodo_inicio": None,
        "periodo_fim": None,
        "periodo_normalizado": raw,
        "periodo_status": "nao_interpretado"
    }


def worksheet_to_grid(ws):
    grid = []
    for row in ws.iter_rows():
        grid.append([value_to_str(cell.value) for cell in row])
    return grid


def non_empty_cells(row):
    return [(i, value_to_str(v)) for i, v in enumerate(row) if value_to_str(v)]


def detect_column_regions(grid, min_non_empty_rows=3, gap_tolerance=2, min_region_width=4):
    if not grid:
        return []

    max_cols = max((len(r) for r in grid), default=0)
    col_scores = []

    for c in range(max_cols):
        count = 0
        for row in grid:
            if c < len(row) and value_to_str(row[c]):
                count += 1
        col_scores.append(count)

    candidate_cols = [i for i, score in enumerate(col_scores) if score >= min_non_empty_rows]
    if not candidate_cols:
        return []

    groups = []
    start = candidate_cols[0]
    prev = candidate_cols[0]

    for col in candidate_cols[1:]:
        if col - prev <= gap_tolerance:
            prev = col
        else:
            groups.append((start, prev))
            start = col
            prev = col
    groups.append((start, prev))

    filtered = []
    for start_col, end_col in groups:
        width = end_col - start_col + 1
        if width >= min_region_width:
            filtered.append({
                "start_col": start_col,
                "end_col": end_col,
                "width": width
            })

    return filtered


def crop_grid_to_region(grid, start_col, end_col):
    cropped = []
    for row in grid:
        if start_col < len(row):
            cropped.append(row[start_col:end_col + 1])
        else:
            cropped.append([])
    return cropped


def find_form_anchors(grid, max_header_rows=20):
    anchors = []

    max_rows = min(len(grid), max_header_rows)
    for r_idx in range(max_rows):
        row = grid[r_idx]
        for c_idx, value in enumerate(row):
            norm = normalize_text(value)
            if not norm:
                continue

            for alias in FORM_TITLE_ALIASES:
                alias_norm = normalize_text(alias)
                if norm == alias_norm or norm.startswith(alias_norm + " |") or norm.startswith(alias_norm + " -"):
                    anchors.append({
                        "row": r_idx,
                        "col": c_idx,
                        "label": value_to_str(value)
                    })
                    break

    anchors.sort(key=lambda x: x["col"])

    deduped = []
    for item in anchors:
        if not deduped:
            deduped.append(item)
            continue

        prev = deduped[-1]
        if abs(prev["col"] - item["col"]) <= 2:
            continue
        deduped.append(item)

    return deduped


def detect_form_regions_from_anchors(grid, anchors, side_margin=2):
    if not anchors:
        return []

    max_cols = max((len(r) for r in grid), default=0)
    regions = []

    for i, anchor in enumerate(anchors):
        if i == 0:
            start_col = max(0, anchor["col"] - side_margin)
        else:
            prev = anchors[i - 1]
            midpoint = (prev["col"] + anchor["col"]) // 2
            start_col = max(0, midpoint)

        if i == len(anchors) - 1:
            end_col = max_cols - 1
        else:
            nxt = anchors[i + 1]
            midpoint = (anchor["col"] + nxt["col"]) // 2
            end_col = max(start_col, midpoint - 1)

        regions.append({
            "start_col": start_col,
            "end_col": end_col,
            "width": end_col - start_col + 1,
            "anchor_col": anchor["col"],
            "anchor_label": anchor["label"]
        })

    return regions


def find_block_headers(grid):
    found = []

    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            raw = value_to_str(value)
            norm = normalize_text(raw)
            if not norm:
                continue

            for block_name, aliases in MAIN_BLOCK_ALIASES.items():
                matched = False
                for alias in aliases:
                    alias_norm = normalize_text(alias)
                    if norm == alias_norm or norm.startswith(alias_norm + " |") or norm.startswith(alias_norm + " -"):
                        found.append({
                            "block": block_name,
                            "row": r_idx,
                            "col": c_idx,
                            "label_found": raw
                        })
                        matched = True
                        break
                if matched:
                    break

    found.sort(key=lambda x: (x["row"], x["col"]))

    deduped = []
    for item in found:
        if not deduped:
            deduped.append(item)
            continue

        prev = deduped[-1]
        if prev["block"] == item["block"] and abs(prev["row"] - item["row"]) <= 1:
            if len(item["label_found"]) > len(prev["label_found"]):
                deduped[-1] = item
        else:
            deduped.append(item)

    return deduped


def is_valid_form_region(headers_found):
    blocks = {h["block"] for h in headers_found}

    if "plano_negocios" not in blocks:
        return False

    meaningful = {
        "plano_negocios",
        "historico",
        "objetivos",
        "descricao_investimento",
        "contrapartidas",
        "contrapartidas_itens_foco",
        "encartes_obrigatorios",
        "cadastros_vinculados",
        "investimentos_extras",
    }

    found_meaningful = blocks.intersection(meaningful)
    return len(found_meaningful) >= 2


def build_row_ranges(headers, total_rows):
    if not headers:
        return []

    ranges = []
    for i, item in enumerate(headers):
        start = item["row"]
        end = total_rows - 1
        if i < len(headers) - 1:
            end = headers[i + 1]["row"] - 1

        ranges.append({
            "block": item["block"],
            "label_found": item["label_found"],
            "start_row": start,
            "end_row": end,
            "start_col": item["col"]
        })
    return ranges


def slice_block_rows(grid, block_info):
    rows = []
    for r_idx in range(block_info["start_row"], block_info["end_row"] + 1):
        if r_idx >= len(grid):
            continue
        row = grid[r_idx]
        filled = non_empty_cells(row)
        if filled:
            rows.append({
                "row_excel": r_idx + 1,
                "cells": [
                    {
                        "col_idx": idx,
                        "col_excel": excel_col_letter(idx),
                        "cell_ref": f"{excel_col_letter(idx)}{r_idx + 1}",
                        "value": val
                    }
                    for idx, val in filled
                ]
            })
    return rows


def row_to_joined(row_obj):
    return " | ".join(cell["value"] for cell in row_obj["cells"])


def build_context(header, sheet_name, form_index):
    return {
        "cliente": header.get("cliente"),
        "periodo_original": header.get("periodo_original"),
        "periodo_inicio": header.get("periodo_inicio"),
        "periodo_fim": header.get("periodo_fim"),
        "periodo_normalizado": header.get("periodo_normalizado"),
        "periodo_status": header.get("periodo_status"),
        "periodo_acoes_original": header.get("periodo_acoes_original"),
        "periodo_acoes_inicio": header.get("periodo_acoes_inicio"),
        "periodo_acoes_fim": header.get("periodo_acoes_fim"),
        "periodo_acoes_normalizado": header.get("periodo_acoes_normalizado"),
        "periodo_acoes_status": header.get("periodo_acoes_status"),
        "numero_acordo": header.get("numero_acordo"),
        "titulo_plano": header.get("titulo_plano"),
        "aba_origem": sheet_name,
        "formulario_index": form_index
    }


def find_label_in_row(cells, aliases):
    aliases_norm = {normalize_text(a).rstrip(":") for a in aliases}
    for pos, cell in enumerate(cells):
        norm = normalize_text(cell["value"]).rstrip(":")
        if norm in aliases_norm:
            return pos
    return None


def value_after_label(cells, label_pos):
    if label_pos is None:
        return None
    if label_pos + 1 < len(cells):
        return cells[label_pos + 1]["value"]
    return None


def extract_header_from_block(rows_raw, label_found):
    header = {
        "cliente": None,
        "periodo_original": None,
        "periodo_inicio": None,
        "periodo_fim": None,
        "periodo_normalizado": None,
        "periodo_status": None,
        "periodo_acoes_original": None,
        "periodo_acoes_inicio": None,
        "periodo_acoes_fim": None,
        "periodo_acoes_normalizado": None,
        "periodo_acoes_status": None,
        "numero_acordo": None,
        "titulo_plano": label_found
    }

    for row in rows_raw:
        cells = row["cells"]

        pos_cliente = find_label_in_row(cells, HEADER_LABEL_ALIASES["cliente"])
        if pos_cliente is not None and not header["cliente"]:
            header["cliente"] = value_after_label(cells, pos_cliente)

        pos_periodo = find_label_in_row(cells, HEADER_LABEL_ALIASES["periodo_plano"])
        if pos_periodo is not None and not header["periodo_original"]:
            v = value_after_label(cells, pos_periodo)
            p = normalize_period(v)
            header["periodo_original"] = p["periodo_original"]
            header["periodo_inicio"] = p["periodo_inicio"]
            header["periodo_fim"] = p["periodo_fim"]
            header["periodo_normalizado"] = p["periodo_normalizado"]
            header["periodo_status"] = p["periodo_status"]

        pos_periodo_acoes = find_label_in_row(cells, HEADER_LABEL_ALIASES["periodo_acoes_plano"])
        if pos_periodo_acoes is not None and not header["periodo_acoes_original"]:
            v = value_after_label(cells, pos_periodo_acoes)
            p = normalize_period(v)
            header["periodo_acoes_original"] = p["periodo_original"]
            header["periodo_acoes_inicio"] = p["periodo_inicio"]
            header["periodo_acoes_fim"] = p["periodo_fim"]
            header["periodo_acoes_normalizado"] = p["periodo_normalizado"]
            header["periodo_acoes_status"] = p["periodo_status"]

        pos_acordo = find_label_in_row(cells, HEADER_LABEL_ALIASES["numero_acordo"])
        if pos_acordo is not None and not header["numero_acordo"]:
            header["numero_acordo"] = value_after_label(cells, pos_acordo)

    return header


def extract_year_from_title(title):
    if not title:
        return None
    m = re.search(r"(20\d{2}|\b\d{2}\b)", normalize_text(title))
    if not m:
        return None
    y = int(m.group(1))
    if y < 100:
        y += 2000
    return y


def parse_kv_list(rows_raw, header, sheet_name, form_index, tipo_registro, title_year=None):
    rows = []
    ctx = build_context(header, sheet_name, form_index)

    for order, row in enumerate(rows_raw, start=1):
        cells = row["cells"]
        joined = row_to_joined(row)
        if len(cells) < 2:
            continue

        label = cells[0]["value"]
        value = cells[-1]["value"]

        norm_label = normalize_text(label)
        if norm_label in {
            "OBJETIVOS",
            "DESCRICAO DO INVESTIMENTO",
            "DESCRIÇÃO DO INVESTIMENTO",
            "HISTORICO",
            "INVESTIMENTOS EXTRAS",
        }:
            continue

        rows.append({
            **ctx,
            "tipo_registro": tipo_registro,
            "descricao": label,
            "valor_original": value,
            "valor_numerico": parse_number(value),
            "unidade": detect_unit(f"{label} {value}"),
            "ano_bloco": title_year,
            "linha_ordem": order,
            "row_excel": row["row_excel"],
            "linha_original": joined
        })

    return rows


def detect_table_header_row(rows_raw, expected_headers):
    expected_norm = set(normalize_text(x) for x in expected_headers.keys())
    best = None
    best_score = -1

    for idx, row in enumerate(rows_raw):
        row_norms = [normalize_text(c["value"]) for c in row["cells"]]
        score = sum(1 for v in row_norms if v in expected_norm)
        if score > best_score:
            best_score = score
            best = idx

    if best_score >= 2:
        return best
    return None


def build_column_map(header_cells, header_dict):
    col_map = {}
    for cell in header_cells:
        norm = normalize_text(cell["value"])
        if norm in header_dict:
            col_map[cell["col_idx"]] = header_dict[norm]
    return col_map


def parse_grid_table(rows_raw, header, sheet_name, form_index, tipo_registro, header_dict, section_rules=False):
    rows = []
    ctx = build_context(header, sheet_name, form_index)
    header_row_idx = detect_table_header_row(rows_raw, header_dict)
    if header_row_idx is None:
        return rows

    table_header_cells = rows_raw[header_row_idx]["cells"]
    col_map = build_column_map(table_header_cells, header_dict)

    current_section = None
    line_order = 0
    normalized_header_labels = set(normalize_text(k) for k in header_dict.keys())

    for idx in range(header_row_idx + 1, len(rows_raw)):
        row = rows_raw[idx]
        values = [c["value"] for c in row["cells"]]
        norms = [normalize_text(v) for v in values]

        if section_rules:
            if len(values) == 1 and normalize_text(values[0]) not in normalized_header_labels:
                current_section = values[0]
                continue

        if sum(1 for n in norms if n in normalized_header_labels) >= 2:
            col_map = build_column_map(row["cells"], header_dict)
            continue

        line_data = {
            **ctx,
            "tipo_registro": tipo_registro,
            "linha_ordem": line_order + 1,
            "row_excel": row["row_excel"],
            "linha_original": row_to_joined(row)
        }

        if current_section:
            line_data["secao_interna"] = current_section

        filled_count = 0
        for cell in row["cells"]:
            col_idx = cell["col_idx"]
            if col_idx in col_map:
                field = col_map[col_idx]
                line_data[field] = cell["value"]
                filled_count += 1

        if filled_count >= 1:
            if "quantidade" in line_data:
                line_data["quantidade_numerica"] = parse_number(line_data.get("quantidade"))
            if "volume_mensal" in line_data:
                line_data["volume_mensal_numerica"] = parse_number(line_data.get("volume_mensal"))
            if "volume_periodo" in line_data:
                line_data["volume_periodo_numerica"] = parse_number(line_data.get("volume_periodo"))
            line_order += 1
            line_data["linha_ordem"] = line_order
            rows.append(line_data)

    return rows


def find_month_header_sections(rows_raw):
    sections = []

    for idx, row in enumerate(rows_raw):
        months = []
        for cell in row["cells"]:
            norm = normalize_text(cell["value"])
            if norm in MONTH_ALIASES:
                months.append((cell["col_idx"], cell["value"], MONTH_ALIASES[norm]))

        if len(months) >= 3:
            sections.append({
                "header_idx": idx,
                "month_cols": months
            })

    return sections


def parse_month_grid(rows_raw, header, sheet_name, form_index):
    rows = []
    ctx = build_context(header, sheet_name, form_index)

    sections = find_month_header_sections(rows_raw)
    if not sections:
        return rows

    global_line_visual = 0

    for s_idx, section in enumerate(sections):
        start_idx = section["header_idx"] + 1
        end_idx = len(rows_raw) - 1
        if s_idx < len(sections) - 1:
            end_idx = sections[s_idx + 1]["header_idx"] - 1

        month_cols = section["month_cols"]

        for idx in range(start_idx, end_idx + 1):
            row = rows_raw[idx]
            cell_by_col = {c["col_idx"]: c["value"] for c in row["cells"]}

            added_any = False
            for col_idx, month_label, month_num in month_cols:
                cell_value = cell_by_col.get(col_idx)
                if not cell_value:
                    continue

                products = split_products_from_cell(cell_value)
                for prod in products:
                    rows.append({
                        **ctx,
                        "tipo_registro": "encarte_obrigatorio",
                        "mes": month_label,
                        "mes_numero": month_num,
                        "produto": prod,
                        "linha_visual": global_line_visual + 1,
                        "row_excel": row["row_excel"],
                        "linha_original": row_to_joined(row)
                    })
                    added_any = True

            if added_any:
                global_line_visual += 1

    return rows


def split_contrapartidas_and_encartes(rows_raw):
    split_idx = None
    split_labels = {
        "SUGESTAO DE ENCARTES",
        "SUGESTÃO DE ENCARTES",
        "ENCARTES OBRIGATORIOS",
        "ENCARTES OBRIGATÓRIOS",
    }

    norm_split_labels = {normalize_text(x) for x in split_labels}

    for idx, row in enumerate(rows_raw):
        joined = normalize_text(row_to_joined(row))
        values = [normalize_text(c["value"]) for c in row["cells"]]

        for label in norm_split_labels:
            if joined == label or joined.startswith(label):
                split_idx = idx
                break

        if split_idx is not None:
            break

        if len(values) == 1 and values[0] in norm_split_labels:
            split_idx = idx
            break

    if split_idx is None:
        return rows_raw, []

    contrapartidas_rows = rows_raw[:split_idx]
    encartes_rows = rows_raw[split_idx + 1:]
    return contrapartidas_rows, encartes_rows


def process_form_grid(grid, sheet_name, form_index, region_meta):
    headers_found = find_block_headers(grid)
    if not is_valid_form_region(headers_found):
        return None

    ranges = build_row_ranges(headers_found, len(grid))

    result = {
        "sheet_name": sheet_name,
        "formulario_index": form_index,
        "region_start_col": region_meta["start_col"] + 1,
        "region_end_col": region_meta["end_col"] + 1,
        "region_width": region_meta["width"],
        "identified_blocks": [h["block"] for h in headers_found],
        "header": {
            "cliente": None,
            "periodo_original": None,
            "periodo_inicio": None,
            "periodo_fim": None,
            "periodo_normalizado": None,
            "periodo_status": None,
            "periodo_acoes_original": None,
            "periodo_acoes_inicio": None,
            "periodo_acoes_fim": None,
            "periodo_acoes_normalizado": None,
            "periodo_acoes_status": None,
            "numero_acordo": None,
            "titulo_plano": None
        },
        "plano_negocios_rows": [],
        "historico_rows": [],
        "objetivos_rows": [],
        "descricao_investimento_rows": [],
        "contrapartidas_rows": [],
        "contrapartidas_itens_foco_rows": [],
        "encartes_obrigatorios_rows": [],
        "cadastros_vinculados_rows": [],
        "investimentos_extras_rows": [],
        "raw_blocks": []
    }

    for block_info in ranges:
        rows_raw = slice_block_rows(grid, block_info)
        title_year = extract_year_from_title(block_info["label_found"])

        result["raw_blocks"].append({
            "block": block_info["block"],
            "label_found": block_info["label_found"],
            "start_row": block_info["start_row"] + 1,
            "end_row": block_info["end_row"] + 1,
            "total_rows_raw": len(rows_raw)
        })

        if block_info["block"] == "plano_negocios":
            result["header"] = extract_header_from_block(rows_raw, block_info["label_found"])
            result["plano_negocios_rows"] = [build_context(result["header"], sheet_name, form_index)]

        elif block_info["block"] == "historico":
            result["historico_rows"].extend(
                parse_kv_list(rows_raw, result["header"], sheet_name, form_index, "historico", title_year)
            )

        elif block_info["block"] == "objetivos":
            result["objetivos_rows"].extend(
                parse_kv_list(rows_raw, result["header"], sheet_name, form_index, "objetivo", title_year)
            )

        elif block_info["block"] == "descricao_investimento":
            result["descricao_investimento_rows"].extend(
                parse_kv_list(rows_raw, result["header"], sheet_name, form_index, "descricao_investimento", None)
            )

        elif block_info["block"] == "contrapartidas":
            contrapartidas_only_rows, encartes_embedded_rows = split_contrapartidas_and_encartes(rows_raw)

            result["contrapartidas_rows"].extend(
                parse_grid_table(
                    contrapartidas_only_rows,
                    result["header"],
                    sheet_name,
                    form_index,
                    "contrapartida",
                    CONTRAPARTIDAS_HEADERS
                )
            )

            if encartes_embedded_rows:
                result["encartes_obrigatorios_rows"].extend(
                    parse_month_grid(
                        encartes_embedded_rows,
                        result["header"],
                        sheet_name,
                        form_index
                    )
                )

                if "encartes_obrigatorios" not in result["identified_blocks"]:
                    result["identified_blocks"].append("encartes_obrigatorios")

        elif block_info["block"] == "contrapartidas_itens_foco":
            result["contrapartidas_itens_foco_rows"].extend(
                parse_grid_table(
                    rows_raw,
                    result["header"],
                    sheet_name,
                    form_index,
                    "contrapartida_item_foco",
                    ITENS_FOCO_HEADERS
                )
            )

        elif block_info["block"] == "encartes_obrigatorios":
            result["encartes_obrigatorios_rows"].extend(
                parse_month_grid(rows_raw, result["header"], sheet_name, form_index)
            )

        elif block_info["block"] == "cadastros_vinculados":
            result["cadastros_vinculados_rows"].extend(
                parse_grid_table(
                    rows_raw,
                    result["header"],
                    sheet_name,
                    form_index,
                    "cadastro_vinculado",
                    CADASTROS_HEADERS,
                    section_rules=True
                )
            )

        elif block_info["block"] == "investimentos_extras":
            result["investimentos_extras_rows"].extend(
                parse_kv_list(rows_raw, result["header"], sheet_name, form_index, "investimento_extra", None)
            )

    if not result["plano_negocios_rows"]:
        result["plano_negocios_rows"] = [build_context(result["header"], sheet_name, form_index)]

    result["identified_blocks"] = list(dict.fromkeys(result["identified_blocks"]))
    return result


def process_sheet(ws):
    full_grid = worksheet_to_grid(ws)

    anchors = find_form_anchors(full_grid)
    if len(anchors) >= 1:
        regions = detect_form_regions_from_anchors(full_grid, anchors)
    else:
        regions = detect_column_regions(full_grid)

    forms = []
    ignored_regions = []

    form_index = 1
    for region in regions:
        region_grid = crop_grid_to_region(full_grid, region["start_col"], region["end_col"])
        headers_found = find_block_headers(region_grid)

        if is_valid_form_region(headers_found):
            processed = process_form_grid(region_grid, ws.title, form_index, region)
            if processed:
                forms.append(processed)
                form_index += 1
        else:
            ignored_regions.append({
                "start_col": region["start_col"] + 1,
                "end_col": region["end_col"] + 1,
                "width": region["width"],
                "motivo": "regiao_sem_blocos_principais"
            })

    return {
        "sheet_name": ws.title,
        "forms": forms,
        "ignored_regions": ignored_regions
    }


def main():
    try:
        if len(sys.argv) < 2:
            raise Exception("Caminho do arquivo não informado.")

        file_path = sys.argv[1]
        if not os.path.exists(file_path):
            raise Exception("Arquivo não encontrado.")

        wb = load_workbook(file_path, data_only=True)

        output = {
            "success": True,
            "file_name": os.path.basename(file_path),
            "processed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_sheets": len(wb.sheetnames),
            "sheets": []
        }

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            output["sheets"].append(process_sheet(ws))

        print(json.dumps(output, ensure_ascii=False))

    except Exception as e:
        print(json.dumps({
            "success": False,
            "error": str(e)
        }, ensure_ascii=False))


if __name__ == "__main__":
    main()