"""
Exporta os dados de SST do FeatureServer FALCAO_SST_BI (ArcGIS) e gera
data/falcao_sst.json, consumido pelo dashboard estático (index.html).

Credenciais: lidas exclusivamente das variáveis de ambiente
ARCGIS_USERNAME e ARCGIS_PASSWORD (nunca de arquivo). No GitHub Actions,
vêm de Settings > Secrets and variables > Actions.
"""

import json
import os
import sys
from datetime import datetime, timezone

import pandas as pd
import requests

# =====================================================================
# CONFIGURAÇÃO
# =====================================================================

USERNAME = os.environ.get("ARCGIS_USERNAME")
PASSWORD = os.environ.get("ARCGIS_PASSWORD")

if not USERNAME or not PASSWORD:
    sys.exit(
        "Erro: defina as variáveis de ambiente ARCGIS_USERNAME e "
        "ARCGIS_PASSWORD antes de rodar este script."
    )

BASE_URL = "https://services6.arcgis.com/QJAp6nG4ishOkuMg/arcgis/rest/services/FALCAO_SST_BI/FeatureServer"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "..", "data", "falcao_sst.json")


# =====================================================================
# AUTENTICAÇÃO
# =====================================================================

def get_token():
    resp = requests.post(
        "https://www.arcgis.com/sharing/rest/generateToken",
        data={
            "username": USERNAME,
            "password": PASSWORD,
            "referer": "https://www.arcgis.com",
            "f": "json",
            "expiration": 1440,
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if "token" not in data:
        raise Exception(f"Falha ao gerar token ArcGIS: {data}")
    return data["token"]


# =====================================================================
# BUSCA DE DADOS (com paginação — Max Record Count do serviço = 1000)
# =====================================================================

def get_layer_data(token, layer_index, out_fields, page_size=1000):
    url = f"{BASE_URL}/{layer_index}/query"
    all_features = []
    offset = 0
    while True:
        params = {
            "where": "1=1",
            "outFields": out_fields,
            "f": "json",
            "token": token,
            "resultRecordCount": page_size,
            "resultOffset": offset,
        }
        r = requests.get(url, params=params, timeout=60)
        r.raise_for_status()
        data = r.json()
        if "error" in data:
            raise Exception(
                f"Erro do ArcGIS na camada {layer_index} (campos: {out_fields}): {data['error']}"
            )
        feats = data.get("features", [])
        if not feats:
            break
        all_features.extend(feats)
        if len(feats) < page_size:
            break
        offset += page_size

    records = [f["attributes"] for f in all_features]
    df = pd.DataFrame.from_records(records)
    print(f"Camada {layer_index}: {len(df)} registros")
    return df


# =====================================================================
# TABELA DE LABELS (equivalente a labelChoice() do Arcade)
# =====================================================================

LABELS = {
    "acidente": "Acidente",
    "desvio": "Desvio",
    "incidente": "Incidente",
    "fatalidade": "Fatalidade",
    "inspecao": "Inspeção SST",
    "afastamento": "Com afastamento",
    "sem_afastamento": "Sem afastamento",
    "aso": "Atestado de Saúde Ocupacional – ASO",
    "epi": "Ficha de Equipamento de Proteção Individual – EPI",
    "integracao": "Integração SST",
    "os": "Ordem de Serviço – OS",
    "treinamentos_nrs": "Treinamentos das NRs",
    "APR": "Análise Preliminar de Risco – APR",
    "ART": "Anotação de Responsabilidade Técnica – ART",
    "AVCB": "Auto de Vistoria do Corpo de Bombeiros – AVCB",
    "checK_epi_equipamento": "Check-list – Ferramentas e Equipamentos",
    "checK_veiculo": "Check-list – Veículo",
    "CIPA": "Comissão Interna de Prevenção de Acidentes e Assédios – CIPA",
    "cx_primeiro_socorros": "Caixa de Primeiros Socorros",
    "extintor": "Extintor de Incêndio",
    "FDS": "Ficha de Dados de Segurança – FDS",
    "GRO": "Gerenciamento de Risco Ocupacional – GRO",
    "inventario_maquina_epi": "Inventário de Máquinas e Equipamentos",
    "laud_iluminacao": "Laudo de Iluminação",
    "laudo_agua_potavel": "Laudo de Potabilidade de Água",
    "laudo_eletrico": "Laudo Elétrico",
    "laudo_ergo": "Laudo Ergonômico",
    "laudo_insalubridade": "Laudo de Insalubridade",
    "laudo_periculosidade": "Laudo de Periculosidade",
    "laudo_ruido": "Laudo de Ruído",
    "layout": "Layout do Canteiro de Obras",
    "LTCAT": "Laudo Técnico de Condições Ambientais de Trabalho – LTCAT",
    "mapa_risco": "Mapa de Risco",
    "PAE": "Plano de Atendimento e Emergência – PAE",
    "PCA": "Programa de Conservação Auditiva – PCA",
    "PCMSO": "Programa de Controle Médico e Saúde Ocupacional – PCMSO",
    "PET": "Permissão de Entrada e Trabalho – PET",
    "PGR": "Programa de Gerenciamento de Risco – PGR",
    "PPR": "Programa de Proteção Respiratória – PPR",
    "procedimentos_valas": "Procedimentos para Escavações e Valas",
    "PT": "Permissão de Trabalho – PT",
    "SESMT": "Registro do SESMT",
    "dds": "Diálogo Semanal de Segurança – DDS",
    "campanha_sst": "Campanha SST",
    "saude_mental": "Saúde Mental",
}


def label(code):
    if code is None:
        return ""
    c = str(code).strip()
    if not c:
        return ""
    return LABELS.get(c, c)


def unique_labels(values, texto_padrao):
    itens = []
    for v in values:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue
        for parte in str(v).split(","):
            p = parte.strip()
            if p:
                itens.append(label(p))
    vistos = []
    for it in itens:
        if it and it not in vistos:
            vistos.append(it)
    return " • ".join(vistos) if vistos else texto_padrao


def _clean_text(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s else None


SITUACAO_LABEL = {"tipico": "Típico", "trajeto": "Trajeto"}


def _detalhe_acidentes(g):
    linhas = []
    acidentes = g[g["isDeflator"]].sort_values("data_registro")
    for _, row in acidentes.iterrows():
        situacao = _clean_text(row.get("situacao_ocorrencia"))
        linhas.append({
            "data": row["data_registro"].strftime("%Y-%m-%d") if pd.notna(row["data_registro"]) else None,
            "tipo_label": row["acidenteLabel"],
            "situacao": SITUACAO_LABEL.get(situacao, situacao) if situacao else None,
            "cid": _clean_text(row.get("cid")),
            "n_cat": _clean_text(row.get("n_cat")),
            "dias_afastado": int(row["dias_afastado"]) if pd.notna(row.get("dias_afastado")) else 0,
            "custo_afastamento": float(row["custo_afastamento"]) if pd.notna(row.get("custo_afastamento")) else 0.0,
            "houve_tratativa": _clean_text(row.get("houve_tratativa")),
            "descricao_tratativa": _clean_text(row.get("descricao_tratativas")),
        })
    return linhas


# =====================================================================
# FUNÇÕES DE PERÍODO (equivalentes ao Arcade — Month() zero-indexado)
# =====================================================================

MESES_LABEL = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]


def mes_ano_tecnico(dt):
    return f"{dt.year}-{dt.month - 1:02d}"


def mes_ano_label(dt):
    return f"{MESES_LABEL[dt.month - 1]}/{dt.year}"


# =====================================================================
# AGRUPAMENTOS POR PC + MÊS
# =====================================================================

def agrupar_docs(df):
    linhas = []
    for (id_pc, mes_ano), g in df.groupby(["id_pc", "mes_ano"]):
        linhas.append({
            "id_pc": id_pc,
            "mes_ano": mes_ano,
            "diretor": g["diretor"].iloc[0],
            "gerente": g["gerente"].iloc[0],
            "projeto": g["projeto"].iloc[0],
            "mes_ano_label": g["mes_ano_label"].iloc[0],
            "soma_doc": g["nota_analise_documental"].fillna(0).astype(float).sum(),
            "qtd_doc": len(g),
            "documentos_reprovados": unique_labels(
                list(g["lista_nao_conforme"]) + list(g["lista_nao_conforme_g2"]),
                "Nenhum documento reprovado no período",
            ),
        })
    return pd.DataFrame(linhas)


def agrupar_eventos(df):
    linhas = []
    for (id_pc, mes_ano), g in df.groupby(["id_pc", "mes_ano"]):
        qual = g[g["qualifica"]]
        linhas.append({
            "id_pc": id_pc,
            "mes_ano": mes_ano,
            "diretor": g["diretor"].iloc[0],
            "gerente": g["gerente"].iloc[0],
            "projeto": g["projeto"].iloc[0],
            "mes_ano_label": g["mes_ano_label"].iloc[0],
            "soma_evt": qual["notaEventoNum"].sum(),
            "qtd_evt": len(qual),
            "qtd_dds": int((qual["tipoEvento"] == "dds").sum()),
            "qtd_campanha_sst": int((qual["tipoEvento"] == "campanha_sst").sum()),
            "qtd_saude_mental": int((qual["tipoEvento"] == "saude_mental").sum()),
            "eventos_realizados": unique_labels(list(qual["tipo_evento"]), "Nenhum evento registrado no período"),
        })
    return pd.DataFrame(linhas)


def agrupar_ocorrencias(df):
    linhas = []
    for (id_pc, mes_ano), g in df.groupby(["id_pc", "mes_ano"]):
        campo = g["notaCampoRow"].dropna()
        sst = g["notaSST"].dropna()
        linhas.append({
            "id_pc": id_pc,
            "mes_ano": mes_ano,
            "diretor": g["diretor"].iloc[0],
            "gerente": g["gerente"].iloc[0],
            "projeto": g["projeto"].iloc[0],
            "mes_ano_label": g["mes_ano_label"].iloc[0],
            "qtd_registros_ocorrencias_total": len(g),
            "soma_ocor": campo.sum(),
            "qtd_ocor": campo.shape[0],
            "soma_def": g["deflatorValor"].sum(),
            "qtd_def": int(g["isDeflator"].sum()),
            "qtd_acidentes": int((g["tipo"] == "acidente").sum()),
            "qtd_fatalidades": int((g["tipo"] == "fatalidade").sum()),
            "qtd_acidentes_afastamento": int(((g["tipo"] == "acidente") & (g["statusAcidente"] == "afastamento")).sum()),
            "qtd_acidentes_sem_afastamento": int(((g["tipo"] == "acidente") & (g["statusAcidente"] == "sem_afastamento")).sum()),
            "soma_analise_sst": sst.sum(),
            "qtd_analise_sst": sst.shape[0],
            "resumo_acidentes": unique_labels(list(g["acidenteLabel"].dropna()), "Nenhum acidente registrado no período"),
            "soma_dias_afastado": int(g.loc[g["isDeflator"], "dias_afastado"].fillna(0).sum()),
            "soma_custo_afastamento": float(g.loc[g["isDeflator"], "custo_afastamento"].fillna(0).sum()),
            "acidentes_detalhe": _detalhe_acidentes(g),
        })
    return pd.DataFrame(linhas)


# =====================================================================
# SERIALIZAÇÃO JSON (limpa NaN / tipos numpy)
# =====================================================================

def clean_value(v):
    if isinstance(v, float) and pd.isna(v):
        return None
    if pd.isna(v) if not isinstance(v, (list, dict)) else False:
        return None
    try:
        import numpy as np
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            return None if pd.isna(v) else float(v)
        if isinstance(v, np.bool_):
            return bool(v)
    except ImportError:
        pass
    return v


def clean_records(records):
    return [{k: clean_value(v) for k, v in row.items()} for row in records]


def main():
    token = get_token()

    pcs = get_layer_data(token, 0, "globalid,id_pc,diretor,gerente,projeto")

    # Limpa espacos extras/duplicados nos campos de texto do ArcGIS (ex:
    # "  PC 367 -  ..." com espacos irregulares). Isso NAO corrige grafias
    # divergentes do mesmo nome (ex: "Cintia" vs "Cíntia") — decidir a
    # grafia certa exige revisao humana dos registros no ArcGIS.
    for _col in ("diretor", "gerente", "projeto"):
        pcs[_col] = pcs[_col].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    ocorrencias = get_layer_data(
        token, 1,
        "parentglobalid,data_registro,tipo_ocorrencia,status_acidente,nota_ocorrencia,nota_analise_sst,"
        "dias_afastado,custo_afastamento,cid,n_cat,situacao_ocorrencia,houve_tratativa,descricao_tratativas"
    )
    docs = get_layer_data(
        token, 2, "parentglobalid,data_analise,nota_analise_documental,lista_nao_conforme,lista_nao_conforme_g2"
    )
    eventos = get_layer_data(
        token, 3, "parentglobalid,data_evento,tipo_evento,nota_evento,total_colaboradores_presente"
    )

    # ---- normalização de GUIDs e datas ----
    pcs["globalid"] = pcs["globalid"].astype(str).str.strip().str.upper()

    ocorrencias = ocorrencias.dropna(subset=["parentglobalid", "data_registro"]).copy()
    docs = docs.dropna(subset=["parentglobalid", "data_analise"]).copy()
    eventos = eventos.dropna(subset=["parentglobalid", "data_evento"]).copy()

    for _df in (ocorrencias, docs, eventos):
        _df["parentglobalid"] = _df["parentglobalid"].astype(str).str.strip().str.upper()

    ocorrencias["data_registro"] = pd.to_datetime(ocorrencias["data_registro"], unit="ms", utc=True)
    docs["data_analise"] = pd.to_datetime(docs["data_analise"], unit="ms", utc=True)
    eventos["data_evento"] = pd.to_datetime(eventos["data_evento"], unit="ms", utc=True)

    # ---- junção com PCs (descarta registros sem contrato correspondente) ----
    ocorrencias = ocorrencias.merge(pcs, left_on="parentglobalid", right_on="globalid", how="inner")
    docs = docs.merge(pcs, left_on="parentglobalid", right_on="globalid", how="inner")
    eventos = eventos.merge(pcs, left_on="parentglobalid", right_on="globalid", how="inner")

    ocorrencias["mes_ano"] = ocorrencias["data_registro"].apply(mes_ano_tecnico)
    ocorrencias["mes_ano_label"] = ocorrencias["data_registro"].apply(mes_ano_label)
    docs["mes_ano"] = docs["data_analise"].apply(mes_ano_tecnico)
    docs["mes_ano_label"] = docs["data_analise"].apply(mes_ano_label)
    eventos["mes_ano"] = eventos["data_evento"].apply(mes_ano_tecnico)
    eventos["mes_ano_label"] = eventos["data_evento"].apply(mes_ano_label)

    # ---- colunas auxiliares — ocorrências ----
    ocorrencias["tipo"] = ocorrencias["tipo_ocorrencia"].fillna("").astype(str).str.strip().str.lower()
    ocorrencias["nota"] = ocorrencias["nota_ocorrencia"].fillna(1000).astype(float)
    ocorrencias["possuiAnaliseSST"] = ocorrencias["nota_analise_sst"].notna()
    ocorrencias["notaSST"] = ocorrencias.apply(
        lambda r: min(max(float(r["nota_analise_sst"]), 0), 100) if r["possuiAnaliseSST"] else None, axis=1
    )
    ocorrencias["statusAcidente"] = ocorrencias["status_acidente"].fillna("").astype(str).str.strip().str.lower()

    def _nota_campo_row(r):
        if r["possuiAnaliseSST"] and r["tipo"] not in ("acidente", "fatalidade"):
            return min(max(r["notaSST"] * 10, 0), 1000)
        elif r["tipo"] in ("desvio", "incidente") and not r["possuiAnaliseSST"]:
            return r["nota"]
        return None

    ocorrencias["notaCampoRow"] = ocorrencias.apply(_nota_campo_row, axis=1)
    ocorrencias["isDeflator"] = ocorrencias["tipo"].isin(["acidente", "fatalidade"])
    ocorrencias["deflatorValor"] = ocorrencias.apply(lambda r: r["nota"] if r["isDeflator"] else 0.0, axis=1)

    def _acidente_label(r):
        if r["tipo"] == "acidente":
            return f"Acidente – {label(r['statusAcidente'])}" if r["statusAcidente"] else "Acidente"
        elif r["tipo"] == "fatalidade":
            return "Fatalidade"
        return None

    ocorrencias["acidenteLabel"] = ocorrencias.apply(_acidente_label, axis=1)

    # ---- colunas auxiliares — eventos ----
    eventos["tipoEvento"] = eventos["tipo_evento"].fillna("").astype(str).str.strip().str.lower()
    eventos["participantes"] = eventos["total_colaboradores_presente"].fillna(0).astype(float)
    eventos["notaEventoNum"] = eventos["nota_evento"].fillna(0).astype(float)
    eventos["qualifica"] = (eventos["participantes"] > 0) & (eventos["notaEventoNum"] > 0)

    docs_grouped = agrupar_docs(docs)
    eventos_grouped = agrupar_eventos(eventos)
    ocorrencias_grouped = agrupar_ocorrencias(ocorrencias)

    # ---- consolidação final ----
    chaves = pd.concat([
        docs_grouped[["id_pc", "mes_ano", "diretor", "gerente", "projeto", "mes_ano_label"]],
        eventos_grouped[["id_pc", "mes_ano", "diretor", "gerente", "projeto", "mes_ano_label"]],
        ocorrencias_grouped[["id_pc", "mes_ano", "diretor", "gerente", "projeto", "mes_ano_label"]],
    ], ignore_index=True).drop_duplicates(subset=["id_pc", "mes_ano"])

    final = chaves.merge(
        docs_grouped.drop(columns=["diretor", "gerente", "projeto", "mes_ano_label"]), on=["id_pc", "mes_ano"], how="left"
    )
    final = final.merge(
        eventos_grouped.drop(columns=["diretor", "gerente", "projeto", "mes_ano_label"]), on=["id_pc", "mes_ano"], how="left"
    )
    final = final.merge(
        ocorrencias_grouped.drop(columns=["diretor", "gerente", "projeto", "mes_ano_label"]), on=["id_pc", "mes_ano"], how="left"
    )

    num_cols = [
        "soma_doc", "qtd_doc", "soma_evt", "qtd_evt", "qtd_dds", "qtd_campanha_sst", "qtd_saude_mental",
        "qtd_registros_ocorrencias_total", "soma_ocor", "qtd_ocor", "soma_def", "qtd_def", "qtd_acidentes",
        "qtd_fatalidades", "qtd_acidentes_afastamento", "qtd_acidentes_sem_afastamento",
        "soma_analise_sst", "qtd_analise_sst", "soma_dias_afastado", "soma_custo_afastamento",
    ]
    final[num_cols] = final[num_cols].fillna(0)
    final["acidentes_detalhe"] = final["acidentes_detalhe"].apply(lambda v: v if isinstance(v, list) else [])

    final["documentos_reprovados"] = final["documentos_reprovados"].fillna("Nenhum documento reprovado no período")
    final["eventos_realizados"] = final["eventos_realizados"].fillna("Nenhum evento registrado no período")
    final["resumo_acidentes"] = final["resumo_acidentes"].fillna("Nenhum acidente registrado no período")

    final["nota_documento"] = final.apply(
        lambda r: round(min(max(r["soma_doc"] / r["qtd_doc"] if r["qtd_doc"] > 0 else 0, 0), 1000), 2), axis=1
    )
    # Teto de 200 confirmado contra o painel oficial (PC 391 - JICA, JUN/2026:
    # soma_evt=375 -> nota_evento=200, nao 300).
    final["nota_evento"] = final["soma_evt"].apply(lambda v: round(min(max(v, 0), 200), 2))
    final["nota_ocorrencia"] = final.apply(
        lambda r: round(min(max(r["soma_ocor"] / r["qtd_ocor"] if r["qtd_ocor"] > 0 else 0, 0), 1000), 2), axis=1
    )
    final["nota_deflatores"] = final["soma_def"].round(2)
    final["nota_analise_sst_media"] = final.apply(
        lambda r: round(min(max(r["soma_analise_sst"] / r["qtd_analise_sst"] if r["qtd_analise_sst"] > 0 else 0, 0), 100), 1),
        axis=1,
    )
    final["nota_effico_sst"] = (
        final["nota_documento"] + final["nota_evento"] + final["nota_ocorrencia"] + final["soma_def"]
    ).clip(-10000, 2300).round(2)

    final["farol"] = final["nota_effico_sst"].apply(
        lambda v: "vermelho" if v < 1000 else ("amarelo" if v < 1600 else "verde")
    )
    final["status_sst"] = final["farol"].map({
        "vermelho": "🔴 VERMELHO", "amarelo": "🟡 AMARELO", "verde": "🟢 VERDE",
    })

    final = final.rename(columns={
        "qtd_def": "qtd_deflatores",
        "qtd_ocor": "qtd_registros_campo_pontuaveis",
        "qtd_doc": "qtd_registros_documento",
        "qtd_evt": "qtd_registros_eventos",
    })

    final = final.sort_values(["mes_ano", "id_pc"]).reset_index(drop=True)

    # ---- evolução mensal (média geral por mês, todos os contratos) ----
    evolucao = (
        final.groupby(["mes_ano", "mes_ano_label"])
        .agg(media_nota_sst=("nota_effico_sst", "mean"), qtd_contratos=("id_pc", "count"))
        .reset_index()
        .sort_values("mes_ano")
    )
    evolucao["media_nota_sst"] = evolucao["media_nota_sst"].round(2)

    # ---- resumo do mês mais recente ----
    if len(final) > 0:
        mes_recente = final["mes_ano"].max()
        recente = final[final["mes_ano"] == mes_recente]
        farol_counts = recente["farol"].value_counts().to_dict()
        resumo = {
            "total_contratos": int(final["id_pc"].nunique()),
            "mes_mais_recente": mes_recente,
            "mes_mais_recente_label": recente["mes_ano_label"].iloc[0],
            "media_nota_sst_mes_recente": round(float(recente["nota_effico_sst"].mean()), 2),
            "farol_mes_recente": {
                "verde": int(farol_counts.get("verde", 0)),
                "amarelo": int(farol_counts.get("amarelo", 0)),
                "vermelho": int(farol_counts.get("vermelho", 0)),
            },
        }
    else:
        resumo = {
            "total_contratos": 0,
            "mes_mais_recente": None,
            "mes_mais_recente_label": None,
            "media_nota_sst_mes_recente": None,
            "farol_mes_recente": {"verde": 0, "amarelo": 0, "vermelho": 0},
        }

    output = {
        "gerado_em": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "resumo": resumo,
        "evolucao_mensal": clean_records(evolucao.to_dict("records")),
        "registros": clean_records(final.to_dict("records")),
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"OK: {len(final)} registros salvos em {os.path.abspath(OUTPUT_PATH)}")


if __name__ == "__main__":
    main()
