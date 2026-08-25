"""
Importa uma exportacao mensal do WeHandle/SABESP e gera/atualiza
data/wehandle_comparativo.json, usado pelo dashboard para mostrar a
comparacao Falcao x SABESP por contrato.

Uso (rodar localmente, nunca no GitHub Actions - o WeHandle nao tem API,
so exportacao manual de planilha):

    python scripts/import_wehandle.py "caminho\\para\\06-jun-EFFICO....xlsx"

O script e ADITIVO: cada execucao atualiza so o(s) mes(es) presentes na
planilha, preservando os meses ja importados anteriormente no JSON.

Metodologia (validada contra os 8 contratos de referencia - a soma dos
4 componentes reconstroi TotalPontosMes com diferenca < 0.5 ponto em
todos os casos, resultado de arredondamento, nao de formula errada):

    documentacao_sabesp = Pont-DocsEmpresa + Pont-DocsFuncionarios + Integracao
      (Integracao entra aqui porque no proprio schema do Falcao ela e um
      tipo de documento - "Integracao SST" - nao um componente separado)
    eventos_sabesp       = DDS + Campanhas + SaudeMental
    campo_sabesp         = Inspecoes
    deflatores_sabesp    = -(Acidentes + Acidentes Fora do Prazo + NaoReportAcidentes)
    nota_final_sabesp    = TotalPontosMes (oficial, nunca recalculado)

Nao inventa idWehandle nao mapeado: se aparecer um idWehandle fora do
MAPA_PC abaixo, o contrato e listado em "nao_mapeados" no resumo impresso
e fica de fora do JSON, em vez de tentar adivinhar o PC Falcao.
"""

import json
import os
import sys
from datetime import datetime

import pandas as pd

# =====================================================================
# MAPA DE CORRELACAO idWehandle -> PC Falcao (fornecido manualmente)
# =====================================================================

MAPA_PC = {
    147167: 10348,
    140862: 10389,
    156613: 10396,
    172821: 10395,
    150925: 10391,
    145344: 10372,
    144023: 10392,
    201131: 10399,
}

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FALCAO_JSON_PATH = os.path.join(SCRIPT_DIR, "..", "data", "falcao_sst.json")
OUTPUT_PATH = os.path.join(SCRIPT_DIR, "..", "data", "wehandle_comparativo.json")

MESES_ARQUIVO = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}
MESES_LABEL = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]


def inferir_competencia(caminho_arquivo, docs_emp):
    """Mes vem do nome do arquivo (ex: '06-jun-...'); ano vem da data
    mais comum na aba de documentos, ja que o nome do arquivo nao traz ano."""
    nome = os.path.basename(caminho_arquivo).lower()
    mes = None
    for abrev, num in MESES_ARQUIVO.items():
        if f"-{abrev}-" in nome or nome.startswith(f"{num:02d}-{abrev}"):
            mes = num
            break
    if mes is None:
        raise ValueError(f"Nao foi possivel inferir o mes a partir do nome do arquivo: {nome}")

    datas = pd.to_datetime(docs_emp["Data da Postagem"], errors="coerce").dropna()
    if len(datas) == 0:
        raise ValueError("Nao ha datas em 'Data da Postagem' para inferir o ano da competencia.")
    ano = int(datas.dt.year.mode()[0])
    return ano, mes


def _clean_text(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s else None


def montar_comparativo(caminho_arquivo):
    xls = pd.ExcelFile(caminho_arquivo)
    abas_esperadas = {"contratos", "documentos empresa", "prestadores funcionarios"}
    faltando = abas_esperadas - set(xls.sheet_names)
    if faltando:
        raise ValueError(f"Abas esperadas ausentes no arquivo: {faltando}")

    contratos = pd.read_excel(caminho_arquivo, sheet_name="contratos")
    docs_emp = pd.read_excel(caminho_arquivo, sheet_name="documentos empresa")
    prest = pd.read_excel(caminho_arquivo, sheet_name="prestadores funcionarios")

    ano, mes = inferir_competencia(caminho_arquivo, docs_emp)
    mes_ano_falcao = f"{ano}-{mes - 1:02d}"  # mesmo esquema zero-indexado do Falcao
    mes_ano_label = f"{MESES_LABEL[mes - 1]}/{ano}"

    ids_contratos = set(contratos["idWehandle"].dropna().astype(int))
    nao_mapeados = sorted(ids_contratos - set(MAPA_PC.keys()))

    # carrega dados do Falcao para casar id_pc + mes_ano
    falcao_por_pc_mes = {}
    if os.path.exists(FALCAO_JSON_PATH):
        with open(FALCAO_JSON_PATH, encoding="utf-8") as f:
            falcao_data = json.load(f)
        for r in falcao_data.get("registros", []):
            falcao_por_pc_mes[(r["id_pc"], r["mes_ano"])] = r

    registros = []
    for _, row in contratos.iterrows():
        id_wh = int(row["idWehandle"])
        if id_wh not in MAPA_PC:
            continue
        id_pc = MAPA_PC[id_wh]

        documentacao_sabesp = round(
            float(row["Pont - DocsEmpresa"] or 0) + float(row["Pont - DocsFuncionarios"] or 0) + float(row["Integração"] or 0), 2
        )
        eventos_sabesp = round(float(row["DDS"] or 0) + float(row["Campanhas"] or 0) + float(row["SaudeMental"] or 0), 2)
        campo_sabesp = round(float(row["Inspecoes"] or 0), 2)
        deflatores_sabesp = round(
            -(float(row["Acidentes"] or 0) + float(row["Acidentes Fora do Prazo"] or 0) + float(row["NaoReportAcidentes"] or 0)), 2
        )
        nota_final_sabesp = round(float(row["TotalPontosMes"] or 0), 2)

        docs_contrato = docs_emp[docs_emp["idWehandle"] == id_wh]
        nao_conforme = docs_contrato[docs_contrato["StatusRanking"] != "Conforme"]
        nomes_nao_conforme = sorted(set(nao_conforme["Documento"].dropna().astype(str).str.strip()))

        prest_contrato = prest[prest["idWehandle"] == id_wh]
        total_colaboradores = int(prest_contrato["wh-pessoa"].nunique())
        pendencias_func = int((prest_contrato["StatusDocumentos"] != "Conforme").sum())

        falcao_rec = falcao_por_pc_mes.get((id_pc, mes_ano_falcao))

        registros.append({
            "id_pc": id_pc,
            "idWehandle": id_wh,
            "mes_ano": mes_ano_falcao,
            "mes_ano_label": mes_ano_label,
            "projeto": falcao_rec.get("projeto") if falcao_rec else _clean_text(row.get("Nome do Consórcio")),
            "falcao": {
                "documentacao": falcao_rec["nota_documento"] if falcao_rec else None,
                "eventos": falcao_rec["nota_evento"] if falcao_rec else None,
                "campo": falcao_rec["nota_ocorrencia"] if falcao_rec else None,
                "deflatores": falcao_rec["nota_deflatores"] if falcao_rec else None,
                "nota_final": falcao_rec["nota_effico_sst"] if falcao_rec else None,
            },
            "sabesp": {
                "documentacao": documentacao_sabesp,
                "eventos": eventos_sabesp,
                "campo": campo_sabesp,
                "deflatores": deflatores_sabesp,
                "nota_final": nota_final_sabesp,
            },
            "auditoria": {
                "total_documentos_empresa": int(len(docs_contrato)),
                "total_nao_conformes_empresa": int(len(nao_conforme)),
                "documentos_nao_conformes": " • ".join(nomes_nao_conforme) if nomes_nao_conforme else "Nenhum documento reprovado no período",
                "total_colaboradores_avaliados": total_colaboradores,
                "total_pendencias_funcionarios": pendencias_func,
                "qtd_pessoas_oficial": int(row["Qtd pessoas"]) if pd.notna(row.get("Qtd pessoas")) else None,
            },
            "fonte": "WeHandle - exportação mensal",
            "arquivo_origem": os.path.basename(caminho_arquivo),
            "gerado_em": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        })

    return registros, nao_mapeados, mes_ano_label


def main():
    if len(sys.argv) < 2:
        sys.exit("Uso: python scripts/import_wehandle.py <caminho_do_arquivo.xlsx>")
    caminho = sys.argv[1]
    if not os.path.exists(caminho):
        sys.exit(f"Arquivo não encontrado: {caminho}")

    registros, nao_mapeados, mes_ano_label = montar_comparativo(caminho)

    if nao_mapeados:
        print(f"AVISO: idWehandle sem mapeamento para PC Falcao (ignorados): {nao_mapeados}")
        print("Se algum desses for um contrato valido, adicione-o em MAPA_PC no script.")

    existentes = []
    if os.path.exists(OUTPUT_PATH):
        with open(OUTPUT_PATH, encoding="utf-8") as f:
            existentes = json.load(f)

    chaves_novas = {(r["id_pc"], r["mes_ano"]) for r in registros}
    mantidos = [r for r in existentes if (r["id_pc"], r["mes_ano"]) not in chaves_novas]
    final = mantidos + registros
    final.sort(key=lambda r: (r["mes_ano"], r["id_pc"]))

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(final, f, ensure_ascii=False, indent=2)

    print(f"OK: {len(registros)} contratos de {mes_ano_label} gravados em {os.path.abspath(OUTPUT_PATH)}")
    print(f"Total de registros no arquivo (todas as competencias importadas ate agora): {len(final)}")
    print("\nProximo passo: revise o arquivo e, se estiver certo, rode:")
    print("  git add data/wehandle_comparativo.json")
    print('  git commit -m "chore: importa comparativo WeHandle <mes/ano>"')
    print("  git push")


if __name__ == "__main__":
    main()
