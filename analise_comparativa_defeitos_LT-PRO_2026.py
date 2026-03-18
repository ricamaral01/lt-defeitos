"""
Análise Comparativa de Erros — Todos os meses de 2026
Executa automaticamente para todos os meses disponíveis na planilha.
Meses futuros sem aba são ignorados.
Para adicionar dados de produção de um novo mês, basta preencher PRODUCAO_RAW abaixo.
"""

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import os
import html as html_module
import sys
import time
import json

# ─── Configurações ───────────────────────────────────────────────────────────

CREDENTIALS_PATH = os.getenv(
    "GOOGLE_CREDENTIALS_PATH",
    r"G:\Outros computadores\Meu laptop\Concrefer\Relatórios Automatizados\Credenciais\sheetsautomacao-477114-9b74d04e460f.json",
)
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID_PRODUCAO", "18Rgr2fWCSrc2NUzgxE9YDO8_LoXhMRxFSNb-jhB_kHU")  # Planilha com dados de produção
SPREADSHEET_ID_DEFEITOS = os.getenv("SPREADSHEET_ID_DEFEITOS", "1OCcDhYan81sNRgrjycYtY6E6tewbRvcCVOPUp6GsQC4")  # Planilha com dados de defeitos
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))  # Mesma pasta do arquivo .py

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def carregar_credenciais():
    """Carrega credenciais via secret JSON (Actions) ou arquivo local."""
    json_secret = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if json_secret:
        info = json.loads(json_secret)
        return Credentials.from_service_account_info(info, scopes=SCOPES)

    if os.path.exists(CREDENTIALS_PATH):
        return Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)

    raise FileNotFoundError(
        "Credenciais não encontradas. Defina GOOGLE_SERVICE_ACCOUNT_JSON ou GOOGLE_CREDENTIALS_PATH válido."
    )

# ─── Meses de 2026 (aba → label) ─────────────────────────────────────────────

MESES_2026 = [
    ("jan26", "Jan/26"),
    ("fev26", "Fev/26"),
    ("mar26", "Mar/26"),
    ("abr26", "Abr/26"),
    ("mai26", "Mai/26"),
    ("jun26", "Jun/26"),
    ("jul26", "Jul/26"),
    ("ago26", "Ago/26"),
    ("set26", "Set/26"),
    ("out26", "Out/26"),
    ("nov26", "Nov/26"),
    ("dez26", "Dez/26"),
]

# ─── Produção por Setor por Mês ──────────────────────────────────────────────
# Para cada mês, cole os dados de produção no formato:
#   "Produto|SETOR N|quantidade"
# Meses sem dados ficam com string vazia — o relatório indica "sem dados".

# ─── Função para ler dados de produção do Google Sheets ──────────────────────

def ler_producao_sheets(creds, spreadsheet_id, aba_nome, celula):
    """Lê um valor de célula específica da planilha Google Sheets."""
    try:
        client = gspread.authorize(creds)
        planilha = client.open_by_key(spreadsheet_id)
        aba = planilha.worksheet(aba_nome)
        valor = aba.acell(celula).value
        if valor:
            valor_str = str(valor).strip()
            # Se for um número puro, formata como "Produção Total|TOTAL|valor"
            try:
                # Tenta converter para int/float (aceita formatação com ponto ou vírgula)
                num_test = valor_str.replace(".", "").replace(",", "").strip()
                if num_test.isdigit():
                    return f"Produção Total|TOTAL|{valor_str}"
            except:
                pass
            # Se contiver pipes, retorna como está (é um CSV formatado)
            if "|" in valor_str:
                return valor_str
            # Se for apenas número, mesmo assim trata como total
            try:
                float(valor_str.replace(",", "."))
                return f"Produção Total|TOTAL|{valor_str}"
            except:
                # Não é número, retorna como está
                return valor_str
        return ""
    except Exception as e:
        print(f"[AVISO] Erro ao ler {aba_nome}!{celula}: {e}")
        return ""


def parse_int_br(valor):
    """Converte texto numérico BR para inteiro (ex.: 1.234, 1234, 1234,0)."""
    if valor is None:
        return None
    s = str(valor).strip()
    if not s:
        return None
    s = s.replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return int(float(s))
    except ValueError:
        return None


def ler_producao_setorial_marco(planilha_prod):
    """Lê produção setorial de março no intervalo Estoque!C700:D704."""
    try:
        aba = planilha_prod.worksheet("Estoque")
        rows = aba.get("C700:D704")
    except Exception as e:
        print(f"[AVISO] Erro ao ler produção setorial em Estoque!C700:D704: {e}")
        return None

    prod = {"SETOR 1": 0, "SETOR 2": 0, "SETOR 3": 0, "SETOR 4": 0}

    for row in rows:
        c = str(row[0]).strip() if len(row) > 0 else ""
        d = str(row[1]).strip() if len(row) > 1 else ""

        qtd_c = parse_int_br(c)
        qtd_d = parse_int_br(d)

        if qtd_d is not None and c:
            setor_txt = c
            qtd = qtd_d
        elif qtd_c is not None and d:
            setor_txt = d
            qtd = qtd_c
        else:
            continue

        setor_key = mapear_setor_producao(setor_txt)
        if setor_key in prod and qtd > 0:
            prod[setor_key] += qtd

    if sum(prod.values()) == 0:
        return None
    return prod

PRODUCAO_RAW = {
    "jan26": """Produção Total|TOTAL|2657""",  # Somente total (sem detalhamento por setor)
    "fev26": """12x1500|SETOR 3|5
24x1000|SETOR 3|1
13x1000|SETOR 3|2
12x1000 AAT|SETOR 3|1
12x1000|SETOR 3|89
11x1000|SETOR 3|39
10,5x1000|SETOR 3|5
10x1000|SETOR 3|3
15x600|SETOR 3|2
12x600 AAT|SETOR 3|2
12x600|SETOR 3|91
11x600|SETOR 3|57
10x600|SETOR 3|8
9x600|SETOR 3|30
13x400|SETOR 3|1
12x400|SETOR 3|140
10x400|SETOR 3|1
9x400|SETOR 3|147
8x400|SETOR 3|1
10x300|SETOR 3|15
9x300|SETOR 3|19
8x300|SETOR 3|3
11x200|SETOR 3|75
10x200|SETOR 3|20
9x200|SETOR 3|207
7x200|SETOR 3|5
18x400 CR|SETOR 3|3
6x90 DT|SETOR 4|47
7,5x90 DT|SETOR 4|147
7,5x200 DT|SETOR 4|75
7,5x300 DT|SETOR 4|98
7,5x400 DT|SETOR 4|5
7,5x600 DT|SETOR 4|25
7,5x700 DT|SETOR 4|10
7,5x800 DT|SETOR 4|1
Cemig DT 7,0x150|SETOR 4|32
Placa Estai 20x10 x 100|SETOR 4|302
Chápeu Sub 100|SETOR 4|112
Chápeu Sub 200 TC|SETOR 4|18
Pedestal Caixa Sabesp|SETOR 4|28
Padrão Completo 2 cx VR|SETOR 2|239
Padrão Completo 2 cx VL|SETOR 1|54
Padrão Completo 1 cx VR|SETOR 2|667
Padrão Completo 1 cx VL|SETOR 2|240
Padrão Completo 1 cx VL (EDP)|SETOR 1|271
Padrão Completo 2 cx VR Elektro|SETOR 1|85
Padrão Completo 2 cx VL Elektro|SETOR 1|42
Padrão Completo 1 cx VR Elektro|SETOR 1|423
Padrão Completo 1 cx VL Elektro|SETOR 1|66
Padrão Completo 4 cx VR|SETOR 1|11
Padrão Completo 4 cx VL|SETOR 1|6
Padrão Completo 3 cx VR|SETOR 1|14
Padrão Completo 3 cx VL|SETOR 1|4
Poste 7,5 x 600 VL c/ TC|SETOR 2|38
Poste 7,5 x 600 VR c/ TC|SETOR 2|37
Poste Subterrâneo 100 A - Elektro|SETOR 1|21
Poste Subterrâneo 200 A - TC|SETOR 1|12
Poste Totem|SETOR 4|3
Poste Visor Aéreo 1 cx VL (7,5x300)|SETOR 2|38
Poste Econômico 1CX VR|SETOR 1|27
Poste Econômico 2CX VR|SETOR 1|1
Poste Econômico 3CX VR|SETOR 1|3""",
    "mar26": "",  # ← Cole aqui quando tiver os dados de Março
    "abr26": "",  # ← Cole aqui quando tiver os dados de Abril
    "mai26": "",  # ← Cole aqui quando tiver os dados de Maio
    "jun26": "",  # ← Cole aqui quando tiver os dados de Junho
    "jul26": "",  # ← Cole aqui quando tiver os dados de Julho
    "ago26": "",  # ← Cole aqui quando tiver os dados de Agosto
    "set26": "",  # ← Cole aqui quando tiver os dados de Setembro
    "out26": "",  # ← Cole aqui quando tiver os dados de Outubro
    "nov26": "",  # ← Cole aqui quando tiver os dados de Novembro
    "dez26": "",  # ← Cole aqui quando tiver os dados de Dezembro
}

# ─── Parsear produção ────────────────────────────────────────────────────────

def parsear_producao(raw_str):
    """Retorna dict {SETOR N: quantidade} a partir da string raw."""
    prod = {"SETOR 1": 0, "SETOR 2": 0, "SETOR 3": 0, "SETOR 4": 0}
    if not raw_str or not raw_str.strip():
        return prod
    for line in raw_str.strip().split("\n"):
        if not line.strip():
            continue
        parts = line.split("|")
        if len(parts) == 3:
            setor = parts[1].strip().upper()
            try:
                qtd = int(parts[2].strip())
            except ValueError:
                continue
            
            # Caso especial: se setor for "TOTAL", distribui entre todos os setores igualmente
            if setor == "TOTAL":
                qtd_por_setor = qtd // 4
                for k in prod:
                    prod[k] = qtd_por_setor
                # Distribui o resto (se houver)
                resto = qtd % 4
                for i in range(resto):
                    prod[f"SETOR {i+1}"] += 1
            else:
                setor_key = setor
                for k in prod:
                    if setor.replace(" ", "").lower() == k.replace(" ", "").lower():
                        setor_key = k
                        break
                if setor_key not in prod:
                    prod[setor_key] = 0
                prod[setor_key] += qtd
    return prod

PRODUCAO = {}
for aba, _ in MESES_2026:
    PRODUCAO[aba] = parsear_producao(PRODUCAO_RAW.get(aba, ""))

for aba, label in MESES_2026:
    total = sum(PRODUCAO[aba].values())
    if total > 0:
        print(f"[INFO] Produção {label} por setor: {PRODUCAO[aba]}, total: {total}")


# ─── Conexão com Google Sheets ────────────────────────────────────────────────

def conectar_planilha(planilha, aba_nome):
    """Lê uma aba da planilha. Retorna DataFrame ou None se a aba não existir."""
    time.sleep(0.5)  # Delay para respeitar quotas da API
    try:
        aba = planilha.worksheet(aba_nome)
    except gspread.exceptions.WorksheetNotFound:
        return None

    all_values = aba.get_all_values()
    if not all_values:
        return None

    # Detecta cabeçalho (linha com mais células preenchidas)
    header_idx = 0
    max_filled = 0
    for i, row in enumerate(all_values[:20]):
        filled = sum(1 for c in row if str(c).strip())
        if filled > max_filled:
            max_filled = filled
            header_idx = i

    headers = all_values[header_idx]
    seen = {}
    clean_headers = []
    for h in headers:
        h = str(h).strip()
        if not h:
            h = "_vazio"
        if h in seen:
            seen[h] += 1
            clean_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            clean_headers.append(h)

    data_rows = all_values[header_idx + 1:]
    df = pd.DataFrame(data_rows, columns=clean_headers)
    df = df.loc[:, ~df.columns.str.startswith("_vazio")]
    df = df.replace("", pd.NA).dropna(how="all").fillna("")

    # Validação: somente linhas com Colunas A(0), C(2), E(4) e G(6) preenchidas
    cols_list = list(df.columns)
    required_positions = [0, 2, 4, 6]  # Colunas A, C, E, G
    valid_positions = [p for p in required_positions if p < len(cols_list)]
    if valid_positions:
        mask = df.apply(
            lambda row: all(str(row.iloc[p]).strip() != "" for p in valid_positions),
            axis=1
        )
    else:
        mask = df.apply(lambda row: sum(1 for v in row if str(v).strip()) >= 3, axis=1)
    df = df[mask].reset_index(drop=True)

    # Remover linhas-artefato
    LIXO = ["total -->", "problema (padrão + comum)", "problema (padrao + comum)",
            "setor", "lista", "controle produtos não conforme",
            "controle produtos nao conforme"]
    for col in df.columns:
        df = df[~df[col].astype(str).str.strip().str.lower().isin(LIXO)]
    df = df.reset_index(drop=True)

    if df.empty:
        return None
    return df


# ─── Mapeamento de colunas ────────────────────────────────────────────────────

def mapear_colunas(df):
    colunas = {col.strip().lower(): col for col in df.columns}
    m = {}
    buscas = {
        "erro": ["tipo defeito", "tipo de defeito", "tipo de erro", "defeito"],
        "setor": ["setor"],
        "data": ["data não conformidade", "data nao conformidade", "data fabricação", "data fabricacao", "data"],
        "como_ocorreu": ["como ocorreu ?", "como ocorreu", "como ocorreu?"],
        "produto": ["modelo do poste", "modelo", "produto"],
        "forma": ["nº forma", "n forma", "forma"],
        "acao": ["ação imediata", "acao imediata"],
        "responsavel": ["responsável pela ação", "responsavel pela acao", "responsável", "responsavel"],
        "retrabalho": ["conseguiu retrabalhar? sim ou não", "conseguiu retrabalhar", "retrabalho"],
        "reinspecionado": ["reinspecionado por", "reinspecionado"],
        "tipo_poste": ["tipo (padrão/comum ou cir / dt)", "tipo padrão", "tipo"],
        "rnc": ["rnc"],
    }
    for chave, termos in buscas.items():
        for t in termos:
            if t in colunas:
                m[chave] = colunas[t]
                break

    # Fallback por posição — abas com cabeçalhos genéricos ("Coluna N")
    # Posições conhecidas: 0=produto, 1=tipo_poste, 2=forma, 3=setor,
    # 4=data(fab), 5=erro, 6=como_ocorreu, 7=acao, 8=responsavel,
    # 9=data(NC), 10=retrabalho, 11=data_retrab, 12=rnc, 13=reinspecionado
    pos_map = {
        "produto": 0, "tipo_poste": 1, "forma": 2, "setor": 3,
        "data": 4, "erro": 5, "como_ocorreu": 6, "acao": 7,
        "responsavel": 8, "retrabalho": 10, "rnc": 12, "reinspecionado": 13,
    }
    cols_list = list(df.columns)
    for chave, pos in pos_map.items():
        if chave not in m and pos < len(cols_list):
            m[chave] = cols_list[pos]

    return m


# ─── Análise completa de uma aba ──────────────────────────────────────────────

def analisar_aba(df, mapeamento, nome_mes):
    r = {"mes": nome_mes, "total": len(df), "mapeamento": mapeamento}

    # Tipo defeito
    if "erro" in mapeamento:
        col = mapeamento["erro"]
        df[col] = df[col].astype(str).str.strip()
        df_com_erro = df[df[col] != ""]
        r["total_com_erro"] = len(df_com_erro)

        contagem = df_com_erro[col].value_counts().reset_index()
        contagem.columns = ["Tipo de Erro", "Frequência"]
        contagem["% do Total"] = (contagem["Frequência"] / contagem["Frequência"].sum() * 100).round(2)
        contagem["% Acumulado"] = contagem["% do Total"].cumsum().round(2)
        r["erros_pareto"] = contagem

        vitais = contagem[contagem["% Acumulado"] <= 80]
        if vitais.empty:
            vitais = contagem.head(1)
        r["erros_vitais"] = vitais
    else:
        r["erros_pareto"] = pd.DataFrame()
        r["total_com_erro"] = 0

    # Setor
    if "setor" in mapeamento:
        col = mapeamento["setor"]
        df[col] = df[col].astype(str).str.strip()
        df_setor = df[df[col] != ""]
        contagem = df_setor[col].value_counts().reset_index()
        contagem.columns = ["Setor", "Frequência"]
        contagem["% do Total"] = (contagem["Frequência"] / contagem["Frequência"].sum() * 100).round(2)
        r["setores"] = contagem

        if "erro" in mapeamento:
            r["cross_erro_setor"] = pd.crosstab(df_setor[mapeamento["erro"]], df_setor[col])
    else:
        r["setores"] = pd.DataFrame()

    # Como ocorreu
    if "como_ocorreu" in mapeamento:
        col = mapeamento["como_ocorreu"]
        df[col] = df[col].astype(str).str.strip()
        df_co = df[df[col] != ""]
        contagem = df_co[col].value_counts().reset_index()
        contagem.columns = ["Como Ocorreu", "Frequência"]
        contagem["% do Total"] = (contagem["Frequência"] / contagem["Frequência"].sum() * 100).round(2)
        r["como_ocorreu"] = contagem
    else:
        r["como_ocorreu"] = pd.DataFrame()

    # Retrabalho
    if "retrabalho" in mapeamento:
        col = mapeamento["retrabalho"]
        df[col] = df[col].astype(str).str.strip().str.upper()
        sim = df[col].isin(["S", "SIM"]).sum()
        nao = df[col].isin(["N", "NÃO", "NAO"]).sum()
        total = sim + nao
        r["retrabalho"] = {
            "sim": int(sim), "nao": int(nao), "total": int(total),
            "taxa": round(sim / total * 100, 1) if total > 0 else 0,
        }
    else:
        r["retrabalho"] = None

    # Formas
    if "forma" in mapeamento:
        col = mapeamento["forma"]
        contagem = df[col].value_counts().head(20).reset_index()
        contagem.columns = ["Nº Forma", "Frequência"]
        r["formas"] = contagem
    else:
        r["formas"] = pd.DataFrame()

    # Produtos
    if "produto" in mapeamento:
        col = mapeamento["produto"]
        contagem = df[col].value_counts().head(15).reset_index()
        contagem.columns = ["Modelo", "Frequência"]
        r["produtos"] = contagem
    else:
        r["produtos"] = pd.DataFrame()

    # Tipo poste
    if "tipo_poste" in mapeamento:
        col = mapeamento["tipo_poste"]
        df[col] = df[col].astype(str).str.strip()
        contagem = df[col].value_counts().reset_index()
        contagem.columns = ["Tipo Poste", "Frequência"]
        r["tipo_poste"] = contagem
    else:
        r["tipo_poste"] = pd.DataFrame()

    # Reinspecionado
    if "reinspecionado" in mapeamento:
        col = mapeamento["reinspecionado"]
        contagem = df[col].value_counts().head(10).reset_index()
        contagem.columns = ["Inspetor", "Frequência"]
        r["inspetores"] = contagem
    else:
        r["inspetores"] = pd.DataFrame()

    r["df"] = df
    return r


# ─── Mapear setor da planilha para setor de produção ──────────────────────────

def mapear_setor_producao(setor_planilha):
    s = str(setor_planilha).strip().lower()
    mapping = {
        "montagem 1": "SETOR 1", "montagem 2": "SETOR 2",
        "montagem 3": "SETOR 3", "montagem 4": "SETOR 4",
        "montagem1": "SETOR 1", "montagem2": "SETOR 2",
        "montagem3": "SETOR 3", "montagem4": "SETOR 4",
        "setor 1": "SETOR 1", "setor 2": "SETOR 2",
        "setor 3": "SETOR 3", "setor 4": "SETOR 4",
        "dt": "SETOR 4", "retrabalho": None,
    }
    for key, val in mapping.items():
        if key in s:
            return val
    return setor_planilha.strip().upper()


# ─── Comparação entre dois meses ─────────────────────────────────────────────

def comparar_dois_meses(anterior, atual, aba_atual):
    """Compara dois resultados de analisar_aba. Retorna dict com comparação."""
    comp = {}
    lbl_ant = anterior["mes"]
    lbl_atu = atual["mes"]
    comp["label_ant"] = lbl_ant
    comp["label_atu"] = lbl_atu

    comp["total_ant"] = anterior["total"]
    comp["total_atu"] = atual["total"]
    comp["variacao_total"] = atual["total"] - anterior["total"]
    if anterior["total"] > 0:
        comp["variacao_pct"] = round((atual["total"] - anterior["total"]) / anterior["total"] * 100, 1)
    else:
        comp["variacao_pct"] = 0

    # Erros
    if not anterior["erros_pareto"].empty and not atual["erros_pareto"].empty:
        ea = anterior["erros_pareto"][["Tipo de Erro", "Frequência"]].rename(columns={"Frequência": lbl_ant})
        eb = atual["erros_pareto"][["Tipo de Erro", "Frequência"]].rename(columns={"Frequência": lbl_atu})
        merged = pd.merge(ea, eb, on="Tipo de Erro", how="outer").fillna(0)
        merged[lbl_ant] = merged[lbl_ant].astype(int)
        merged[lbl_atu] = merged[lbl_atu].astype(int)
        merged["Variação"] = merged[lbl_atu] - merged[lbl_ant]
        merged["Var %"] = merged.apply(
            lambda r: round((r[lbl_atu] - r[lbl_ant]) / r[lbl_ant] * 100, 1) if r[lbl_ant] > 0
            else (100.0 if r[lbl_atu] > 0 else 0), axis=1
        )
        merged = merged.sort_values(lbl_atu, ascending=False).reset_index(drop=True)
        comp["erros_comp"] = merged
    else:
        comp["erros_comp"] = pd.DataFrame()

    # Setores
    prod = PRODUCAO.get(aba_atual, {})
    prod_total = sum(prod.values())
    comp["producao"] = prod
    comp["producao_total"] = prod_total

    if not anterior["setores"].empty and not atual["setores"].empty:
        sa = anterior["setores"][["Setor", "Frequência"]].rename(columns={"Frequência": lbl_ant})
        sb = atual["setores"][["Setor", "Frequência"]].rename(columns={"Frequência": lbl_atu})
        merged_s = pd.merge(sa, sb, on="Setor", how="outer").fillna(0)
        merged_s[lbl_ant] = merged_s[lbl_ant].astype(int)
        merged_s[lbl_atu] = merged_s[lbl_atu].astype(int)
        merged_s["Variação"] = merged_s[lbl_atu] - merged_s[lbl_ant]
        merged_s["Setor Produção"] = merged_s["Setor"].apply(mapear_setor_producao)
        merged_s["Produção"] = merged_s["Setor Produção"].apply(lambda x: prod.get(x, 0))
        merged_s["Taxa Defeito (%)"] = merged_s.apply(
            lambda r: round(r[lbl_atu] / r["Produção"] * 100, 2) if r["Produção"] > 0 else 0, axis=1
        )
        merged_s = merged_s.sort_values(lbl_atu, ascending=False).reset_index(drop=True)
        comp["setores_comp"] = merged_s
    elif not atual["setores"].empty:
        sc = atual["setores"].copy()
        sc.rename(columns={"Frequência": lbl_atu}, inplace=True)
        sc[lbl_ant] = 0
        sc["Variação"] = sc[lbl_atu]
        sc["Setor Produção"] = sc["Setor"].apply(mapear_setor_producao)
        sc["Produção"] = sc["Setor Produção"].apply(lambda x: prod.get(x, 0))
        sc["Taxa Defeito (%)"] = sc.apply(
            lambda r: round(r[lbl_atu] / r["Produção"] * 100, 2) if r["Produção"] > 0 else 0, axis=1
        )
        comp["setores_comp"] = sc
    else:
        comp["setores_comp"] = pd.DataFrame()

    # Como ocorreu
    if not anterior["como_ocorreu"].empty and not atual["como_ocorreu"].empty:
        ca = anterior["como_ocorreu"][["Como Ocorreu", "Frequência"]].rename(columns={"Frequência": lbl_ant})
        cb = atual["como_ocorreu"][["Como Ocorreu", "Frequência"]].rename(columns={"Frequência": lbl_atu})
        merged_c = pd.merge(ca, cb, on="Como Ocorreu", how="outer").fillna(0)
        merged_c[lbl_ant] = merged_c[lbl_ant].astype(int)
        merged_c[lbl_atu] = merged_c[lbl_atu].astype(int)
        merged_c["Variação"] = merged_c[lbl_atu] - merged_c[lbl_ant]
        merged_c = merged_c.sort_values(lbl_atu, ascending=False).reset_index(drop=True)
        comp["como_comp"] = merged_c
    else:
        comp["como_comp"] = pd.DataFrame()

    # Retrabalho
    if anterior["retrabalho"] and atual["retrabalho"]:
        comp["retrabalho_ant"] = anterior["retrabalho"]
        comp["retrabalho_atu"] = atual["retrabalho"]

    # Taxa geral
    comp["taxa_defeito_geral"] = round(atual["total"] / prod_total * 100, 2) if prod_total > 0 else 0

    return comp


# ─── Plano de ação DMAIC — top 2 ─────────────────────────────────────────────

def gerar_planos(erros_top2, comp_erros, resultado_atual, aba_atual):
    planos = []
    if erros_top2.empty:
        return planos

    lbl_atu = resultado_atual["mes"]
    prod = PRODUCAO.get(aba_atual, {})
    prod_total = sum(prod.values())
    lbl_ant = comp_erros.columns[1] if not comp_erros.empty and len(comp_erros.columns) > 1 else ""

    for rank, (_, row) in enumerate(erros_top2.iterrows(), 1):
        erro = row["Tipo de Erro"]
        freq = row["Frequência"]
        pct = row["% do Total"]

        tendencia = ""
        ant_v = 0
        if not comp_erros.empty and lbl_ant:
            match = comp_erros[comp_erros["Tipo de Erro"] == erro]
            if not match.empty:
                var = match.iloc[0]["Variação"]
                ant_v = int(match.iloc[0][lbl_ant])
                if var > 0:
                    tendencia = f"⬆️ PIORANDO: +{var} ocorrências vs mês anterior ({ant_v}→{freq})"
                elif var < 0:
                    tendencia = f"⬇️ MELHORANDO: {var} ocorrências vs mês anterior ({ant_v}→{freq})"
                else:
                    tendencia = f"➡️ ESTÁVEL: mesma quantidade que mês anterior ({freq})"

        prioridade = "CRÍTICA" if pct > 20 else ("ALTA" if pct > 10 else ("MÉDIA" if pct > 5 else "BAIXA"))

        setores_detalhe = ""
        if "cross_erro_setor" in resultado_atual:
            cross = resultado_atual["cross_erro_setor"]
            if erro in cross.index:
                row_cross = cross.loc[erro]
                partes = [f"{col}: {int(v)} ({v/freq*100:.1f}%)" for col, v in row_cross.items() if v > 0]
                setores_detalhe = "; ".join(partes)

        meta_90d = max(1, int(freq * 0.5))
        meta_semanal = round(freq / 4, 1)
        taxa_prod = f"{freq/prod_total*100:.2f}%" if prod_total > 0 else "N/D (sem dados de produção)"

        plano = {
            "erro": erro, "frequencia": freq, "percentual": pct,
            "prioridade": prioridade, "tendencia": tendencia, "rank": rank,
            "dmaic": {
                "Define": (
                    f"<strong>Problema:</strong> '{erro}' é o <strong>#{rank} defeito mais frequente</strong>, "
                    f"representando <strong>{pct:.1f}%</strong> de todos os defeitos "
                    f"(<strong>{freq} ocorrências</strong> em {lbl_atu}). "
                    f"{'Variação: ' + tendencia + '.' if tendencia else ''}<br>"
                    f"<strong>Escopo:</strong> Todos os setores produtivos afetados. "
                    f"{'Distribuição: ' + setores_detalhe + '.' if setores_detalhe else ''}<br>"
                    f"<strong>Objetivo:</strong> Reduzir a ocorrência de '{erro}' em 50% nos próximos 90 dias.<br>"
                    f"<strong>Impacto:</strong> Taxa sobre produção: {taxa_prod}.<br>"
                    f"<strong>Equipe sugerida:</strong> Supervisor de produção, Líder de qualidade, Operadores dos setores afetados, Manutenção."
                ),
                "Measure": (
                    f"<strong>Situação atual (baseline):</strong> {freq} ocorrências em {lbl_atu} "
                    f"({pct:.1f}% do total de defeitos). Mês anterior: {ant_v} ocorrências.<br>"
                    f"<strong>Meta:</strong> Reduzir para no máximo <strong>{meta_90d} ocorrências/mês</strong> em 90 dias (redução de 50%).<br>"
                    f"<strong>Meta intermediária:</strong> Máximo {meta_semanal} ocorrências/semana.<br>"
                    f"<strong>KPIs de acompanhamento:</strong><br>"
                    f"&nbsp;&nbsp;• Taxa de '{erro}' por 1.000 unidades produzidas<br>"
                    f"&nbsp;&nbsp;• Nº de ocorrências semanais (carta de tendência)<br>"
                    f"&nbsp;&nbsp;• % de retrabalho com sucesso para este tipo de defeito<br>"
                    f"&nbsp;&nbsp;• Custo estimado de retrabalho/sucata por ocorrência<br>"
                    f"<strong>Sistema de medição:</strong> Validar MSA (Análise do Sistema de Medição) — "
                    f"os inspetores classificam '{erro}' da mesma forma? Existe acordo entre turnos?"
                ),
                "Analyze": (
                    f"<strong>Ishikawa (6M) para '{erro}':</strong><br>"
                    f"&nbsp;&nbsp;• <strong>Mão de obra:</strong> Treinamento adequado? Fadiga? Rotatividade de operadores?<br>"
                    f"&nbsp;&nbsp;• <strong>Método:</strong> Procedimento padrão existe e é seguido? Instrução de trabalho visual disponível?<br>"
                    f"&nbsp;&nbsp;• <strong>Máquina:</strong> Formas/moldes desgastados? Manutenção preventiva em dia? Vibrador calibrado?<br>"
                    f"&nbsp;&nbsp;• <strong>Material:</strong> Concreto na especificação? Agregados com granulometria adequada?<br>"
                    f"&nbsp;&nbsp;• <strong>Meio ambiente:</strong> Temperatura, umidade afetam a cura? Condições do galpão?<br>"
                    f"&nbsp;&nbsp;• <strong>Medição:</strong> Critério de inspeção claro? Calibração dos instrumentos?<br><br>"
                    f"<strong>5 Porquês:</strong> Realizar sessão com equipe multidisciplinar, partindo de cada ocorrência "
                    f"recente de '{erro}'. Documentar as cadeias causais.<br>"
                    f"<strong>Análise estatística:</strong> Estratificar por setor, forma, turno e operador. "
                    f"Verificar se há concentração.<br>"
                    f"<strong>Gemba Walk:</strong> Ir ao chão de fábrica observar o processo nos pontos onde '{erro}' mais ocorre."
                ),
                "Improve": (
                    f"<strong>Ações de curto prazo (semana 1-2):</strong><br>"
                    f"&nbsp;&nbsp;1. Alerta visual (Andon) nos postos com maior incidência de '{erro}'<br>"
                    f"&nbsp;&nbsp;2. Reunião de conscientização com operadores dos setores afetados<br>"
                    f"&nbsp;&nbsp;3. Checklist de verificação pré-operação (setup da forma, preparo do concreto)<br><br>"
                    f"<strong>Ações de médio prazo (semana 3-8):</strong><br>"
                    f"&nbsp;&nbsp;4. Poka-Yoke: dispositivos à prova de erro nos pontos críticos<br>"
                    f"&nbsp;&nbsp;5. Padronização do processo (Work Instruction visual) para '{erro}'<br>"
                    f"&nbsp;&nbsp;6. Treinamento prático focado com os operadores (OJT - On the Job Training)<br>"
                    f"&nbsp;&nbsp;7. Manutenção preventiva nas formas/moldes com maior incidência<br><br>"
                    f"<strong>Ações de longo prazo (semana 9-12):</strong><br>"
                    f"&nbsp;&nbsp;8. Kaizen Blitz dedicado a '{erro}' com equipe cross-funcional<br>"
                    f"&nbsp;&nbsp;9. Análise de investimento em melhoria de equipamento/processo<br>"
                    f"&nbsp;&nbsp;10. Benchmark com outras plantas ou melhores práticas do setor"
                ),
                "Control": (
                    f"<strong>Monitoramento contínuo:</strong><br>"
                    f"&nbsp;&nbsp;• Carta de controle X-barra R para '{erro}' — atualização diária<br>"
                    f"&nbsp;&nbsp;• Dashboard visual na área de produção (gestão à vista)<br>"
                    f"&nbsp;&nbsp;• Reunião diária de 5 min (Daily Management) com revisão de defeitos do dia anterior<br><br>"
                    f"<strong>Auditorias:</strong><br>"
                    f"&nbsp;&nbsp;• LPA (Layered Process Audit): operador diário, líder semanal, gerente mensal<br>"
                    f"&nbsp;&nbsp;• Verificação semanal do cumprimento das ações implementadas<br><br>"
                    f"<strong>Gatilhos de escalonamento:</strong><br>"
                    f"&nbsp;&nbsp;• Se ocorrências semanais > {meta_semanal}: revisão imediata com supervisor<br>"
                    f"&nbsp;&nbsp;• Se 2 semanas consecutivas acima da meta: reabertura da fase Analyze<br>"
                    f"&nbsp;&nbsp;• Se meta de 90 dias não atingida: escalonar para gerência de qualidade<br><br>"
                    f"<strong>Sustentação:</strong><br>"
                    f"&nbsp;&nbsp;• Incorporar melhorias no procedimento operacional padrão (POP)<br>"
                    f"&nbsp;&nbsp;• Incluir no treinamento de integração de novos colaboradores<br>"
                    f"&nbsp;&nbsp;• Revisão trimestral dos resultados para validar sustentação"
                ),
            },
            "wald": {
                "perguntas": [
                    f"Quantas vezes '{erro}' ocorreu mas NÃO foi registrado? O sistema de inspeção capta todos os casos ou há seleção no ponto de controle?",
                    f"Existem defeitos piores que '{erro}' que não aparecem nas estatísticas porque a peça foi descartada antes da inspeção final?",
                    f"Setores com poucos registros de '{erro}' realmente têm qualidade superior ou há subnotificação por diferença de critério entre inspetores?",
                    f"Há formas/moldes com ZERO registros de '{erro}'? É excelência real ou simplesmente não foram inspecionados com o mesmo rigor?",
                    f"Os defeitos mais FREQUENTES são os mais CUSTOSOS? O custo total (retrabalho + sucata + garantia + perda de imagem) está sendo medido e priorizado?",
                    f"Peças que passaram pela inspeção e foram ao cliente sem '{erro}' detectado — quantas reclamações de campo existem? Esses são os 'aviões que não voltaram'.",
                ],
                "recomendacoes": [
                    "Auditar pontos de inspeção em TODOS os estágios do processo, não só na inspeção final.",
                    "Investigar áreas com 'zero defeitos' como possível subnotificação. Realizar auditorias surpresa.",
                    "Coletar dados em pontos intermediários do processo produtivo para captar defeitos onde eles NASCEM.",
                    "Cruzar reclamações de campo (garantia, devoluções) com dados internos — a diferença revela os defeitos que escaparam.",
                    "Comparar a taxa de detecção entre inspetores (estudo de concordância) para garantir que o sistema de medição é confiável.",
                ],
            },
        }
        planos.append(plano)
    return planos


# ─── Médias históricas ────────────────────────────────────────────────────────

MEDIA_2024 = 5.3   # % de defeito médio em 2024
MEDIA_2025 = 7.6   # % de defeito médio em 2025
META_RETRABALHO = 4.5  # meta mensal %


# ─── Leitura do Resumo 2025 (dados históricos) ───────────────────────────────

MESES_LABELS_2025 = ["Jan/25", "Fev/25", "Mar/25", "Abr/25", "Mai/25", "Jun/25",
                     "Jul/25", "Ago/25", "Set/25", "Out/25", "Nov/25", "Dez/25"]

def ler_resumo_2025(planilha):
    """Lê Resumo 2025 AB1:AO17 e retorna dict com dados de 2025."""
    time.sleep(0.5)  # Delay para respeitar quotas da API
    try:
        aba = planilha.worksheet("Resumo 2025")
    except gspread.exceptions.WorksheetNotFound:
        return None

    dados = aba.get("AB1:AO17")
    if not dados or len(dados) < 14:
        return None

    resumo = {"categorias": [], "retrabalho_mensal": [], "producao_mensal": [],
              "pct_mensal": [], "meta_mensal": []}

    # Rows 2-12: error categories with monthly counts
    for row in dados[1:12]:
        cat_name = row[0].strip() if row else ""
        valores = []
        for v in row[1:13]:
            try:
                valores.append(float(str(v).replace(",", ".").replace("#REF!", "0")))
            except (ValueError, IndexError):
                valores.append(0)
        resumo["categorias"].append({"nome": cat_name, "valores": valores})

    # Row 14: Quantidade Total Retrabalho
    if len(dados) >= 14:
        row14 = dados[13]
        for v in row14[1:13]:
            try:
                resumo["retrabalho_mensal"].append(float(str(v).replace(",", ".").replace("#REF!", "0")))
            except (ValueError, IndexError):
                resumo["retrabalho_mensal"].append(0)

    # Row 15: Produção Total
    if len(dados) >= 15:
        row15 = dados[14]
        for v in row15[1:13]:
            try:
                resumo["producao_mensal"].append(float(str(v).replace(",", ".").replace(".", "").replace("#REF!", "0")))
            except (ValueError, IndexError):
                resumo["producao_mensal"].append(0)

    # Row 16: Percentual do retrabalho
    if len(dados) >= 16:
        row16 = dados[15]
        for v in row16[1:13]:
            try:
                resumo["pct_mensal"].append(float(str(v).replace(",", ".").replace("#REF!", "0")))
            except (ValueError, IndexError):
                resumo["pct_mensal"].append(0)

    return resumo


# ─── Construir dados Resumo 2026 a partir da análise ─────────────────────────

def construir_resumo_2026(todos_resultados, abas_disponiveis):
    """Constrói a tabela Resumo 2026 a partir dos dados de análise mensal."""
    # Categorias padrão (prefixo letra)
    CATEGORIAS = [
        "A -  Falha na concretagem",
        "B - Cano entupido",
        "C - Excesso de Bolhas",
        "D - Prob. na caixa do Relogio ou Disjuntor",
        "E - Furos Obstruidos ou Faltando",
        "F - Identificação do poste (Carimbos)",
        "G - Presença de Trincas",
        "H - Bolha nas caixas",
        "I - Trincas em Toda a extenção do poste",
        "J - Pequenas Avarias",
        "K - Manchas Brancas superficiais",
        "L - Bucha das caixas e/ou bucha fixação da abraçadeira poste EDP",
        "M - Parafuso lacre da caixa do medidor",
        "N - Checagem de todos os itens pré-moldagem",
    ]

    meses_abas = [aba for aba, _ in MESES_2026]
    meses_labels = [lbl for _, lbl in MESES_2026]

    resumo = {
        "categorias": [],
        "meses_labels": meses_labels,
        "meses_abas": meses_abas,
        "retrabalho_mensal": [],
        "producao_mensal": [],
        "pct_mensal": [],
    }

    # Para cada categoria, contar ocorrências em cada mês
    for cat in CATEGORIAS:
        prefixo = cat.split(" - ")[0].strip().upper()  # "A", "B", etc.
        valores = []
        for aba in meses_abas:
            if aba in todos_resultados:
                pareto = todos_resultados[aba].get("erros_pareto", pd.DataFrame())
                if not pareto.empty:
                    # Match by prefix letter
                    match = pareto[pareto["Tipo de Erro"].str.strip().str.upper().str.startswith(prefixo + " ")]
                    freq = int(match["Frequência"].sum()) if not match.empty else 0
                    valores.append(freq)
                else:
                    valores.append(0)
            else:
                valores.append(None)  # mês sem dados
        resumo["categorias"].append({"nome": cat, "prefixo": prefixo, "valores": valores})

    # Retrabalho, produção e % por mês
    for aba in meses_abas:
        if aba in todos_resultados:
            r = todos_resultados[aba]
            ret = r.get("retrabalho", None)
            total = r.get("total", 0)
            prod = PRODUCAO.get(aba, {})
            prod_total = sum(prod.values())
            resumo["retrabalho_mensal"].append(total)
            resumo["producao_mensal"].append(prod_total)
            if prod_total > 0:
                resumo["pct_mensal"].append(round(total / prod_total * 100, 1))
            else:
                resumo["pct_mensal"].append(None)
        else:
            resumo["retrabalho_mensal"].append(None)
            resumo["producao_mensal"].append(None)
            resumo["pct_mensal"].append(None)

    return resumo


# ─── HTML helpers ─────────────────────────────────────────────────────────────

def esc(text):
    return html_module.escape(str(text))


def trend_icon(val):
    if val > 0:
        return f'<span class="trend-up">▲ +{val}</span>'
    elif val < 0:
        return f'<span class="trend-down">▼ {val}</span>'
    return '<span class="trend-flat">● 0</span>'


def trend_pct_icon(val):
    if val > 0:
        return f'<span class="trend-up">+{val}%</span>'
    elif val < 0:
        return f'<span class="trend-down">{val}%</span>'
    return '<span class="trend-flat">0%</span>'


def bar_html(items, label_col, val_col, colors=None):
    if items.empty:
        return ""
    max_v = items[val_col].max()
    default_colors = ["#3498db", "#e74c3c", "#2ecc71", "#f39c12", "#9b59b6",
                      "#1abc9c", "#e67e22", "#34495e", "#e91e63", "#00bcd4"]
    html_out = '<div class="bar-chart">'
    for idx, (_, row) in enumerate(items.head(12).iterrows()):
        pct = (row[val_col] / max_v * 100) if max_v > 0 else 0
        cor = (colors[idx % len(colors)] if colors else default_colors[idx % len(default_colors)])
        nome = esc(str(row[label_col])[:35])
        html_out += f'''
            <div class="bar-row">
                <div class="bar-label" title="{esc(row[label_col])}">{nome}</div>
                <div class="bar-track">
                    <div class="bar-fill" style="width:{pct:.1f}%;background:{cor};">{row[val_col]}</div>
                </div>
            </div>'''
    html_out += '</div>'
    return html_out


# ─── Geração do HTML ─────────────────────────────────────────────────────────

def gerar_html(todos_resultados, abas_disponiveis, comp, planos, resumo_2026, resumo_2025):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    # Últimos dois meses (usar mesmos meses do comp, não abas_disponiveis)
    aba_atu = comp["aba_atu"]
    lbl_atu = comp["label_atu"]
    lbl_ant = comp["label_ant"]
    atual = todos_resultados[aba_atu]
    prod = comp["producao"]
    prod_total = comp["producao_total"]

    # Última data registrada na lista de defeitos do mês atual
    ultima_data_defeitos = "N/D"
    col_data = atual.get("mapeamento", {}).get("data")
    df_atual = atual.get("df", pd.DataFrame())
    if col_data and col_data in df_atual.columns:
        serie_datas = df_atual[col_data].astype(str).str.strip()
        serie_datas = serie_datas[serie_datas != ""]
        if not serie_datas.empty:
            datas_parseadas = pd.to_datetime(serie_datas, dayfirst=True, errors="coerce")
            datas_parseadas = datas_parseadas.dropna()
            if not datas_parseadas.empty:
                ultima_data_defeitos = datas_parseadas.max().strftime("%d/%m/%y")

    # ── Tabela evolução mensal ──
    tab_evolucao = '<div class="table-scroll"><table class="data-table"><thead><tr><th>Mês</th><th>Defeitos</th><th>Variação</th><th>Produção</th><th>Taxa Defeito</th></tr></thead><tbody>'
    prev_total = None
    for aba in abas_disponiveis:
        r = todos_resultados[aba]
        p = PRODUCAO.get(aba, {})
        pt = sum(p.values())
        taxa = f"{r['total']/pt*100:.2f}%" if pt > 0 else "—"
        prod_str = f"{pt:,}" if pt > 0 else "—"
        if prev_total is not None:
            var = r["total"] - prev_total
            var_html = trend_icon(var)
        else:
            var_html = "—"
        tab_evolucao += f'<tr><td><strong>{r["mes"]}</strong></td><td>{r["total"]}</td><td>{var_html}</td><td>{prod_str}</td><td>{taxa}</td></tr>'
        prev_total = r["total"]
    tab_evolucao += '</tbody></table></div>'

    # ── Tabela comparativa de erros (últimos 2 meses) ──
    tab_erros_comp = ""
    if not comp["erros_comp"].empty:
        tab_erros_comp = f'<table class="data-table"><thead><tr><th>Tipo de Defeito</th><th>{lbl_ant}</th><th>{lbl_atu}</th><th>Variação</th><th>Var %</th><th>Tendência</th></tr></thead><tbody>'
        for _, row in comp["erros_comp"].iterrows():
            tab_erros_comp += f'''<tr>
                <td>{esc(row["Tipo de Erro"])}</td>
                <td>{row[lbl_ant]}</td><td>{row[lbl_atu]}</td>
                <td>{trend_icon(row["Variação"])}</td>
                <td>{trend_pct_icon(row["Var %"])}</td>
                <td>{"🔴 Piorou" if row["Variação"] > 0 else ("🟢 Melhorou" if row["Variação"] < 0 else "⚪ Estável")}</td>
            </tr>'''
        tab_erros_comp += '</tbody></table>'

    # ── Tabela comparativa setores ──
    tab_setores_comp = ""
    if not comp["setores_comp"].empty:
        tab_setores_comp = f'<table class="data-table"><thead><tr><th>Setor</th><th>{lbl_ant}</th><th>{lbl_atu}</th><th>Variação</th><th>Produção {lbl_atu}</th><th>Taxa Defeito (%)</th></tr></thead><tbody>'
        for _, row in comp["setores_comp"].iterrows():
            taxa = row["Taxa Defeito (%)"]
            taxa_cls = "heat-high" if taxa > 20 else ("heat-med" if taxa > 10 else "")
            prod_v = int(row["Produção"])
            prod_str = f"{prod_v:,}" if prod_v > 0 else "—"
            taxa_str = f"{taxa:.2f}%" if prod_v > 0 else "—"
            tab_setores_comp += f'''<tr>
                <td>{esc(row["Setor"])}</td>
                <td>{row[lbl_ant]}</td><td>{row[lbl_atu]}</td>
                <td>{trend_icon(row["Variação"])}</td>
                <td>{prod_str}</td>
                <td class="{taxa_cls}"><strong>{taxa_str}</strong></td>
            </tr>'''
        tab_setores_comp += '</tbody></table>'

    # ── Tabela % Erros por Setor — Produzido vs Defeito ──
    tab_setor_prod_vs_def = ""
    if not comp["setores_comp"].empty and prod_total > 0:
        total_def = comp["setores_comp"][lbl_atu].sum()
        tab_setor_prod_vs_def = f'''<table class="data-table"><thead><tr>
            <th>Setor</th><th>Produção {lbl_atu}</th><th>Defeitos {lbl_atu}</th>
            <th>Taxa de Defeito (%)</th><th>% dos Defeitos Totais</th><th>% da Produção Total</th>
        </tr></thead><tbody>'''
        for _, row in comp["setores_comp"].iterrows():
            p = int(row["Produção"])
            d = int(row[lbl_atu])
            taxa = round(d / p * 100, 2) if p > 0 else 0
            pct_def = round(d / total_def * 100, 1) if total_def > 0 else 0
            pct_prod = round(p / prod_total * 100, 1) if prod_total > 0 else 0
            taxa_cls = "heat-high" if taxa > 20 else ("heat-med" if taxa > 10 else "")
            p_str = f"{p:,}" if p > 0 else "—"
            taxa_str = f"{taxa:.2f}%" if p > 0 else "—"
            tab_setor_prod_vs_def += f'''<tr>
                <td><strong>{esc(row["Setor"])}</strong></td>
                <td>{p_str}</td><td>{d}</td>
                <td class="{taxa_cls}"><strong>{taxa_str}</strong></td>
                <td>{pct_def:.1f}%</td><td>{pct_prod:.1f}%</td>
            </tr>'''
        total_taxa = round(total_def / prod_total * 100, 2)
        tab_setor_prod_vs_def += f'''<tr style="font-weight:700;background:#edf2f7">
            <td>TOTAL</td><td>{prod_total:,}</td><td>{int(total_def)}</td>
            <td><strong>{total_taxa:.2f}%</strong></td><td>100%</td><td>100%</td></tr>'''
        tab_setor_prod_vs_def += '</tbody></table>'
    elif not comp["setores_comp"].empty:
        tab_setor_prod_vs_def = '<p class="section-desc" style="color:#e53e3e;">⚠️ Dados de produção não disponíveis para este mês. Preencha PRODUCAO_RAW no script.</p>'

    # ── Pareto mês atual ──
    pareto_atu_bar = ""
    if not atual["erros_pareto"].empty:
        cores = []
        for _, row in atual["erros_pareto"].head(12).iterrows():
            p = row["% do Total"]
            cores.append("#e74c3c" if p > 20 else ("#f39c12" if p > 10 else ("#3498db" if p > 5 else "#2ecc71")))
        pareto_atu_bar = bar_html(atual["erros_pareto"].head(12), "Tipo de Erro", "Frequência", cores)

    # ── Pareto mês anterior ──
    aba_ant = comp["aba_ant"]
    anterior = todos_resultados[aba_ant]
    pareto_ant_bar = ""
    if not anterior["erros_pareto"].empty:
        pareto_ant_bar = bar_html(anterior["erros_pareto"].head(12), "Tipo de Erro", "Frequência")

    # ── Como ocorreu ──
    tab_como_comp = ""
    if not comp.get("como_comp", pd.DataFrame()).empty:
        tab_como_comp = f'<table class="data-table"><thead><tr><th>Como Ocorreu</th><th>{lbl_ant}</th><th>{lbl_atu}</th><th>Variação</th></tr></thead><tbody>'
        for _, row in comp["como_comp"].iterrows():
            tab_como_comp += f'<tr><td>{esc(row["Como Ocorreu"])}</td><td>{row[lbl_ant]}</td><td>{row[lbl_atu]}</td><td>{trend_icon(row["Variação"])}</td></tr>'
        tab_como_comp += '</tbody></table>'

    # ── Retrabalho ──
    retrabalho_html = ""
    if "retrabalho_ant" in comp and "retrabalho_atu" in comp:
        ra = comp["retrabalho_ant"]
        rb = comp["retrabalho_atu"]
        taxa_var = round(rb["taxa"] - ra["taxa"], 1)
        cor_a = "#38a169" if ra["taxa"] >= 80 else ("#d69e2e" if ra["taxa"] >= 50 else "#e53e3e")
        cor_b = "#38a169" if rb["taxa"] >= 80 else ("#d69e2e" if rb["taxa"] >= 50 else "#e53e3e")
        retrabalho_html = f'''
        <div class="kpi-grid" style="grid-template-columns:repeat(5,1fr)">
            <div class="kpi-card" style="border-top-color:{cor_a}"><div class="kpi-value" style="color:{cor_a}">{ra["taxa"]}%</div><div class="kpi-label">Taxa Retrab. {lbl_ant}</div></div>
            <div class="kpi-card" style="border-top-color:{cor_b}"><div class="kpi-value" style="color:{cor_b}">{rb["taxa"]}%</div><div class="kpi-label">Taxa Retrab. {lbl_atu}</div></div>
            <div class="kpi-card"><div class="kpi-value">{ra["sim"]}/{ra["total"]}</div><div class="kpi-label">Retrab. {lbl_ant} (S/Total)</div></div>
            <div class="kpi-card"><div class="kpi-value">{rb["sim"]}/{rb["total"]}</div><div class="kpi-label">Retrab. {lbl_atu} (S/Total)</div></div>
            <div class="kpi-card"><div class="kpi-value">{trend_pct_icon(taxa_var)}</div><div class="kpi-label">Variação Taxa</div></div>
        </div>'''

    # ── Vitais ──
    vitais_html = ""
    if "erros_vitais" in atual and not atual["erros_vitais"].empty:
        v = atual["erros_vitais"]
        vitais_html = f'''<div class="pareto-highlight">
            <h3>🎯 Pareto {lbl_atu} — Poucos Vitais (80/20)</h3>
            <p><strong>{len(v)} tipo(s) de defeito</strong> representam ~80% das ocorrências.</p><ul>'''
        for _, row in v.iterrows():
            vitais_html += f'<li><strong>{esc(row["Tipo de Erro"])}</strong> — {row["Frequência"]} ({row["% do Total"]:.1f}%)</li>'
        vitais_html += '</ul></div>'

    # ── Cross erro x setor ──
    cross_html = ""
    if "cross_erro_setor" in atual:
        cross = atual["cross_erro_setor"]
        cross_html = '<div class="table-scroll"><table class="data-table cross-table"><thead><tr><th>Defeito \\ Setor</th>'
        for col in cross.columns:
            cross_html += f"<th>{esc(col)}</th>"
        cross_html += "<th>Total</th></tr></thead><tbody>"
        for idx_name, row in cross.iterrows():
            if str(idx_name).strip() == "":
                continue
            total_row = row.sum()
            cross_html += f"<tr><td><strong>{esc(idx_name)}</strong></td>"
            for val in row:
                heat = "heat-high" if val > 0 and val >= cross.values.max() * 0.7 else (
                    "heat-med" if val > 0 and val >= cross.values.max() * 0.3 else (
                    "heat-low" if val > 0 else ""))
                cross_html += f'<td class="{heat}">{val}</td>'
            cross_html += f'<td><strong>{total_row}</strong></td></tr>'
        cross_html += "</tbody></table></div>"

    # ── Formas ──
    formas_html = ""
    if not atual["formas"].empty:
        formas_html = f'<table class="data-table"><thead><tr><th>Nº Forma</th><th>Defeitos {lbl_atu}</th></tr></thead><tbody>'
        for _, row in atual["formas"].head(15).iterrows():
            formas_html += f'<tr><td>{esc(row["Nº Forma"])}</td><td>{row["Frequência"]}</td></tr>'
        formas_html += '</tbody></table>'

    # ── Produção por setor ──
    producao_html = ""
    if prod_total > 0:
        producao_html = f'<table class="data-table"><thead><tr><th>Setor</th><th>Produção {lbl_atu}</th><th>% da Produção</th></tr></thead><tbody>'
        for setor, qtd in sorted(prod.items()):
            pct = round(qtd / prod_total * 100, 1)
            producao_html += f'<tr><td>{esc(setor)}</td><td>{qtd:,}</td><td>{pct}%</td></tr>'
        producao_html += f'<tr style="font-weight:700;background:#edf2f7"><td>TOTAL</td><td>{prod_total:,}</td><td>100%</td></tr></tbody></table>'
    else:
        producao_html = '<p class="section-desc" style="color:#e53e3e;">⚠️ Dados de produção não disponíveis. Preencha PRODUCAO_RAW no script.</p>'

    # ── Planos ──
    html_planos = ""
    for p in planos:
        pr_cls = {"CRÍTICA": "critico", "ALTA": "alto", "MÉDIA": "medio", "BAIXA": "baixo"}.get(p["prioridade"], "baixo")
        tend_html = f'<div class="tendencia-badge">{p["tendencia"]}</div>' if p["tendencia"] else ""
        html_planos += f'''
        <div class="plano-card {pr_cls}-border">
            <div class="plano-header">
                <h3>🔧 {esc(p["erro"])}</h3>
                <div class="plano-meta">
                    <span class="badge {pr_cls}">{p["prioridade"]}</span>
                    <span class="meta-info">{p["frequencia"]} ocorrências ({p["percentual"]:.1f}%)</span>
                </div>
            </div>
            {tend_html}
            <div class="dmaic-grid">
                <div class="dmaic-step define"><h4>📋 DEFINE</h4><p>{p["dmaic"]["Define"]}</p></div>
                <div class="dmaic-step measure"><h4>📏 MEASURE</h4><p>{p["dmaic"]["Measure"]}</p></div>
                <div class="dmaic-step analyze"><h4>🔍 ANALYZE</h4><p>{p["dmaic"]["Analyze"]}</p></div>
                <div class="dmaic-step improve"><h4>⚡ IMPROVE</h4><p>{p["dmaic"]["Improve"]}</p></div>
                <div class="dmaic-step control"><h4>🎯 CONTROL</h4><p>{p["dmaic"]["Control"]}</p></div>
            </div>
            <div class="wald-section">
                <h4>🛩️ Viés de Sobrevivência (Abraham Wald)</h4>
                <h5>Perguntas Críticas:</h5>
                <ul class="wald-questions">'''
        for q in p["wald"]["perguntas"]:
            html_planos += f"<li>{q}</li>"
        html_planos += '</ul><h5>Recomendações:</h5><ul class="wald-recs">'
        for r in p["wald"]["recomendacoes"]:
            html_planos += f"<li>{esc(r)}</li>"
        html_planos += '</ul></div></div>'

    # ── Indicadores ──
    erros_novos = 0
    erros_eliminados = 0
    if not comp["erros_comp"].empty:
        erros_novos = len(comp["erros_comp"][(comp["erros_comp"][lbl_ant] == 0) & (comp["erros_comp"][lbl_atu] > 0)])
        erros_eliminados = len(comp["erros_comp"][(comp["erros_comp"][lbl_ant] > 0) & (comp["erros_comp"][lbl_atu] == 0)])

    var_cor = "#e53e3e" if comp["variacao_pct"] > 0 else ("#38a169" if comp["variacao_pct"] < 0 else "#718096")
    var_seta = "▲" if comp["variacao_pct"] > 0 else ("▼" if comp["variacao_pct"] < 0 else "●")

    taxa_geral = comp["taxa_defeito_geral"]
    taxa_cor = '#e53e3e' if taxa_geral > 15 else '#f39c12' if taxa_geral > 5 else '#38a169'
    taxa_str = f"{taxa_geral}%" if prod_total > 0 else "N/D"

    # Lista de meses no header
    meses_str = " → ".join(todos_resultados[a]["mes"] for a in abas_disponiveis)
    prod_badge = f'🏭 Produção {lbl_atu}: {prod_total:,} un.' if prod_total > 0 else ""

    # ── Tabela evolução por tipo de erro (todos os meses) ──
    tab_evolucao_erros = ""
    if len(abas_disponiveis) >= 2:
        all_erros = set()
        for aba in abas_disponiveis:
            if not todos_resultados[aba]["erros_pareto"].empty:
                all_erros.update(todos_resultados[aba]["erros_pareto"]["Tipo de Erro"].tolist())
        if all_erros:
            labels = [todos_resultados[a]["mes"] for a in abas_disponiveis]
            tab_evolucao_erros = '<div class="table-scroll"><table class="data-table"><thead><tr><th>Tipo de Defeito</th>'
            for lbl in labels:
                tab_evolucao_erros += f"<th>{lbl}</th>"
            tab_evolucao_erros += '<th>Tendência</th></tr></thead><tbody>'
            # Ordenar por último mês
            ultimo = abas_disponiveis[-1]
            pareto_ultimo = todos_resultados[ultimo]["erros_pareto"]
            erros_ord = pareto_ultimo["Tipo de Erro"].tolist() if not pareto_ultimo.empty else sorted(all_erros)
            outros = sorted(all_erros - set(erros_ord))
            erros_ord += outros
            for erro in erros_ord:
                vals = []
                for aba in abas_disponiveis:
                    par = todos_resultados[aba]["erros_pareto"]
                    if not par.empty:
                        m = par[par["Tipo de Erro"] == erro]
                        vals.append(int(m["Frequência"].iloc[0]) if not m.empty else 0)
                    else:
                        vals.append(0)
                tab_evolucao_erros += f"<tr><td>{esc(erro)}</td>"
                for v in vals:
                    tab_evolucao_erros += f"<td>{v}</td>"
                if len(vals) >= 2:
                    diff = vals[-1] - vals[-2]
                    tab_evolucao_erros += f"<td>{trend_icon(diff)}</td>"
                else:
                    tab_evolucao_erros += "<td>—</td>"
                tab_evolucao_erros += "</tr>"
            tab_evolucao_erros += '</tbody></table></div>'

    # ── Tabela Resumo 2026 (estilo Resumo da planilha) ──
    resumo_tab1_html = ""
    resumo_tab2_html = ""
    resumo_comp_html = ""

    if resumo_2026:
        meses_lbl = resumo_2026["meses_labels"]  # 12 labels
        meses_abas = resumo_2026["meses_abas"]

        # ── Tabela 1: Tipo de Defeito × Meses ──
        resumo_tab1_html = '<div class="table-scroll"><table class="resumo-table"><thead><tr><th class="resumo-cat">Tipo de Defeito</th>'
        for lbl in meses_lbl:
            resumo_tab1_html += f'<th>{lbl}</th>'
        resumo_tab1_html += '<th>Total</th><th>% Total</th></tr></thead><tbody>'

        grand_total = 0
        cat_totals = []
        for cat in resumo_2026["categorias"]:
            vals = cat["valores"]
            total_cat = sum(v for v in vals if v is not None and v > 0)
            cat_totals.append(total_cat)
            grand_total += total_cat

        for idx, cat in enumerate(resumo_2026["categorias"]):
            vals = cat["valores"]
            total_cat = cat_totals[idx]
            pct_total = round(total_cat / grand_total * 100, 1) if grand_total > 0 else 0
            nome = esc(cat["nome"])

            # Row color based on magnitude
            resumo_tab1_html += f'<tr><td class="resumo-cat"><strong>{nome}</strong></td>'
            for v in vals:
                if v is None:
                    resumo_tab1_html += '<td class="resumo-empty">—</td>'
                elif v == 0:
                    resumo_tab1_html += '<td class="resumo-zero">0</td>'
                else:
                    heat = "resumo-high" if v >= 50 else ("resumo-med" if v >= 20 else ("resumo-low" if v >= 5 else ""))
                    resumo_tab1_html += f'<td class="{heat}">{v}</td>'
            pct_cls = "resumo-high" if pct_total > 20 else ("resumo-med" if pct_total > 10 else "")
            resumo_tab1_html += f'<td class="resumo-total"><strong>{total_cat}</strong></td>'
            resumo_tab1_html += f'<td class="{pct_cls}"><strong>{pct_total}%</strong></td></tr>'

        # Totals row
        resumo_tab1_html += '<tr class="resumo-grand-total"><td class="resumo-cat"><strong>TOTAL GERAL</strong></td>'
        for i in range(12):
            col_sum = sum(c["valores"][i] for c in resumo_2026["categorias"] if c["valores"][i] is not None)
            if any(c["valores"][i] is not None for c in resumo_2026["categorias"]):
                resumo_tab1_html += f'<td><strong>{col_sum}</strong></td>'
            else:
                resumo_tab1_html += '<td class="resumo-empty">—</td>'
        resumo_tab1_html += f'<td><strong>{grand_total}</strong></td><td><strong>100%</strong></td></tr>'
        resumo_tab1_html += '</tbody></table></div>'

        # ── Tabela 2: Produção × Retrabalho × % × Meta ──
        ret_m = resumo_2026["retrabalho_mensal"]
        prod_m = resumo_2026["producao_mensal"]
        pct_m = resumo_2026["pct_mensal"]

        resumo_tab2_html = '<div class="table-scroll"><table class="resumo-table"><thead><tr><th class="resumo-cat">Indicador</th>'
        for lbl in meses_lbl:
            resumo_tab2_html += f'<th>{lbl}</th>'
        resumo_tab2_html += '<th>Média 2026</th><th>Média 2025</th><th>Média 2024</th></tr></thead><tbody>'

        # Row: Quantidade Total Defeitos
        resumo_tab2_html += '<tr><td class="resumo-cat"><strong>Quantidade Total Defeitos</strong></td>'
        valid_ret = [v for v in ret_m if v is not None and v > 0]
        for v in ret_m:
            if v is None:
                resumo_tab2_html += '<td class="resumo-empty">—</td>'
            else:
                resumo_tab2_html += f'<td>{v}</td>'
        media_ret = round(sum(valid_ret) / len(valid_ret), 0) if valid_ret else 0
        resumo_tab2_html += f'<td class="resumo-total"><strong>{int(media_ret)}</strong></td>'
        resumo_tab2_html += '<td>—</td><td>—</td></tr>'

        # Row: Produção Total
        resumo_tab2_html += '<tr><td class="resumo-cat"><strong>Produção Total (peças)</strong></td>'
        valid_prod = [v for v in prod_m if v is not None and v > 0]
        for v in prod_m:
            if v is None:
                resumo_tab2_html += '<td class="resumo-empty">—</td>'
            elif v == 0:
                resumo_tab2_html += '<td class="resumo-empty">—</td>'
            else:
                resumo_tab2_html += f'<td>{int(v):,}</td>'
        media_prod = round(sum(valid_prod) / len(valid_prod), 0) if valid_prod else 0
        resumo_tab2_html += f'<td class="resumo-total"><strong>{int(media_prod):,}</strong></td>'
        resumo_tab2_html += '<td>—</td><td>—</td></tr>'

        # Row: Percentual de Defeitos (%)
        resumo_tab2_html += '<tr class="resumo-pct-row"><td class="resumo-cat"><strong>Percentual de Defeitos (%)</strong></td>'
        valid_pct = [v for v in pct_m if v is not None]
        for v in pct_m:
            if v is None:
                resumo_tab2_html += '<td class="resumo-empty">—</td>'
            else:
                cls = "resumo-acima-meta" if v > META_RETRABALHO else "resumo-abaixo-meta"
                resumo_tab2_html += f'<td class="{cls}"><strong>{v}%</strong></td>'
        media_pct_2026 = round(sum(valid_pct) / len(valid_pct), 1) if valid_pct else 0
        media_cls_2026 = "resumo-acima-meta" if media_pct_2026 > META_RETRABALHO else "resumo-abaixo-meta"
        media_cls_2025 = "resumo-acima-meta" if MEDIA_2025 > META_RETRABALHO else "resumo-abaixo-meta"
        media_cls_2024 = "resumo-acima-meta" if MEDIA_2024 > META_RETRABALHO else "resumo-abaixo-meta"
        resumo_tab2_html += f'<td class="{media_cls_2026}"><strong>{media_pct_2026}%</strong></td>'
        resumo_tab2_html += f'<td class="{media_cls_2025}"><strong>{MEDIA_2025}%</strong></td>'
        resumo_tab2_html += f'<td class="{media_cls_2024}"><strong>{MEDIA_2024}%</strong></td></tr>'

        # Row: Meta
        resumo_tab2_html += '<tr class="resumo-meta-row"><td class="resumo-cat"><strong>Meta</strong></td>'
        for _ in range(12):
            resumo_tab2_html += f'<td>{META_RETRABALHO}%</td>'
        resumo_tab2_html += f'<td>{META_RETRABALHO}%</td><td>{META_RETRABALHO}%</td><td>{META_RETRABALHO}%</td></tr>'
        resumo_tab2_html += '</tbody></table></div>'

        # ── Comparação média mês corrente vs anos anteriores ──
        if valid_pct:
            ultimo_pct = valid_pct[-1]
            # Find which month is the latest
            ultimo_idx = None
            for i in range(11, -1, -1):
                if pct_m[i] is not None:
                    ultimo_idx = i
                    break
            ultimo_label = meses_lbl[ultimo_idx] if ultimo_idx is not None else "—"

            diff_2025 = round(ultimo_pct - MEDIA_2025, 1)
            diff_2024 = round(ultimo_pct - MEDIA_2024, 1)
            diff_media = round(media_pct_2026 - MEDIA_2025, 1)

            cor_vs_25 = "#38a169" if diff_2025 < 0 else ("#e53e3e" if diff_2025 > 0 else "#718096")
            cor_vs_24 = "#38a169" if diff_2024 < 0 else ("#e53e3e" if diff_2024 > 0 else "#718096")
            cor_media = "#38a169" if diff_media < 0 else ("#e53e3e" if diff_media > 0 else "#718096")

            seta_25 = "▼" if diff_2025 < 0 else ("▲" if diff_2025 > 0 else "●")
            seta_24 = "▼" if diff_2024 < 0 else ("▲" if diff_2024 > 0 else "●")
            seta_m = "▼" if diff_media < 0 else ("▲" if diff_media > 0 else "●")

            resumo_comp_html = f'''
            <div class="kpi-grid" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr));margin-top:20px">
                <div class="kpi-card" style="border-top-color:#3182ce">
                    <div class="kpi-value" style="color:#3182ce">{ultimo_pct}%</div>
                    <div class="kpi-label">Taxa Defeito {ultimo_label}</div>
                </div>
                <div class="kpi-card" style="border-top-color:{cor_vs_25}">
                    <div class="kpi-value" style="color:{cor_vs_25}">{seta_25} {diff_2025:+.1f}pp</div>
                    <div class="kpi-label">vs Média 2025 ({MEDIA_2025}%)</div>
                </div>
                <div class="kpi-card" style="border-top-color:{cor_vs_24}">
                    <div class="kpi-value" style="color:{cor_vs_24}">{seta_24} {diff_2024:+.1f}pp</div>
                    <div class="kpi-label">vs Média 2024 ({MEDIA_2024}%)</div>
                </div>
                <div class="kpi-card" style="border-top-color:{cor_media}">
                    <div class="kpi-value" style="color:{cor_media}">{media_pct_2026}%</div>
                    <div class="kpi-label">Média 2026 (acumulada)</div>
                </div>
                <div class="kpi-card" style="border-top-color:#805ad5">
                    <div class="kpi-value" style="color:#805ad5">{seta_m} {diff_media:+.1f}pp</div>
                    <div class="kpi-label">Média 2026 vs 2025</div>
                </div>
                <div class="kpi-card" style="border-top-color:#dd6b20">
                    <div class="kpi-value">{META_RETRABALHO}%</div>
                    <div class="kpi-label">Meta Retrabalho</div>
                </div>
            </div>'''

    # ─── HTML final ───
    html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Análise Comparativa 2026 — Lean Six Sigma + Wald</title>
<style>
:root {{
    --primary: #1a365d; --secondary: #2d3748; --accent: #3182ce;
    --success: #38a169; --warning: #d69e2e; --danger: #e53e3e;
    --bg: #f7fafc; --card-bg: #fff; --border: #e2e8f0;
    --text: #2d3748; --text-light: #718096;
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Segoe UI',system-ui,sans-serif; background:var(--bg); color:var(--text); line-height:1.6; }}
.container {{ max-width:1300px; margin:0 auto; padding:20px; }}
.report-header {{
    background:linear-gradient(135deg, #1a365d 0%, #2c5282 50%, #2b6cb0 100%);
    color:#fff; padding:40px; border-radius:12px; margin-bottom:30px;
    box-shadow:0 4px 15px rgba(0,0,0,0.2);
}}
.report-header h1 {{ font-size:2.2em; margin-bottom:6px; }}
.report-header .subtitle {{ opacity:0.9; font-size:1.1em; margin-bottom:15px; }}
.header-badges {{ display:flex; gap:12px; flex-wrap:wrap; }}
.header-badge {{ background:rgba(255,255,255,0.15); padding:6px 14px; border-radius:20px; font-size:0.85em; }}
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; margin-bottom:25px; }}
.kpi-card {{ background:var(--card-bg); border-radius:10px; padding:18px; text-align:center;
    box-shadow:0 2px 8px rgba(0,0,0,0.06); border-top:4px solid var(--accent); }}
.kpi-card .kpi-value {{ font-size:2em; font-weight:700; color:var(--primary); }}
.kpi-card .kpi-label {{ font-size:0.8em; color:var(--text-light); text-transform:uppercase; letter-spacing:1px; }}
.report-section {{ background:var(--card-bg); border-radius:10px; padding:28px; margin-bottom:22px;
    box-shadow:0 2px 8px rgba(0,0,0,0.06); }}
.report-section h2 {{ color:var(--primary); border-bottom:2px solid var(--border); padding-bottom:10px; margin-bottom:18px; font-size:1.3em; }}
.section-desc {{ color:var(--text-light); margin-bottom:12px; font-size:0.95em; }}
.data-table {{ width:100%; border-collapse:collapse; margin-top:12px; font-size:0.9em; }}
.data-table th {{ background:var(--primary); color:#fff; padding:10px 14px; text-align:center; font-weight:600; font-size:0.82em; text-transform:uppercase; letter-spacing:0.5px; }}
.data-table td {{ padding:8px 14px; border-bottom:1px solid var(--border); text-align:center; }}
.data-table tr:hover {{ background:#edf2f7; }}
.table-scroll {{ overflow-x:auto; }}
.heat-high {{ background:#fed7d7 !important; color:#9b2c2c; font-weight:700; }}
.heat-med {{ background:#fefcbf !important; color:#975a16; }}
.heat-low {{ background:#f0fff4 !important; }}
.badge {{ display:inline-block; padding:3px 10px; border-radius:12px; font-size:0.72em; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; }}
.critico {{ background:#fed7d7; color:#9b2c2c; }} .alto {{ background:#feebc8; color:#975a16; }}
.medio {{ background:#fefcbf; color:#975a16; }} .baixo {{ background:#c6f6d5; color:#276749; }}
.bar-chart {{ display:flex; flex-direction:column; gap:6px; margin:15px 0; }}
.bar-row {{ display:flex; align-items:center; gap:8px; }}
.bar-label {{ min-width:180px; max-width:180px; font-size:0.82em; text-align:right; overflow:hidden;
    text-overflow:ellipsis; white-space:nowrap; color:var(--secondary); }}
.bar-track {{ flex:1; background:#edf2f7; border-radius:6px; height:28px; overflow:hidden; }}
.bar-fill {{ height:100%; border-radius:6px; display:flex; align-items:center; padding:0 8px;
    color:#fff; font-size:0.78em; font-weight:600; white-space:nowrap; min-width:fit-content; }}
.pareto-highlight {{ background:linear-gradient(135deg,#ebf8ff,#e6fffa); border:2px solid #3182ce;
    border-radius:10px; padding:22px; margin:18px 0; }}
.pareto-highlight h3 {{ color:var(--primary); margin-bottom:8px; }}
.pareto-highlight ul {{ margin-top:8px; padding-left:18px; }}
.plano-card {{ background:var(--card-bg); border-radius:10px; padding:22px; margin-bottom:22px;
    box-shadow:0 2px 10px rgba(0,0,0,0.08); border-left:5px solid var(--accent); }}
.critico-border {{ border-left-color:var(--danger); }} .alto-border {{ border-left-color:#dd6b20; }}
.medio-border {{ border-left-color:var(--warning); }} .baixo-border {{ border-left-color:var(--success); }}
.plano-header {{ display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:8px; margin-bottom:15px; }}
.plano-header h3 {{ color:var(--primary); font-size:1.2em; }}
.plano-meta {{ display:flex; align-items:center; gap:8px; }}
.meta-info {{ color:var(--text-light); font-size:0.85em; }}
.tendencia-badge {{ background:#edf2f7; padding:8px 15px; border-radius:6px; margin-bottom:15px; font-size:0.9em; font-weight:600; }}
.dmaic-grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:12px; margin-bottom:18px; }}
.dmaic-step {{ padding:14px; border-radius:8px; border:1px solid var(--border); }}
.dmaic-step h4 {{ margin-bottom:6px; font-size:0.9em; }}
.dmaic-step p {{ font-size:0.82em; color:var(--secondary); }}
.define {{ background:#ebf8ff; border-color:#90cdf4; }}
.measure {{ background:#e9d8fd; border-color:#b794f4; }}
.analyze {{ background:#fefcbf; border-color:#f6e05e; }}
.improve {{ background:#c6f6d5; border-color:#68d391; }}
.control {{ background:#fed7e2; border-color:#feb2b2; }}
.wald-section {{ background:#fffbeb; border:1px solid #f6e05e; border-radius:8px; padding:18px; margin-top:12px; }}
.wald-section h4 {{ color:#975a16; margin-bottom:10px; }}
.wald-section h5 {{ color:#975a16; margin:12px 0 6px; }}
.wald-questions li {{ margin-bottom:6px; color:#92400e; font-size:0.85em; }}
.wald-recs li {{ margin-bottom:6px; color:#065f46; font-size:0.85em; }}
.wald-questions, .wald-recs {{ padding-left:18px; }}
.trend-up {{ color:#e53e3e; font-weight:700; }}
.trend-down {{ color:#38a169; font-weight:700; }}
.trend-flat {{ color:#718096; font-weight:600; }}
.cross-table th {{ font-size:0.72em; padding:6px; text-align:center; }}
.cross-table td {{ padding:6px; font-size:0.82em; text-align:center; }}
.comp-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; }}
@media (max-width:900px) {{ .comp-grid {{ grid-template-columns:1fr; }} }}

/* Resumo tables */
.resumo-table {{ width:100%; border-collapse:collapse; margin-top:12px; font-size:0.82em; }}
.resumo-table th {{ background:linear-gradient(135deg, #1a365d, #2c5282); color:#fff; padding:8px 6px;
    text-align:center; font-weight:600; font-size:0.78em; text-transform:uppercase; letter-spacing:0.3px;
    position:sticky; top:0; }}
.resumo-table td {{ padding:7px 6px; border-bottom:1px solid #e2e8f0; text-align:center; font-size:0.88em; }}
.resumo-table tr:nth-child(even) {{ background:#f7fafc; }}
.resumo-table tr:hover {{ background:#edf2f7; }}
.resumo-cat {{ text-align:left !important; min-width:200px; white-space:nowrap; }}
.resumo-empty {{ color:#cbd5e0; font-style:italic; }}
.resumo-zero {{ color:#a0aec0; }}
.resumo-high {{ background:#fed7d7 !important; color:#9b2c2c; font-weight:700; }}
.resumo-med {{ background:#fefcbf !important; color:#975a16; font-weight:600; }}
.resumo-low {{ background:#fef3c7 !important; color:#92400e; }}
.resumo-total {{ background:#edf2f7 !important; font-weight:700; }}
.resumo-grand-total td {{ background:#2d3748 !important; color:#fff !important; font-weight:700; font-size:0.9em; }}
.resumo-acima-meta {{ background:linear-gradient(135deg, #fed7d7, #feb2b2) !important; color:#9b2c2c; font-weight:700; }}
.resumo-abaixo-meta {{ background:linear-gradient(135deg, #c6f6d5, #9ae6b4) !important; color:#276749; font-weight:700; }}
.resumo-meta-row td {{ background:#ebf8ff !important; color:#2c5282; font-weight:600; font-style:italic; }}
.resumo-pct-row td {{ font-size:0.92em; }}
.resumo-section-header {{ display:flex; align-items:center; gap:12px; margin-bottom:8px; }}
.resumo-section-header .badge-year {{ display:inline-block; padding:4px 12px; border-radius:16px;
    font-size:0.72em; font-weight:700; text-transform:uppercase; letter-spacing:0.5px; }}
.badge-2024 {{ background:#e9d8fd; color:#553c9a; }}
.badge-2025 {{ background:#feebc8; color:#975a16; }}
.badge-2026 {{ background:#c6f6d5; color:#276749; }}

.report-footer {{ text-align:center; padding:25px; color:var(--text-light); font-size:0.82em;
    border-top:1px solid var(--border); margin-top:25px; }}
@media print {{
    body {{ background:#fff; }}
    .report-header {{ -webkit-print-color-adjust:exact; }}
    .plano-card {{ break-inside:avoid; }}
}}
</style>
</head>
<body>
<div class="container">

<div class="report-header">
    <h1>Análise Comparativa de Defeitos — 2026</h1>
    <p class="subtitle">Master Black Belt · Lean Six Sigma (DMAIC) · Viés de Sobrevivência (Abraham Wald)</p>
    <div class="header-badges">
        <span class="header-badge">📅 {agora}</span>
        <span class="header-badge">🗓️ Última atualização da lista: {ultima_data_defeitos}</span>
        <span class="header-badge">📊 {meses_str}</span>
        {"<span class='header-badge'>" + prod_badge + "</span>" if prod_badge else ""}
        <span class="header-badge">🛩️ Análise de Wald</span>
    </div>
</div>

<!-- KPIs Executivos -->
<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-value">{anterior["total"]}</div>
        <div class="kpi-label">Defeitos {lbl_ant}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-value">{atual["total"]}</div>
        <div class="kpi-label">Defeitos {lbl_atu}</div>
    </div>
    <div class="kpi-card" style="border-top-color:{var_cor}">
        <div class="kpi-value" style="color:{var_cor}">{var_seta} {comp["variacao_pct"]}%</div>
        <div class="kpi-label">Variação {lbl_ant}→{lbl_atu}</div>
    </div>
    {"<div class='kpi-card'><div class='kpi-value'>" + f"{prod_total:,}" + "</div><div class='kpi-label'>Produção " + lbl_atu + "</div></div>" if prod_total > 0 else ""}
    <div class="kpi-card" style="border-top-color:{taxa_cor}">
        <div class="kpi-value">{taxa_str}</div>
        <div class="kpi-label">Taxa Defeito {lbl_atu}</div>
    </div>
    <div class="kpi-card" style="border-top-color:#e53e3e">
        <div class="kpi-value">{erros_novos}</div>
        <div class="kpi-label">Novos tipos em {lbl_atu}</div>
    </div>
    <div class="kpi-card" style="border-top-color:#38a169">
        <div class="kpi-value">{erros_eliminados}</div>
        <div class="kpi-label">Eliminados em {lbl_atu}</div>
    </div>
</div>

<!-- Evolução Mensal -->
<section class="report-section">
    <h2>📈 Evolução Mensal — {meses_str}</h2>
    <p class="section-desc">Resumo da evolução dos defeitos ao longo dos meses disponíveis.</p>
    {tab_evolucao}
</section>

<!-- Resumo 2026 — Tabela de Defeitos por Categoria -->
{"" if not resumo_tab1_html else f"""
<section class="report-section">
    <h2>📋 Resumo 2026 — Defeitos por Categoria</h2>
    <div class="resumo-section-header">
        <p class="section-desc" style="margin:0">Quantidade de defeitos por tipo (categorias A-N) em cada mês de 2026.
        Células coloridas indicam concentração: <span class="badge critico" style="font-size:0.7em">≥50</span>
        <span class="badge alto" style="font-size:0.7em">≥20</span>
        <span class="badge medio" style="font-size:0.7em">≥5</span></p>
    </div>
    {resumo_tab1_html}
</section>"""}

<!-- Resumo 2026 — Produção vs Defeitos vs Meta -->
{"" if not resumo_tab2_html else f"""
<section class="report-section">
    <h2>🏭 Resumo 2026 — Produção × Defeitos × Meta</h2>
    <div class="resumo-section-header">
        <p class="section-desc" style="margin:0">Comparativo mensal de produção, defeitos e taxa percentual.
        <span class="badge-year badge-2024">Média 2024: {MEDIA_2024}%</span>
        <span class="badge-year badge-2025">Média 2025: {MEDIA_2025}%</span>
        <span class="badge-year badge-2026">Meta: {META_RETRABALHO}%</span></p>
    </div>
    {resumo_tab2_html}
    {resumo_comp_html}
</section>"""}

<!-- Evolução por Tipo de Erro -->
{"" if not tab_evolucao_erros else f"""
<section class="report-section">
    <h2>📊 Evolução por Tipo de Defeito — Todos os Meses</h2>
    <p class="section-desc">Quantidade de cada tipo de defeito mês a mês. Tendência compara os dois últimos meses.</p>
    {tab_evolucao_erros}
</section>"""}

<!-- Comparativo de Erros (2 últimos) -->
<section class="report-section">
    <h2>📊 Comparativo de Defeitos — {lbl_ant} vs {lbl_atu}</h2>
    <p class="section-desc">Evolução entre os dois últimos meses. Setas vermelhas indicam piora, verdes indicam melhora.</p>
    <div class="comp-grid">
        <div><h3 style="color:var(--text-light);margin-bottom:10px;">{lbl_ant}</h3>{pareto_ant_bar}</div>
        <div><h3 style="color:var(--text-light);margin-bottom:10px;">{lbl_atu}</h3>{pareto_atu_bar}</div>
    </div>
    {tab_erros_comp}
</section>

<!-- Pareto -->
<section class="report-section">
    <h2>🎯 Análise de Pareto — {lbl_atu}</h2>
    {vitais_html}
</section>

<!-- Produção por Setor -->
<section class="report-section">
    <h2>🏭 Produção por Setor — {lbl_atu}</h2>
    <p class="section-desc">Volume produzido por setor.</p>
    {producao_html}
</section>

<!-- % Erros por Setor — Produzido vs Defeito -->
<section class="report-section">
    <h2>📍 % de Erros por Setor — {lbl_atu} (Produzido vs Defeito)</h2>
    <p class="section-desc">Volume produzido vs quantidade de defeitos por setor, com taxa percentual de defeito.</p>
    {tab_setor_prod_vs_def}
</section>

<!-- Comparativo Setores -->
<section class="report-section">
    <h2>📍 Comparativo por Setor — {lbl_ant} vs {lbl_atu}</h2>
    <p class="section-desc">Evolução dos defeitos por setor entre os dois últimos meses.</p>
    {tab_setores_comp}
</section>

<!-- Cross Erro x Setor -->
<section class="report-section">
    <h2>🔀 Matriz Defeito × Setor — {lbl_atu} (Heat Map)</h2>
    {cross_html}
</section>

<!-- Como Ocorreu -->
{"" if not tab_como_comp else f"""
<section class="report-section">
    <h2>❓ Como os Defeitos Ocorreram — {lbl_ant} vs {lbl_atu}</h2>
    {tab_como_comp}
</section>"""}

<!-- Retrabalho -->
<section class="report-section">
    <h2>🔄 Taxa de Retrabalho — {lbl_ant} vs {lbl_atu}</h2>
    {retrabalho_html}
</section>

<!-- Formas -->
{"" if not formas_html else f"""
<section class="report-section">
    <h2>🔩 Top Formas com Defeitos — {lbl_atu}</h2>
    <p class="section-desc">Moldes com maior concentração de defeitos.</p>
    {formas_html}
</section>"""}

<!-- Planos de Ação -->
<section class="report-section">
    <h2>🔧 Planos de Ação DMAIC + Análise de Wald — {lbl_atu}</h2>
    <p class="section-desc">Plano completo para os <strong>2 principais defeitos</strong> de {lbl_atu}. DMAIC detalhado + Análise de Viés de Sobrevivência.</p>
    {html_planos}
</section>

<!-- Wald Geral -->
<section class="report-section">
    <h2>🛩️ Análise Geral — O Olhar de Abraham Wald</h2>
    <div class="wald-section">
        <h4>O que os dados NÃO mostram é tão importante quanto o que mostram</h4>
        <p style="font-style:italic;color:#744210;margin:10px 0">
            Abraham Wald mostrou que reforçar as partes dos aviões SEM buracos salvava mais vidas —
            porque os aviões atingidos nessas partes não voltavam.
        </p>
        <ul class="wald-recs" style="margin-top:12px;">
            <li><strong>Comparação {lbl_ant} vs {lbl_atu}:</strong> Defeitos que DESAPARECERAM
            podem indicar melhoria real OU mudança no método de registro/inspeção.</li>
            <li><strong>Defeitos NOVOS em {lbl_atu}:</strong> {erros_novos} tipo(s) aparecem pela primeira vez.
            São realmente novos ou antes eram classificados diferentemente?</li>
            <li><strong>Setores com taxa zero:</strong> Se um setor produziu e não tem defeitos,
            verificar se a inspeção realmente cobriu toda a produção.</li>
            <li><strong>Taxa de defeito vs produção:</strong> {f"Taxa de {taxa_geral}% sobre {prod_total:,} un." if prod_total > 0 else "Sem dados de produção disponíveis."}
            Compare com benchmarks do setor.</li>
            <li><strong>Os dados de campo:</strong> Reclamações, devoluções e quebras em serviço
            revelam os defeitos que escaparam — os "aviões que não voltaram".</li>
        </ul>
    </div>
</section>

<div class="report-footer">
    <p>Relatório Comparativo 2026 · {meses_str} · Lean Six Sigma DMAIC · Abraham Wald</p>
    <p>Master Black Belt Analysis · {agora}</p>
</div>

</div>
</body>
</html>'''

    return html


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  ANÁLISE COMPARATIVA 2026  —  Lean Six Sigma + Wald")
    print("=" * 70)

    # Conectar às planilhas
    print("\n[INFO] Conectando ao Google Sheets...")
    creds = carregar_credenciais()
    
    # Planilha de produção
    print("[INFO] Acessando planilha de produção...")
    client_prod = gspread.authorize(creds)
    planilha_prod = client_prod.open_by_key(SPREADSHEET_ID)
    
    # Listar abas disponíveis na planilha de produção
    abas_prod = [ws.title for ws in planilha_prod.worksheets()]
    print(f"[INFO] Abas na planilha de produção: {abas_prod}")
    
    # Planilha de defeitos
    print("[INFO] Acessando planilha de defeitos...")
    client_def = gspread.authorize(creds)
    planilha = client_def.open_by_key(SPREADSHEET_ID_DEFEITOS)

    # Ler produção de Março por setor em Estoque!C700:D704
    print("[INFO] Lendo produção setorial de Março (Estoque!C700:D704)...")
    producao_marco_setorial = ler_producao_setorial_marco(planilha_prod)

    if producao_marco_setorial:
        linhas_marco = [f"Produção Setorial|{setor}|{qtd}" for setor, qtd in producao_marco_setorial.items()]
        PRODUCAO_RAW["mar26"] = "\n".join(linhas_marco)
        print(f"[INFO] Produção Março por setor carregada: {producao_marco_setorial}")
    else:
        # Fallback: total em Estoque!C278
        print("[AVISO] Sem dados setoriais válidos em C700:D704. Tentando total em C278...")
        producao_marco_str = ler_producao_sheets(creds, SPREADSHEET_ID, "Estoque", "C278")
        if producao_marco_str:
            PRODUCAO_RAW["mar26"] = producao_marco_str
            print(f"[INFO] Produção Março total carregada (C278): {producao_marco_str}")
        else:
            print("[AVISO] Nenhum dado de produção encontrado para Março.")

    # Reprocessar PRODUCAO após atualizar PRODUCAO_RAW
    for aba, _ in MESES_2026:
        PRODUCAO[aba] = parsear_producao(PRODUCAO_RAW.get(aba, ""))

    for aba, label in MESES_2026:
        total = sum(PRODUCAO[aba].values())
        if total > 0:
            print(f"[INFO] Produção {label} por setor: {PRODUCAO[aba]}, total: {total}")

    # Listar abas existentes
    abas_existentes = [ws.title.lower() for ws in planilha.worksheets()]
    print(f"[INFO] Abas encontradas na planilha: {len(abas_existentes)}")

    # Ler todas as abas disponíveis
    todos_resultados = {}
    abas_disponiveis = []

    for aba, label in MESES_2026:
        if aba not in abas_existentes:
            print(f"[SKIP] Aba '{aba}' não existe — mês futuro.")
            continue

        print(f"\n[INFO] Lendo aba '{aba}'...")
        time.sleep(1)  # Delay para respeitar quotas da API
        df = conectar_planilha(planilha, aba)
        if df is None or df.empty:
            print(f"[SKIP] Aba '{aba}' vazia ou sem dados válidos.")
            continue

        print(f"[INFO] {len(df)} registros lidos da aba '{aba}'")
        mapa = mapear_colunas(df)
        resultado = analisar_aba(df, mapa, label)
        todos_resultados[aba] = resultado
        abas_disponiveis.append(aba)

    if len(abas_disponiveis) < 2:
        print(f"\n[ERRO] Necessário pelo menos 2 meses com dados. Encontrados: {len(abas_disponiveis)}")
        if len(abas_disponiveis) == 1:
            print(f"[INFO] Apenas '{abas_disponiveis[0]}' tem dados.")
        sys.exit(1)

    print(f"\n[INFO] {len(abas_disponiveis)} meses disponíveis: {', '.join(abas_disponiveis)}")

    # Identificar meses com dados de erro (erros_pareto não vazio)
    abas_com_erros = [a for a in abas_disponiveis
                      if not todos_resultados[a].get("erros_pareto", pd.DataFrame()).empty]
    if len(abas_com_erros) >= 2:
        aba_ant = abas_com_erros[-2]
        aba_atu = abas_com_erros[-1]
        print(f"[INFO] Comparação detalhada: {aba_ant} vs {aba_atu} (últimos 2 meses com dados de defeitos)")
    elif len(abas_com_erros) == 1:
        aba_ant = abas_com_erros[0]
        aba_atu = abas_com_erros[0]
        print(f"[INFO] Apenas {aba_atu} tem dados de defeitos. Sem comparação.")
    else:
        aba_ant = abas_disponiveis[-2]
        aba_atu = abas_disponiveis[-1]
        print(f"[AVISO] Nenhum mês com dados de tipo de defeito. Comparação genérica.")

    anterior = todos_resultados[aba_ant]
    atual = todos_resultados[aba_atu]

    comp = comparar_dois_meses(anterior, atual, aba_atu)
    comp["aba_ant"] = aba_ant
    comp["aba_atu"] = aba_atu

    # Planos de ação — top 2 do mês mais recente com dados de erro
    top2 = atual.get("erros_pareto", pd.DataFrame()).head(2)
    planos = gerar_planos(top2, comp.get("erros_comp", pd.DataFrame()), atual, aba_atu)

    # Ler Resumo 2025 (dados históricos)
    print("[INFO] Lendo Resumo 2025...")
    time.sleep(1)  # Delay para respeitar quotas da API
    resumo_2025 = ler_resumo_2025(planilha)

    # Construir Resumo 2026 a partir dos dados analisados
    print("[INFO] Construindo Resumo 2026...")
    resumo_2026 = construir_resumo_2026(todos_resultados, abas_disponiveis)

    # Nome do arquivo
    meses_label = f"{abas_disponiveis[0]}_{abas_disponiveis[-1]}"
    output_file = os.path.join(OUTPUT_DIR, f"relatorio_comparativo_{meses_label}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")

    # Gerar HTML
    html = gerar_html(todos_resultados, abas_disponiveis, comp, planos, resumo_2026, resumo_2025)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n[SUCESSO] Relatório salvo em: {output_file}")
    for aba in abas_disponiveis:
        r = todos_resultados[aba]
        print(f"  {r['mes']}: {r['total']} defeitos")
    print(f"[INFO] Variação {anterior['mes']}→{atual['mes']}: {comp['variacao_pct']}% ({comp['variacao_total']:+d})")
    prod_total = comp["producao_total"]
    if prod_total > 0:
        print(f"[INFO] Produção {atual['mes']}: {prod_total:,} unidades")
        print(f"[INFO] Taxa defeito geral {atual['mes']}: {comp['taxa_defeito_geral']}%")
    print(f"[INFO] {len(planos)} planos de ação gerados")


if __name__ == "__main__":
    main()
