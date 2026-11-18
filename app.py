import io
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

ARQUIVO = Path(__file__).with_name("controle_pedidos.xlsx")

STATUS = [
    "Aberta", "Em cotação", "Em processamento", "Em separação", "Entregue",
    "Entregue c/ divergencia", "Não Entregue", "Atraso", "Cancelado",
]
FECHADOS = {"Entregue", "Entregue c/ divergencia", "Cancelado"}
ENTREGUES = {"Entregue", "Entregue c/ divergencia"}

# Colunas que você digita (mesma ordem da aba Base do Excel modelo)
COLUNAS = [
    "C.c", "LA", "Data da Solicitação", "PN", "Descrição", "UND",
    "Quantidade total", "FORNECEDOR", "OBS", "STATUS", "PC",
    "PC ENVIADO EM", "PRAZO DE ENTREGA", "ENTREGA CONFIRMADA",
]
DATAS = ["Data da Solicitação", "PC ENVIADO EM", "PRAZO DE ENTREGA", "ENTREGA CONFIRMADA"]
# Colunas calculadas (não editáveis)
CALCULADAS = [
    "DIAS PARA ENTREGA", "Dias faltantes", "Lead Time Compras (Dias)",
    "Tempo Fornecedor (Dias)", "Lead Time Total (Dias)",
]

# Nomes aceitos na importação -> nome padrão
APELIDOS = {
    "c.c": "C.c", "cc": "C.c", "centro de custo": "C.c",
    "la": "LA", "projeto": "LA", "nome do projeto": "LA",
    "data da solicitacao": "Data da Solicitação", "pn": "PN", "descricao": "Descrição",
    "und": "UND", "quantidade total": "Quantidade total", "quantidade": "Quantidade total",
    "qtd": "Quantidade total", "fornecedor": "FORNECEDOR", "obs": "OBS", "status": "STATUS",
    "pc": "PC", "pc enviado em": "PC ENVIADO EM", "prazo de entrega": "PRAZO DE ENTREGA",
    "entrega confirmada": "ENTREGA CONFIRMADA", "entregue em": "ENTREGA CONFIRMADA",
}


def normalizar(txt) -> str:
    txt = unicodedata.normalize("NFD", str(txt))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.split()).lower()


def tabela_vazia() -> pd.DataFrame:
    df = pd.DataFrame(columns=COLUNAS)
    for c in DATAS:
        df[c] = pd.to_datetime(df[c])
    return df


def preparar(df: pd.DataFrame, resetar_indice: bool = True) -> pd.DataFrame:
    """Garante colunas, tipos e limpa textos."""
    for c in COLUNAS:
        if c not in df.columns:
            df[c] = None
    df = df[COLUNAS].copy()
    for c in DATAS:
        df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
    for c in COLUNAS:
        if c in DATAS or c == "Quantidade total":
            continue
        df[c] = (
            df[c].astype("string").str.replace("_x000D_", "", regex=False).str.strip()
        ).fillna("")
    df["Quantidade total"] = pd.to_numeric(df["Quantidade total"], errors="coerce")
    df["STATUS"] = df["STATUS"].map(
        lambda s: next((x for x in STATUS if normalizar(x) == normalizar(s)), s)
    )
    return df.reset_index(drop=True) if resetar_indice else df


def importar_excel(arquivo) -> pd.DataFrame:
    """Lê a aba Base do modelo (procura a linha de títulos pela coluna 'Descrição')."""
    xls = pd.ExcelFile(arquivo)
    aba = "Base" if "Base" in xls.sheet_names else xls.sheet_names[0]
    bruto = pd.read_excel(xls, sheet_name=aba, header=None)
    linha = None
    for i in range(min(len(bruto), 60)):
        if any(normalizar(v) == "descricao" for v in bruto.iloc[i].tolist()):
            linha = i
            break
    if linha is None:
        raise ValueError("Não encontrei a coluna 'Descrição'. Use a aba Base do modelo.")
    cab = [APELIDOS.get(normalizar(v)) for v in bruto.iloc[linha].tolist()]
    dados = bruto.iloc[linha + 1:].copy()
    dados.columns = [c if c else f"_x{j}" for j, c in enumerate(cab)]
    dados = dados.loc[:, ~dados.columns.duplicated()]
    dados = dados[[c for c in dados.columns if c in COLUNAS]]
    df = preparar(dados)
    vazio = (df["C.c"] == "") & (df["LA"] == "") & (df["Descrição"] == "") & (df["PN"] == "")
    return df[~vazio].reset_index(drop=True)


def calcular(df: pd.DataFrame) -> pd.DataFrame:
    """Mesmas fórmulas do Excel modelo."""
    hoje = pd.Timestamp.today().normalize()
    out = df.copy()
    sol, env = out["Data da Solicitação"], out["PC ENVIADO EM"]
    prazo, ent = out["PRAZO DE ENTREGA"], out["ENTREGA CONFIRMADA"]
    status = out["STATUS"]
    cancelado = status == "Cancelado"
    fechado = status.isin(FECHADOS) | ent.notna()

    out["DIAS PARA ENTREGA"] = (prazo - env).dt.days

    faltam = (prazo - hoje).dt.days
    out["Dias faltantes"] = faltam.where(env.notna() & ~fechado & prazo.notna())

    fim_compras = env.fillna(hoje)
    ltc = (fim_compras - sol).dt.days.fillna(0)
    out["Lead Time Compras (Dias)"] = ltc.where(~cancelado & (status != ""), 0)

    fim = ent.fillna(prazo)
    tf = (fim - env).dt.days.where(env.notna() & fim.notna(), 0)
    out["Tempo Fornecedor (Dias)"] = tf.where(~cancelado, 0)

    out["Lead Time Total (Dias)"] = out["Lead Time Compras (Dias)"] + out["Tempo Fornecedor (Dias)"]
    return out


def carregar_inicial() -> pd.DataFrame:
    if ARQUIVO.exists():
        try:
            return importar_excel(ARQUIVO)
        except Exception:
            pass
    return tabela_vazia()


def para_excel(df: pd.DataFrame) -> bytes:
    completo = calcular(df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", datetime_format="DD/MM/YYYY") as w:
        completo[COLUNAS + CALCULADAS].to_excel(w, sheet_name="Base", index=False)
        ws = w.sheets["Base"]
        for col, largura in zip("ABCDEFGHIJKLMNOPQRS", [8, 28, 14, 16, 50, 7, 10, 18, 28, 20, 12, 14, 14, 16, 10, 10, 12, 12, 12]):
            ws.column_dimensions[col].width = largura
    return buf.getvalue()


def salvar(df: pd.DataFrame) -> None:
    ARQUIVO.write_bytes(para_excel(df))


# ---------------------------------------------------------------- tela
st.set_page_config(page_title="Controle de Pedidos", layout="wide")
st.markdown(
    """
    <style>
      .stApp { background: #eef1f6; }
      h1, h2, h3 { color: #1d2736; }
      [data-testid="stMetric"] { background: #fff; border: 1px solid #d9e0eb; border-radius: 8px; padding: 10px 14px; }
      [data-testid="stSidebar"] { background: #e4eaf4; }
      .stButton > button[kind="primary"] { background: #2f5fa7; border-color: #2f5fa7; }
    </style>
    """,
    unsafe_allow_html=True,
)

if "df" not in st.session_state:
    st.session_state.df = carregar_inicial()
    st.session_state.msg = ""

st.title("Controle de Pedidos")
st.caption("Pedidos de compra, prazos de entrega e lead time por projeto")

# ---- barra lateral: importar / exportar / salvar
with st.sidebar:
    st.header("Arquivo")
    up = st.file_uploader("Importar Excel (aba Base)", type=["xlsx", "xls"])
    if up is not None and st.button("Adicionar pedidos do arquivo"):
        try:
            novos = importar_excel(up)
            st.session_state.df = pd.concat([st.session_state.df, novos], ignore_index=True)
            st.session_state.msg = f"{len(novos)} pedidos importados. Clique em Salvar."
            st.rerun()
        except Exception as e:
            st.error(str(e))
    # CORREÇÃO 1: use_container_width=True no lugar de width="stretch"
    if st.button("Salvar", type="primary", use_container_width=True):
        salvar(st.session_state.df)
        st.session_state.msg = f"Salvo em {ARQUIVO.name}"
    st.download_button(
        "Baixar Excel",
        data=para_excel(st.session_state.df),
        file_name="controle-de-pedidos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    if st.session_state.get("msg"):
        st.success(st.session_state.msg)
    st.divider()
    st.header("Revisão")
    rev = st.date_input("Última revisão", value=st.session_state.get("rev", pd.Timestamp.today().date()), format="DD/MM/YYYY")
    st.session_state.rev = rev
    st.write(f"Dias sem revisão: **{(pd.Timestamp.today().date() - rev).days}**")

completo = calcular(st.session_state.df)

# ---- indicadores
aberto = ~completo["STATUS"].isin(FECHADOS)
atraso = aberto & (completo["Dias faltantes"] < 0)
sete = aberto & completo["Dias faltantes"].between(0, 7)
entregues = completo["STATUS"].isin(ENTREGUES)
lt = completo.loc[entregues & (completo["Lead Time Total (Dias)"] > 0), "Lead Time Total (Dias)"]

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Em aberto", int(aberto.sum()))
c2.metric("Em atraso", int(atraso.sum()))
c3.metric("Vencem em até 7 dias", int(sete.sum()))
c4.metric("Entregues", int(entregues.sum()))
c5.metric("Lead time total médio (entregues)", f"{lt.mean():.0f} dias" if len(lt) else "–")

aba_pedidos, aba_resumo = st.tabs(["Pedidos", "Resumo por projeto"])

# ---- aba pedidos
with aba_pedidos:
    f1, f2, f3, f4 = st.columns([2, 2, 2, 2])
    busca = f1.text_input("Buscar", placeholder="PN, descrição, fornecedor, PC…")
    projetos = sorted(x for x in completo["LA"].unique() if x)
    f_proj = f2.multiselect("Projeto (LA)", projetos)
    f_status = f3.multiselect("Status", STATUS)
    f_kpi = f4.selectbox("Mostrar", ["Todos", "Em aberto", "Em atraso", "Vencem em até 7 dias", "Entregues"])

    mask = pd.Series(True, index=completo.index)
    if busca:
        alvo = completo[["C.c", "LA", "PN", "Descrição", "FORNECEDOR", "PC", "OBS"]].astype(str).agg(" ".join, axis=1).map(normalizar)
        mask &= alvo.str.contains(normalizar(busca), regex=False)
    if f_proj:
        mask &= completo["LA"].isin(f_proj)
    if f_status:
        mask &= completo["STATUS"].isin(f_status)
    mask &= {"Todos": True, "Em aberto": aberto, "Em atraso": atraso,
             "Vencem em até 7 dias": sete, "Entregues": entregues}[f_kpi]

    visivel = completo[mask]
    st.caption(f"{len(visivel)} de {len(completo)} pedidos. Clique numa célula para editar; use a última linha da tabela para adicionar um pedido. "
               "Depois de editar, clique em Salvar na barra lateral.")

    cfg = {
        "STATUS": st.column_config.SelectboxColumn("STATUS", options=STATUS, required=False),
        "Quantidade total": st.column_config.NumberColumn("Quantidade total", format="%g"),
        "Descrição": st.column_config.TextColumn("Descrição", width="large"),
        "LA": st.column_config.TextColumn("Projeto (LA)", width="medium"),
    }
    for c in DATAS:
        cfg[c] = st.column_config.DateColumn(c, format="DD/MM/YYYY")
    for c in CALCULADAS:
        cfg[c] = st.column_config.NumberColumn(c, format="%d", disabled=True,
                                               help="Calculado automaticamente")

    editado = st.data_editor(
        visivel[COLUNAS + CALCULADAS],
        column_config=cfg,
        num_rows="dynamic",
        use_container_width=True,
        height=520,
        key="editor",
    )

    # CORREÇÃO 2: resetar_indice=False mantém a posição original das linhas editadas
    base = st.session_state.df
    edicoes_limpas = preparar(editado[COLUNAS], resetar_indice=False)
    novo = pd.concat([base.drop(index=visivel.index), edicoes_limpas], ignore_index=False)
    novo = novo.sort_index(kind="stable").reset_index(drop=True)
    if not novo.equals(base):
        st.session_state.df = novo
        st.rerun()

    # botão para confirmar entrega de hoje
    abertos = completo[aberto]
    with st.expander("Confirmar entrega"):
        opcoes = {f"{i} · {r['LA']} · {r['PN'] or r['Descrição'][:40]}": i for i, r in abertos.iterrows()}
        escolha = st.selectbox("Pedido", [""] + list(opcoes.keys()))
        data = st.date_input("Data da entrega", value=pd.Timestamp.today().date(), format="DD/MM/YYYY", key="data_ent")
        if st.button("Confirmar entrega") and escolha:
            i = opcoes[escolha]
            st.session_state.df.loc[i, "ENTREGA CONFIRMADA"] = pd.Timestamp(data)
            st.session_state.df.loc[i, "STATUS"] = "Entregue"
            st.rerun()

# ---- aba resumo
with aba_resumo:
    if completo.empty:
        st.info("Os projetos aparecem aqui conforme você cadastra pedidos.")
    else:
        t = completo.copy()
        t["Em processamento"] = t["STATUS"].isin(["Em processamento", "Em separação"])
        t["Entregues"] = t["STATUS"].isin(ENTREGUES)
        t["Cancelados"] = t["STATUS"] == "Cancelado"
        t["Em atraso"] = (~t["STATUS"].isin(FECHADOS)) & (t["Dias faltantes"] < 0)
        t["Sem processo"] = ~(t["Em processamento"] | t["Entregues"] | t["Cancelados"])
        r = t.groupby(["C.c", "LA"], dropna=False).agg(
            **{"Total de itens": ("STATUS", "size"), "Sem processo": ("Sem processo", "sum"),
               "Em processamento": ("Em processamento", "sum"), "Em atraso": ("Em atraso", "sum"),
               "Entregues": ("Entregues", "sum"), "Cancelados": ("Cancelados", "sum")}
        ).reset_index()
        base_ent = (r["Total de itens"] - r["Cancelados"]).replace(0, pd.NA)
        r["Andamento"] = (r["Entregues"] / base_ent).fillna(0)
        st.dataframe(
            r.drop(columns="Cancelados"),
            use_container_width=True, hide_index=True,
            column_config={"Andamento": st.column_config.ProgressColumn("Andamento", format="percent", min_value=0, max_value=1)},
        )
