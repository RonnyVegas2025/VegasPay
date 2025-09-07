
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from io import BytesIO

# =========================
# Config & helpers
# =========================
st.set_page_config(page_title="Vegas Pay — Dashboard", layout="wide")

def fmt_brl(v):
    try:
        return "R$ {:,.2f}".format(float(v)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return v

def fmt_pct(v):
    try:
        return "{:,.2f}%".format(float(v)).replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return v

def norm_bandeira(x: str) -> str:
    if pd.isna(x): return ""
    x = str(x).strip().upper()
    x = x.replace("MASTERCARD", "MASTER").replace("MAESTRO", "MASTER").replace("VISA ELECTRON", "VISA")
    return x

def norm_prod(x: str) -> str:
    if pd.isna(x): return ""
    x = str(x).strip().upper()
    return "DEBITO" if x.startswith("D") else "CREDITO"

def ensure_period_str(s):
    return pd.to_datetime(s, errors="coerce").dt.to_period("M").astype(str)

# =========================
# Sessão e navegação
# =========================
if "data" not in st.session_state:
    st.session_state.data = {"vendas": None, "custos": None, "pix": None, "comercios": None}

page = st.sidebar.radio("Navegação", ["📤 Upload", "📊 Dashboard"])

# =========================
# Página: Upload
# =========================
if page == "📤 Upload":
    st.title("📤 Upload de Planilhas")

    with st.expander("Como usar", expanded=True):
        st.markdown("""
        **Arquivos aceitos**
        1) **Vendas (Cartões)** — colunas: `MCC`, `Bandeira`, `Produto`, `Valor`, `Data Transacao`,
        `MDR (R$)`, `Tarifa Antecipacao (R$)`, *(opcionais)* `Vendedor`, `CNPJ Estabelecimento`, `Estabelecimento`.
        2) **Tabela de Custos** — colunas: `MCC`, `BANDEIRA`, `PRODUTO`, `Taxas` (ou `Taxa`), `Taxa Antecipação`,
        `Imposto`, *(opcional)* `Categoria_MCC`.
        3) **Vendas PIX** — colunas: `Nome_Fantasia`, `CNPJ`, `MCC`, `Tipo_MCC`, `Valor`, `Data da Transação`
           (custo fixo **R$0,25** e MDR Vegas **0,24%** aplicados automaticamente).
        4) **Comércios Novos** *(opcional)* — colunas: `CNPJ`, `Nome_Fantasia`, `Vendedor`, `MCC`, `Previsao_Mensal_R$`,
        `Data_Fechamento` (ou `Mes_Referencia`).
        """)

    c1, c2 = st.columns(2)
    with c1:
        vendas_file = st.file_uploader("Vendas (Cartões) — .xlsx", type=["xlsx"], key="up_vendas")
        custos_file = st.file_uploader("Tabela de Custos — .xlsx", type=["xlsx"], key="up_custos")
    with c2:
        pix_file = st.file_uploader("Vendas PIX — .xlsx (opcional)", type=["xlsx"], key="up_pix")
        com_file = st.file_uploader("Comércios Novos — .xlsx (opcional)", type=["xlsx"], key="up_com")

    if st.button("📥 Carregar dados"):
        def rd(f):
            try:
                return pd.read_excel(f) if f is not None else None
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")
                return None

        st.session_state.data["vendas"]    = rd(vendas_file)
        st.session_state.data["custos"]    = rd(custos_file)
        st.session_state.data["pix"]       = rd(pix_file)
        st.session_state.data["comercios"] = rd(com_file)
        st.success("Planilhas carregadas para a sessão! Vá para **📊 Dashboard**.")

# =========================
# Página: Dashboard
# =========================
if page == "📊 Dashboard":
    st.title("📊 Vegas Pay — Dashboard de Vendas e Rentabilidade")

    vendas_raw = st.session_state.data.get("vendas")
    custos_raw = st.session_state.data.get("custos")

    if vendas_raw is None or custos_raw is None:
        st.info("Faça o upload de **Vendas (Cartões)** e **Tabela de Custos** na página **📤 Upload**.")
        st.stop()

    pix_raw       = st.session_state.data.get("pix")
    comercios_raw = st.session_state.data.get("comercios")

    # =========================
    # Preparação — Cartões
    # =========================
    vendas = vendas_raw.copy()
    custos = custos_raw.copy().rename(columns={"Taxas": "Taxa"})
    for df in (vendas, custos):
        if "MCC" in df.columns:
            df["MCC"] = df["MCC"].astype(str).str.strip()

    vendas["_BANDEIRA_KEY"] = vendas["Bandeira"].apply(norm_bandeira)
    vendas["_PRODUTO_KEY"]  = vendas["Produto"].apply(norm_prod)
    custos["_BANDEIRA_KEY"] = custos["BANDEIRA"].astype(str).str.upper().str.strip()
    custos["_PRODUTO_KEY"]  = custos["PRODUTO"].astype(str).str.upper().str.strip()

    # Categoria MCC (para rótulo do filtro)
    cat_map = {}
    if "Categoria_MCC" in custos.columns:
        cat_map = custos[["MCC", "Categoria_MCC"]].drop_duplicates().set_index("MCC")["Categoria_MCC"].to_dict()

    # Merge de custos
    cols_needed = ["MCC","_BANDEIRA_KEY","_PRODUTO_KEY","Taxa","Taxa Antecipação","Imposto"]
    custos_slim = custos[cols_needed].drop_duplicates()
    merged = vendas.merge(custos_slim, how="left", on=["MCC","_BANDEIRA_KEY","_PRODUTO_KEY"])

    # Defaults
    merged["Taxa"] = merged["Taxa"].fillna(0.0)
    merged["Taxa Antecipação"] = merged["Taxa Antecipação"].fillna(0.0)
    merged["Imposto"] = merged["Imposto"].fillna(0.115)

    # Cálculos Cartões — **Modelo que bate com a planilha**
    # Parte da antecipação da Vegas:
    total_ant = 0.0189
    intrepay_ant = 0.0147
    prop_intrepay = intrepay_ant / total_ant  # ~0.7778
    prop_vegas    = 1 - prop_intrepay         # ~0.2222

    # Coluna de antecipação na base de Vendas pode ter dois nomes:
    ant_candidates = ["Tarifa Antecipacao (R$) ", "Tarifa Antecipacao (R$)"]
    ant_col = next((c for c in ant_candidates if c in merged.columns), None)
    merged["Tarifa_Ant"] = merged[ant_col].fillna(0.0) if ant_col else 0.0

    # Custos/Receitas
    merged["Valor"] = pd.to_numeric(merged["Valor"], errors="coerce").fillna(0.0)
    merged["MDR (R$)"] = pd.to_numeric(merged["MDR (R$)"], errors="coerce").fillna(0.0)
    merged["Custo_Entrepay_MDR (R$)"] = merged["Valor"] * merged["Taxa"]
    merged["Receita_Vegas_Ant (R$)"] = merged["Tarifa_Ant"] * prop_vegas
    merged["Vegas_Bruto (R$)"] = merged["MDR (R$)"] - merged["Custo_Entrepay_MDR (R$)"] + merged["Receita_Vegas_Ant (R$)"]
    merged["Imposto_sobre_VegasBruto (R$)"] = merged["Vegas_Bruto (R$)"] * merged["Imposto"]
    merged["MDR Líquido Vegas (R$)"] = merged["Vegas_Bruto (R$)"] - merged["Imposto_sobre_VegasBruto (R$)"]

    merged["Mes"] = ensure_period_str(merged["Data Transacao"])

    # =========================
    # Preparação — PIX (opcional)
    # =========================
    pix = None
    resumo_pix = pd.DataFrame()
    if pix_raw is not None and len(pix_raw) > 0:
        pix = pix_raw.copy()

        # normaliza nomes
        # data
        if not any(c == "Data Transacao" for c in pix.columns):
            for c in pix.columns:
                if c.lower().startswith("data"):
                    pix.rename(columns={c: "Data Transacao"}, inplace=True)
                    break
        # valor
        if "Valor" not in pix.columns:
            cand = [c for c in pix.columns if "valor" in c.lower()]
            if cand: pix.rename(columns={cand[0]: "Valor"}, inplace=True)

        pix["Valor"] = pd.to_numeric(pix["Valor"], errors="coerce").fillna(0.0)
        pix["Mes"]   = ensure_period_str(pix["Data Transacao"])
        pix["Qtde"]  = 1

        # Regras PIX
        mdr_pix   = 0.0024  # 0,24%
        custo_fxo = 0.25    # R$ por transação
        aliq_imp  = merged["Imposto"].median() if not merged["Imposto"].empty else 0.115

        pix["Receita_Vegas (R$)"]    = pix["Valor"] * mdr_pix
        pix["Imposto (R$)"]          = pix["Receita_Vegas (R$)"] * aliq_imp
        pix["Custo_Fixo (R$)"]       = pix["Qtde"] * custo_fxo
        pix["MDR Líquido Vegas (R$)"]= pix["Receita_Vegas (R$)"] - pix["Imposto (R$)"] - pix["Custo_Fixo (R$)"]

    # =========================
    # Filtros (MCC com categoria)
    # =========================
    st.sidebar.header("Filtros")

    merged["_MCC_LABEL"] = merged["MCC"].astype(str) + " - " + merged["MCC"].astype(str).map(cat_map).fillna("Sem categoria")
    mcc_options = sorted(merged["_MCC_LABEL"].unique())

    mcc_sel  = st.sidebar.multiselect("MCC (c/ categoria)", mcc_options)
    band_sel = st.sidebar.multiselect("Bandeira", sorted(merged["_BANDEIRA_KEY"].unique()))
    prod_sel = st.sidebar.multiselect("Produto",  sorted(merged["_PRODUTO_KEY"].unique()))
    vend_sel = st.sidebar.multiselect("Vendedor", sorted(v for v in merged.get("Vendedor", pd.Series([])).astype(str).unique() if v and v.lower()!="nan"))
    mes_sel  = st.sidebar.multiselect("Mês",      sorted(merged["Mes"].unique()))
    meio_sel = st.sidebar.multiselect("Meio",     ["Cartão","PIX"])  # para visão geral

    # aplica filtros em cartões
    filt_cart = merged.copy()
    if mcc_sel:
        filt_cart = filt_cart[filt_cart["_MCC_LABEL"].isin(mcc_sel)]
    if band_sel:
        filt_cart = filt_cart[filt_cart["_BANDEIRA_KEY"].isin(band_sel)]
    if prod_sel:
        filt_cart = filt_cart[filt_cart["_PRODUTO_KEY"].isin(prod_sel)]
    if vend_sel and "Vendedor" in filt_cart.columns:
        filt_cart = filt_cart[filt_cart["Vendedor"].astype(str).isin(vend_sel)]
    if mes_sel:
        filt_cart = filt_cart[filt_cart["Mes"].isin(mes_sel)]

    # aplica filtros em PIX (se houver)
    filt_pix = None
    if pix is not None and len(pix) > 0:
        filt_pix = pix.copy()
        if mes_sel:
            filt_pix = filt_pix[filt_pix["Mes"].isin(mes_sel)]
        if "MCC" in filt_pix.columns and mcc_sel:
            lbl = filt_pix["MCC"].astype(str) + " - " + filt_pix["MCC"].astype(str).map(cat_map).fillna("Sem categoria")
            filt_pix = filt_pix[lbl.isin(mcc_sel)]

    # =========================
    # Visão Geral (Cartões + PIX)
    # =========================
    st.subheader("Visão Geral — Movimentação + MDR (Cartões + PIX)")
    use_cart = ("Cartão" in meio_sel) if meio_sel else True
    use_pix  = (filt_pix is not None) and (("PIX" in meio_sel) if meio_sel else True)

    total_vendas = filt_cart["Valor"].sum() if use_cart else 0.0
    total_mdr_liq_cart = filt_cart["MDR Líquido Vegas (R$)"].sum() if use_cart else 0.0

    total_mdr_liq_pix = 0.0
    if use_pix:
        total_vendas     += filt_pix["Valor"].sum()
        total_mdr_liq_pix = filt_pix["MDR Líquido Vegas (R$)"].sum()

    total_mdr_liq = total_mdr_liq_cart + total_mdr_liq_pix
    mdr_liq_pct   = (total_mdr_liq / total_vendas * 100) if total_vendas > 0 else 0.0

    k1,k2,k3 = st.columns(3)
    k1.metric("Vendas Brutas (R$) — Geral", fmt_brl(total_vendas))
    k2.metric("MDR Líquido (R$) — Geral",   fmt_brl(total_mdr_liq))
    k3.metric("MDR Líquido (%) — Geral",    fmt_pct(mdr_liq_pct))

    st.divider()

    # =========================
    # Seção Cartões — KPIs e Resumo
    # =========================
    st.subheader("Cartões — KPIs e Resumo")

    # KPIs detalhando “Brutos”
    vendas_brutas      = filt_cart["Valor"].sum()
    mdr_bruto_cart     = filt_cart["MDR (R$)"].sum()
    mdr_bruto_cart_pct = (mdr_bruto_cart / vendas_brutas * 100) if vendas_brutas>0 else 0.0

    ant_r              = filt_cart["Tarifa_Ant"].sum()
    # parte da Vegas na antecipação (bruta)
    ant_vegas_r        = filt_cart["Receita_Vegas_Ant (R$)"].sum()
    ant_vegas_pct      = (ant_vegas_r / vendas_brutas * 100) if vendas_brutas>0 else 0.0

    mdr_bruto_pix      = filt_pix["Receita_Vegas (R$)"].sum() if filt_pix is not None else 0.0
    mdr_bruto_pix_pct  = (mdr_bruto_pix / vendas_brutas * 100) if vendas_brutas>0 else 0.0

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Vendas Brutas (R$)", fmt_brl(vendas_brutas))
    c2.metric("MDR (R$) Cartões **Bruto**", f"{fmt_brl(mdr_bruto_cart)} ({fmt_pct(mdr_bruto_cart_pct)})")
    c3.metric("MDR (R$) Antecipação **Bruto** (parte Vegas)", f"{fmt_brl(ant_vegas_r)} ({fmt_pct(ant_vegas_pct)})")
    c4.metric("MDR (R$) PIX **Bruto**", f"{fmt_brl(mdr_bruto_pix)} ({fmt_pct(mdr_bruto_pix_pct)})")
    c5.metric("MDR Líquido (R$)", fmt_brl(filt_cart["MDR Líquido Vegas (R$)"].sum()))

    # Resumo mensal (cartões)
    resumo_cart = pd.DataFrame()
    if not filt_cart.empty:
        resumo_cart = filt_cart.groupby("Mes", as_index=False).agg(
            **{
                "Vendas Brutas (R$)": ("Valor","sum"),
                "MDR (R$)": ("MDR (R$)","sum"),
                "MDR Líquido Vegas (R$)": ("MDR Líquido Vegas (R$)","sum"),
            }
        )
        resumo_cart["MDR Líquido (%)"] = np.where(
            resumo_cart["Vendas Brutas (R$)"]>0,
            resumo_cart["MDR Líquido Vegas (R$)"]/resumo_cart["Vendas Brutas (R$)"]*100, 0
        )

    st.dataframe(
        resumo_cart.style.format({
            "Vendas Brutas (R$)":fmt_brl, "MDR (R$)":fmt_brl,
            "MDR Líquido Vegas (R$)":fmt_brl, "MDR Líquido (%)":fmt_pct
        }),
        use_container_width=True
    )

    # Auditoria por mês (Cartões)
    aud = pd.DataFrame()
    if not filt_cart.empty:
        aud = filt_cart.groupby("Mes", as_index=False).agg(
            **{
                "Vendas_Brutas_R$": ("Valor","sum"),
                "MDR_R$": ("MDR (R$)","sum"),
                "Tarifa_Ant_R$": ("Tarifa_Ant","sum"),
                "Receita_Ant_Vegas": ("Receita_Vegas_Ant (R$)","sum"),
                "Custo_Entrepay_MDR_R$": ("Custo_Entrepay_MDR (R$)","sum"),
                "Vegas_Bruto_R$": ("Vegas_Bruto (R$)","sum"),
                "Imposto_R$": ("Imposto_sobre_VegasBruto (R$)","sum"),
                "MDR_Liquido_R$": ("MDR Líquido Vegas (R$)","sum"),
            }
        )
        aud["MDR_Líquido_%"] = np.where(aud["Vendas_Brutas_R$"]>0, aud["MDR_Liquido_R$"]/aud["Vendas_Brutas_R$"]*100,0)

    with st.expander("Auditoria de Cálculo (Cartões)", expanded=False):
        st.dataframe(
            aud.style.format({
                "Vendas_Brutas_R$":fmt_brl, "MDR_R$":fmt_brl, "Tarifa_Ant_R$":fmt_brl,
                "Receita_Ant_Vegas":fmt_brl, "Custo_Entrepay_MDR_R$":fmt_brl,
                "Vegas_Bruto_R$":fmt_brl, "Imposto_R$":fmt_brl,
                "MDR_Liquido_R$":fmt_brl, "MDR_Líquido_%":fmt_pct
            }),
            use_container_width=True
        )

    # Gráfico evolução (Cartões) — protegido
    if not resumo_cart.empty:
        try:
            plot_df = resumo_cart.melt("Mes", value_vars=["Vendas Brutas (R$)", "MDR Líquido Vegas (R$)"],
                                       var_name="Métrica", value_name="Valor")
            chart_cart = alt.Chart(plot_df).mark_line(point=True).encode(
                x="Mes:N", y=alt.Y("Valor:Q", title="R$"),
                color=alt.Color("Métrica:N", scale=alt.Scale(range=["#1f77b4","#2ca02c"])),
                tooltip=["Mes","Métrica","Valor"]
            )
            st.altair_chart(chart_cart, use_container_width=True)
        except Exception as e:
            st.warning(f"Não foi possível renderizar o gráfico (Cartões). Detalhes: {e}")

    st.divider()

    # =========================
    # Seção PIX (se houver)
    # =========================
    if filt_pix is not None and not filt_pix.empty:
        st.subheader("PIX — KPIs e Resumo")
        p1,p2,p3,p4 = st.columns(4)
        p1.metric("Valor Bruto (R$)", fmt_brl(filt_pix["Valor"].sum()))
        p2.metric("MDR (R$) Bruto (0,24%)", fmt_brl(filt_pix["Receita_Vegas (R$)"].sum()))
        p3.metric("Imposto (R$)", fmt_brl(filt_pix["Imposto (R$)"].sum()))
        p4.metric("MDR Líquido (R$)", fmt_brl(filt_pix["MDR Líquido Vegas (R$)"].sum()))

        resumo_pix = filt_pix.groupby("Mes", as_index=False).agg(
            **{
                "Valor Bruto (R$)": ("Valor","sum"),
                "MDR Bruto (R$)": ("Receita_Vegas (R$)","sum"),
                "Imposto (R$)": ("Imposto (R$)","sum"),
                "MDR Líquido (R$)": ("MDR Líquido Vegas (R$)","sum"),
            }
        )
        resumo_pix["MDR Líquido (%)"] = np.where(
            resumo_pix["Valor Bruto (R$)"]>0,
            resumo_pix["MDR Líquido (R$)"]/resumo_pix["Valor Bruto (R$)"]*100,0
        )
        st.dataframe(
            resumo_pix.style.format({
                "Valor Bruto (R$)":fmt_brl, "MDR Bruto (R$)":fmt_brl,
                "Imposto (R$)":fmt_brl, "MDR Líquido (R$)":fmt_brl, "MDR Líquido (%)":fmt_pct
            }),
            use_container_width=True
        )

        if not resumo_pix.empty:
            try:
                plot_pix = resumo_pix.melt("Mes", value_vars=["Valor Bruto (R$)", "MDR Líquido (R$)"],
                                           var_name="Métrica", value_name="Valor")
                chart_pix = alt.Chart(plot_pix).mark_line(point=True).encode(
                    x="Mes:N", y=alt.Y("Valor:Q", title="R$"),
                    color=alt.Color("Métrica:N", scale=alt.Scale(range=["#1f77b4","#2ca02c"])),
                    tooltip=["Mes","Métrica","Valor"]
                )
                st.altair_chart(chart_pix, use_container_width=True)
            except Exception as e:
                st.warning(f"Não foi possível renderizar o gráfico (PIX). Detalhes: {e}")

    st.divider()

    # =========================
    # Comércios Novos (se houver)
    # =========================
    if comercios_raw is not None and len(comercios_raw) > 0:
        st.subheader("Comércios Novos — Acompanhamento")

        com = comercios_raw.copy()
        # normaliza datas e mês
        if "Mes_Referencia" in com.columns:
            com["Mes"] = com["Mes_Referencia"].astype(str)
        else:
            com["Mes"] = ensure_period_str(com.get("Data_Fechamento"))

        if "Previsao_Mensal_R$" in com.columns:
            com["Previsao_Mensal_R$"] = pd.to_numeric(com["Previsao_Mensal_R$"], errors="coerce").fillna(0.0)
        else:
            com["Previsao_Mensal_R$"] = 0.0

        # agrega realizado (Cartões+PIX) por CNPJ e mês
        realized = []
        if "CNPJ Estabelecimento" in merged.columns:
            card_real = filt_cart.groupby(["Mes","CNPJ Estabelecimento"], as_index=False)["Valor"].sum()
            card_real.rename(columns={"CNPJ Estabelecimento":"CNPJ","Valor":"Realizado_Cartao_R$"}, inplace=True)
            realized.append(card_real)
        if filt_pix is not None and "CNPJ" in filt_pix.columns:
            pix_real = filt_pix.groupby(["Mes","CNPJ"], as_index=False)["Valor"].sum()
            pix_real.rename(columns={"Valor":"Realizado_PIX_R$"}, inplace=True)
            realized.append(pix_real)

        if realized:
            from functools import reduce
            df_real = reduce(lambda a,b: pd.merge(a,b, how="outer", on=["Mes","CNPJ"]), realized)
            for c in ["Realizado_Cartao_R$","Realizado_PIX_R$"]:
                if c not in df_real.columns: df_real[c]=0.0
            df_real["Realizado_R$"] = df_real["Realizado_Cartao_R$"].fillna(0)+df_real["Realizado_PIX_R$"].fillna(0)

            if "CNPJ" in com.columns:
                com["CNPJ"] = com["CNPJ"].astype(str).str.replace(r"\D","",regex=True)
                df_real["CNPJ"] = df_real["CNPJ"].astype(str).str.replace(r"\D","",regex=True)
                com = com.merge(df_real[["Mes","CNPJ","Realizado_R$"]], how="left", on=["Mes","CNPJ"])
                com["Realizado_R$"] = com["Realizado_R$"].fillna(0.0)
        else:
            com["Realizado_R$"] = 0.0

        # filtros de comércios (respeitam filtros globais Mes / Vendedor)
        if mes_sel:
            com = com[com["Mes"].isin(mes_sel)]
        if vend_sel and "Vendedor" in com.columns:
            com = com[com["Vendedor"].astype(str).isin(vend_sel)]

        com["Atingimento_%"] = np.where(
            com["Previsao_Mensal_R$"]>0,
            com["Realizado_R$"]/com["Previsao_Mensal_R$"]*100, 0.0
        )

        cA,cB,cC,cD = st.columns(4)
        cA.metric("Qtd Comércios", f"{len(com):,}".replace(",","."))  # separador simples
        cB.metric("Previsão Total (R$)", fmt_brl(com["Previsao_Mensal_R$"].sum()))
        cC.metric("Realizado Total (R$)", fmt_brl(com["Realizado_R$"].sum()))
        cD.metric("Atingimento Médio (%)", fmt_pct(com["Atingimento_%"].mean() if len(com)>0 else 0))

        cols_show = []
        for c in ["Mes","Nome_Fantasia","CNPJ","Vendedor","MCC"]:
            if c in com.columns: cols_show.append(c)
        cols_show += ["Previsao_Mensal_R$","Realizado_R$","Atingimento_%"]

        st.dataframe(
            com[cols_show].style.format({
                "Previsao_Mensal_R$":fmt_brl, "Realizado_R$":fmt_brl, "Atingimento_%":fmt_pct
            }),
            use_container_width=True
        )

    # =========================
    # Exportação (Excel) — resp. aos filtros
    # =========================
    st.subheader("📤 Exportar (filtro atual)")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        try:
            filt_cart.to_excel(writer, sheet_name="Cartoes_Filtrados", index=False)
            if not resumo_cart.empty:
                resumo_cart.to_excel(writer, sheet_name="Resumo_Cartoes", index=False)
            if not aud.empty:
                aud.to_excel(writer, sheet_name="Auditoria_Cartoes", index=False)
            if filt_pix is not None and not filt_pix.empty:
                filt_pix.to_excel(writer, sheet_name="PIX_Filtrado", index=False)
                if not resumo_pix.empty:
                    resumo_pix.to_excel(writer, sheet_name="Resumo_PIX", index=False)
            if comercios_raw is not None and len(comercios_raw)>0 and not com.empty:
                com[cols_show].to_excel(writer, sheet_name="Comercios_Novos", index=False)
        except Exception as e:
            st.warning(f"Exportação: parte dos dados não pôde ser gravada ({e}).")
    st.download_button("Baixar Excel", data=buffer.getvalue(),
                       file_name="vegas_pay_export.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
