import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from io import BytesIO

st.set_page_config(page_title="Vegas Pay - Dashboard", layout="wide")
st.title("📊 Vegas Pay — Dashboard de Vendas e Rentabilidade")

# -------------------------- Ajuda --------------------------
with st.expander("ℹ️ Como usar", expanded=True):
    st.write(
        "- Faça upload de **duas planilhas**: *Vendas* e *Tabela de Custos*.\n"
        "- A planilha de **Vendas** deve ter colunas: `MCC`, `Bandeira`, `Produto`, `Valor`, `Data Transacao`, `MDR (R$)`, `Tarifa Antecipacao (R$)` e (opcional) `Vendedor`, `CNPJ Estabelecimento`, `Estabelecimento`.\n"
        "- A planilha de **Custos** deve ter colunas: `MCC`, `BANDEIRA`, `PRODUTO`, `Taxas` (ou `Taxa`), `Taxa Antecipação`, `Imposto` e (opcional) `Categoria_MCC`.\n"
        "- O app calcula o **MDR Líquido da Vegas** (imposto + custo Entrepay considerados) e gera **Resumo Mensal** + filtros.\n"
        "- (Opcional) Faça upload de **Comércios Novos** e de **Vendas PIX** para os painéis extras."
    )

# -------------------------- Helpers --------------------------
@st.cache_data
def load_excel(file, sheet_name=None):
    return pd.read_excel(file, sheet_name=sheet_name or 0)

def norm_bandeira(x: str) -> str:
    if pd.isna(x): return ""
    x = str(x).strip().upper()
    x = x.replace("MASTERCARD", "MASTER").replace("MAESTRO", "MASTER").replace("VISA ELECTRON", "VISA")
    return x

def norm_produto(x: str) -> str:
    if pd.isna(x): return ""
    x = str(x).strip().upper()
    if x.startswith("D"): return "DEBITO"
    return "CREDITO"

def fmt_brl(v: float) -> str:
    try:
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

# -------------------------- Core: merge/cálculos --------------------------
def prepare(sales_df, costs_df):
    costs_df = costs_df.rename(columns={"Taxas": "Taxa"})
    for df in (sales_df, costs_df):
        if "MCC" in df.columns:
            df["MCC"] = df["MCC"].astype(str).str.strip()
    sales_df["_BANDEIRA_KEY"] = sales_df["Bandeira"].apply(norm_bandeira)
    sales_df["_PRODUTO_KEY"]  = sales_df["Produto"].apply(norm_produto)
    costs_df["_BANDEIRA_KEY"] = costs_df["BANDEIRA"].astype(str).str.upper().str.strip()
    costs_df["_PRODUTO_KEY"]  = costs_df["PRODUTO"].astype(str).str.upper().str.strip()

    cols_needed = ["MCC","_BANDEIRA_KEY","_PRODUTO_KEY","Taxa","Taxa Antecipação","Imposto"]
    extra_cols = [c for c in ["Categoria_MCC"] if c in costs_df.columns]
    costs = costs_df[cols_needed + extra_cols].drop_duplicates()

    merged = sales_df.merge(costs, how="left", on=["MCC","_BANDEIRA_KEY","_PRODUTO_KEY"])

    merged["Taxa"] = merged["Taxa"].fillna(0.0)
    merged["Taxa Antecipação"] = merged["Taxa Antecipação"].fillna(0.0)
    merged["Imposto"] = merged["Imposto"].fillna(0.115)

    merged["Custo_Entrepay_MDR (R$)"] = merged["Valor"] * merged["Taxa"]

    total_ant = 0.0189
    intrepay_ant = 0.0147
    prop_intrepay = intrepay_ant / total_ant  # ≈ 0.77778
    prop_vegas    = 1 - prop_intrepay         # ≈ 0.22222

    ant_candidates = ["Tarifa Antecipacao (R$) ", "Tarifa Antecipacao (R$)"]
    ant_col = next((c for c in ant_candidates if c in merged.columns), None)
    merged["Tarifa_Ant"] = merged[ant_col].fillna(0) if ant_col else 0

    merged["Custo_Entrepay_Ant (R$)"] = merged["Tarifa_Ant"] * prop_intrepay
    merged["Receita_Vegas_Ant (R$)"]  = merged["Tarifa_Ant"] * prop_vegas

    merged["MDR (R$)"] = merged["MDR (R$)"].fillna(0)
    merged["Receita_Bruta_Vegas (R$)"] = merged["MDR (R$)"] + merged["Receita_Vegas_Ant (R$)"]
    merged["Imposto_sobre_Receita (R$)"] = merged["Receita_Bruta_Vegas (R$)"] * merged["Imposto"]

    merged["MDR Líquido Vegas (R$)"] = (
        merged["Receita_Bruta_Vegas (R$)"] -
        merged["Imposto_sobre_Receita (R$)"] -
        merged["Custo_Entrepay_MDR (R$)"] -
        merged["Custo_Entrepay_Ant (R$)"]
    )

    merged["Mes"] = pd.to_datetime(merged["Data Transacao"]).dt.to_period("M").astype(str)

    # Normalizações opcionais
    if "Vendedor" in merged.columns:
        merged["Vendedor"] = merged["Vendedor"].astype(str).str.strip()
    if "CNPJ Estabelecimento" in merged.columns:
        merged["CNPJ Estabelecimento"] = (
            merged["CNPJ Estabelecimento"].astype(str).str.replace(r"\D","", regex=True).str.zfill(14)
        )
    return merged

# -------------------------- Uploads base --------------------------
left, right = st.columns(2)
with left:
    vendas_file = st.file_uploader("📥 Upload — Planilha de Vendas (.xlsx)", type=["xlsx"], key="vendas")
with right:
    custos_file = st.file_uploader("📥 Upload — Tabela de Custos (.xlsx)", type=["xlsx"], key="custos")

# Extras
extra_left, extra_right = st.columns(2)
with extra_left:
    novos_file = st.file_uploader("🆕 Upload — Comércios Novos (.xlsx) [opcional]", type=["xlsx"], key="novos")
with extra_right:
    pix_file = st.file_uploader("💸 Upload — Vendas PIX (.xlsx) [opcional]", type=["xlsx"], key="pix")

if vendas_file and custos_file:
    vendas = load_excel(vendas_file)
    custos = load_excel(custos_file)
    merged = prepare(vendas.copy(), custos.copy())

    # -------------------------- Filtros --------------------------
    st.sidebar.header("Filtros")

    # MCC label (código + categoria se existir)
    if "Categoria_MCC" in custos.columns:
        mcc_map = (
            custos[["MCC","Categoria_MCC"]]
            .drop_duplicates()
            .assign(label=lambda d: d["MCC"].astype(str).str.zfill(4) + " - " + d["Categoria_MCC"].fillna(""))
        )
        mcc_to_label = dict(zip(mcc_map["MCC"], mcc_map["label"]))
        opcoes_mcc = sorted(merged["MCC"].dropna().unique().tolist())
        labels_mcc = [mcc_to_label.get(m, str(m).zfill(4)) for m in opcoes_mcc]
        sel_labels = st.sidebar.multiselect("MCC", options=labels_mcc)
        sel_mcc = {k for k, v in mcc_to_label.items() if v in sel_labels} if sel_labels else []
    else:
        sel_mcc = st.sidebar.multiselect("MCC", sorted(merged["MCC"].unique()))

    band_sel = st.sidebar.multiselect("Bandeira", sorted(merged["_BANDEIRA_KEY"].unique()))
    prod_sel = st.sidebar.multiselect("Produto", sorted(merged["_PRODUTO_KEY"].unique()))
    mes_sel  = st.sidebar.multiselect("Mês", sorted(merged["Mes"].unique()))

    vend_sel = []
    if "Vendedor" in merged.columns:
        vend_sel = st.sidebar.multiselect("Vendedor", sorted([v for v in merged["Vendedor"].dropna().unique() if v]))

    # Aplicação dos filtros
    filt = merged.copy()
    if sel_mcc:  filt = filt[filt["MCC"].isin(sel_mcc)]
    if band_sel: filt = filt[filt["_BANDEIRA_KEY"].isin(band_sel)]
    if prod_sel: filt = filt[filt["_PRODUTO_KEY"].isin(prod_sel)]
    if mes_sel:  filt = filt[filt["Mes"].isin(mes_sel)]
    if vend_sel and "Vendedor" in filt.columns:
        filt = filt[filt["Vendedor"].isin(vend_sel)]

    # -------------------------- KPIs (R$ + %) --------------------------
    vendas_brutas = float(filt["Valor"].sum())
    mdr_cobrado_r = float(filt["MDR (R$)"].sum())
    ant_r = float(filt["Tarifa_Ant"].sum())
    mdr_liq_r = float(filt["MDR Líquido Vegas (R$)"].sum())

    mdr_pct_cobrado = (mdr_cobrado_r / vendas_brutas * 100) if vendas_brutas else 0.0
    mdr_pct_liquido = (mdr_liq_r / vendas_brutas * 100) if vendas_brutas else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Vendas Brutas (R$)", f"{vendas_brutas:,.2f}")
    with c2:
        st.metric("MDR (R$) cobrado", f"{mdr_cobrado_r:,.2f}")
        st.metric("MDR (%) cobrado", f"{mdr_pct_cobrado:.2f}%", help="Fórmula: MDR(R$) / Vendas Brutas (R$) * 100")
    c3.metric("Antecipação (R$)", f"{ant_r:,.2f}")
    with c4:
        st.metric("MDR Líquido Vegas (R$)", f"{mdr_liq_r:,.2f}")
        st.metric("MDR Líquido (%)", f"{mdr_pct_liquido:.2f}%", help="Fórmula: MDR Líquido (R$) / Vendas Brutas (R$) * 100")

    # -------------------------- Resumo Mensal --------------------------
    resumo = filt.groupby("Mes", as_index=False).agg(**{
        "Vendas Brutas (R$)": ("Valor","sum"),
        "MDR (R$)": ("MDR (R$)","sum"),
        "MDR Líquido Vegas (R$)": ("MDR Líquido Vegas (R$)","sum")
    })
    resumo["MDR Líquido (%)"] = resumo.apply(
        lambda r: (r["MDR Líquido Vegas (R$)"]/r["Vendas Brutas (R$)"]*100) if r["Vendas Brutas (R$)"] else 0.0, axis=1
    )

    st.subheader("Resumo Mensal")
    st.dataframe(
        resumo.style.format({
            "Vendas Brutas (R$)": fmt_brl,
            "MDR (R$)": fmt_brl,
            "MDR Líquido Vegas (R$)": fmt_brl,
            "MDR Líquido (%)": "{:.2f}%".format
        })
    )

    # -------------------------- Gráfico (cores distintas) --------------------------
    st.subheader("Evolução Mensal")
    base = resumo.melt(id_vars="Mes",
                       value_vars=["Vendas Brutas (R$)", "MDR Líquido Vegas (R$)"],
                       var_name="Métrica", value_name="Valor")
    palette = {"Vendas Brutas (R$)": "#1f77b4", "MDR Líquido Vegas (R$)": "#2ca02c"}  # azul × verde
    chart = alt.Chart(base).mark_line(point=True).encode(
        x=alt.X("Mes:N", sort=None, title="Mês"),
        y=alt.Y("Valor:Q", title="R$"),
        color=alt.Color("Métrica:N", scale=alt.Scale(domain=list(palette.keys()),
                                                    range=list(palette.values())))
    ).properties(height=320)
    st.altair_chart(chart, use_container_width=True)

    # -------------------------- Ranking de Vendedores (se houver) --------------------------
    if "Vendedor" in filt.columns and len(filt):
        st.subheader("🏅 Ranking de Vendedores (Vendas Brutas)")
        rank = filt.groupby("Vendedor", as_index=False)["Valor"].sum().sort_values("Valor", ascending=False)
        st.dataframe(rank.head(15).style.format({"Valor": fmt_brl}))

    # -------------------------- Detalhe das Transações --------------------------
    st.subheader("Detalhe das Transações")
    cols_to_show = ["Data Transacao","Estabelecimento","CNPJ Estabelecimento","MCC",
                    "_BANDEIRA_KEY","_PRODUTO_KEY","Vendedor",
                    "Valor","MDR (R$)","Tarifa_Ant",
                    "Custo_Entrepay_MDR (R$)","Custo_Entrepay_Ant (R$)",
                    "Receita_Bruta_Vegas (R$)","Imposto_sobre_Receita (R$)","MDR Líquido Vegas (R$)"]
    cols_exist = [c for c in cols_to_show if c in filt.columns]
    st.dataframe(filt[cols_exist])

    # -------------------------- Comércios Novos (opcional) --------------------------
    painel_novos = None
    if novos_file is not None:
        novos = pd.read_excel(novos_file)
        # Colunas mínimas: CNPJ, Nome_Fantasia, MCC, Vendedor, Data_Fechamento, Previsao_Mensal_R$, Mes_Referencia (opcional)
        colmap = {c.lower(): c for c in novos.columns}
        # Normalizações
        if "CNPJ" in novos.columns:
            novos["CNPJ"] = novos["CNPJ"].astype(str).str.replace(r"\D","", regex=True).str.zfill(14)
        if "Mes_Referencia" in novos.columns:
            novos["MesRef"] = pd.to_datetime(novos["Mes_Referencia"]).dt.to_period("M").dt.to_timestamp()
        elif "Data_Fechamento" in novos.columns:
            novos["MesRef"] = pd.to_datetime(novos["Data_Fechamento"]).dt.to_period("M").dt.to_timestamp()
        else:
            novos["MesRef"] = pd.Timestamp.now().to_period("M").to_timestamp()

        # Vendas realizadas por CNPJ/mês (a partir do universo filtrado por mês/vendedor, se desejar)
        vendas_cnpj_mes = (
            merged.assign(
                CNPJ=lambda d: d.get("CNPJ Estabelecimento","").astype(str).str.replace(r"\D","", regex=True).str.zfill(14),
                MesRef=lambda d: pd.to_datetime(d["Data Transacao"]).dt.to_period("M").dt.to_timestamp()
            )
            .groupby(["CNPJ","MesRef"], as_index=False)["Valor"].sum()
            .rename(columns={"Valor":"Realizado_R$"})
        )

        painel_novos = novos.merge(vendas_cnpj_mes, on=["CNPJ","MesRef"], how="left").fillna({"Realizado_R$":0})
        if "Previsao_Mensal_R$" in painel_novos.columns:
            painel_novos["Atingimento_%"] = painel_novos.apply(
                lambda r: (r["Realizado_R$"]/r["Previsao_Mensal_R$"]*100) if r["Previsao_Mensal_R$"] else 0.0, axis=1
            )
        else:
            painel_novos["Previsao_Mensal_R$"] = 0.0
            painel_novos["Atingimento_%"] = 0.0
        painel_novos["Alerta"] = np.where(painel_novos["Atingimento_%"] < 70, "⚠️ <70%", "")

        st.subheader("🆕 Comércios Novos — Acompanhamento")
        cols_show_novos = [c for c in ["Vendedor","CNPJ","Nome_Fantasia","MesRef",
                                       "Previsao_Mensal_R$","Realizado_R$","Atingimento_%","Alerta"] if c in painel_novos.columns]
        st.dataframe(
            painel_novos[cols_show_novos]
            .sort_values(["Atingimento_%","Realizado_R$"], ascending=[True, False])
            .style.format({
                "Previsao_Mensal_R$": fmt_brl,
                "Realizado_R$": fmt_brl,
                "Atingimento_%": "{:.1f}%".format
            })
        )

    # -------------------------- PIX (opcional) --------------------------
    pix_resumo = None
    if pix_file is not None:
        df_pix = pd.read_excel(pix_file)
        # Esperado: CNPJ, Nome Fantasia, MCC, Tipo MCC, Valor (R$), Data da Transação
        if "CNPJ" in df_pix.columns:
            df_pix["CNPJ"] = df_pix["CNPJ"].astype(str).str.replace(r"\D","", regex=True).str.zfill(14)
        if "Data da Transação" in df_pix.columns:
            df_pix["Data"] = pd.to_datetime(df_pix["Data da Transação"])
        elif "Data" in df_pix.columns:
            df_pix["Data"] = pd.to_datetime(df_pix["Data"])
        else:
            df_pix["Data"] = pd.Timestamp.now()
        df_pix["Mes"] = df_pix["Data"].dt.to_period("M").dt.to_timestamp()

        # Regras PIX
        imposto_pix = st.sidebar.number_input("Imposto s/ Receita (PIX) %", min_value=0.0, max_value=30.0, value=11.5, step=0.1)
        val_col = "Valor (R$)" if "Valor (R$)" in df_pix.columns else "Valor"
        df_pix[val_col] = pd.to_numeric(df_pix[val_col], errors="coerce").fillna(0.0)

        df_pix["Receita_Vegas"] = df_pix[val_col] * 0.0024
        df_pix["Custo_Fixo"] = 0.25
        df_pix["Imposto_R$"] = df_pix["Receita_Vegas"] * (imposto_pix/100.0)
        df_pix["MDR_Liq_Vegas_PIX"] = df_pix["Receita_Vegas"] - df_pix["Imposto_R$"] - df_pix["Custo_Fixo"]

        # KPIs PIX
        k1, k2, k3 = st.columns(3)
        k1.metric("PIX — Valor Bruto (R$)", f"{df_pix[val_col].sum():,.2f}")
        k2.metric("PIX — Receita Vegas (R$)", f"{df_pix['Receita_Vegas'].sum():,.2f}")
        k3.metric("PIX — MDR Líquido (R$)", f"{df_pix['MDR_Liq_Vegas_PIX'].sum():,.2f}")

        # Resumo mensal PIX
        pix_resumo = df_pix.groupby("Mes", as_index=False).agg({
            val_col:"sum",
            "Receita_Vegas":"sum",
            "MDR_Liq_Vegas_PIX":"sum"
        }).rename(columns={val_col:"Valor Bruto (R$)"})

        st.subheader("Resumo Mensal — PIX")
        st.dataframe(
            pix_resumo.style.format({
                "Valor Bruto (R$)": fmt_brl,
                "Receita_Vegas": fmt_brl,
                "MDR_Liq_Vegas_PIX": fmt_brl
            })
        )

    # -------------------------- Exportação Excel --------------------------
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filt.to_excel(writer, sheet_name="Lancamentos_Filtrados", index=False)
        resumo.to_excel(writer, sheet_name="Resumo_Mensal", index=False)
        if painel_novos is not None:
            painel_novos.to_excel(writer, sheet_name="Comercios_Novos", index=False)
        if pix_resumo is not None:
            pix_resumo.to_excel(writer, sheet_name="PIX_Resumo", index=False)

    st.download_button("📤 Baixar Excel (filtro atual + extras)",
                       data=buffer.getvalue(),
                       file_name="vegas_pay_resumo.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("Faça upload das duas planilhas para ver o painel.")
