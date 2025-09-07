
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Vegas Pay - Dashboard", layout="wide")

st.title("📊 Vegas Pay — Dashboard de Vendas e Rentabilidade")

with st.expander("ℹ️ Como usar", expanded=True):
    st.write(
        "- Faça upload de **duas planilhas**: *Vendas* e *Tabela de Custos*.\n"
        "- A planilha de **Vendas** deve ter colunas: `MCC`, `Bandeira`, `Produto`, `Valor`, `Data Transacao`, `MDR (R$)`, `Tarifa Antecipacao (R$)`.\n"
        "- A planilha de **Custos** deve ter colunas: `MCC`, `BANDEIRA`, `PRODUTO`, `Taxas` (ou `Taxa`), `Taxa Antecipação`, `Imposto`.\n"
        "- O app calcula o **MDR Líquido da Vegas** (já com imposto e custo Entrepay) e gera **resumo mensal** + filtros."
    )

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
    costs = costs_df[cols_needed].drop_duplicates()

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
    return merged

left, right = st.columns(2)
with left:
    vendas_file = st.file_uploader("📥 Upload — Planilha de Vendas (.xlsx)", type=["xlsx"], key="vendas")
with right:
    custos_file = st.file_uploader("📥 Upload — Tabela de Custos (.xlsx)", type=["xlsx"], key="custos")

if vendas_file and custos_file:
    vendas = load_excel(vendas_file)
    custos = load_excel(custos_file)
    merged = prepare(vendas.copy(), custos.copy())

    st.sidebar.header("Filtros")
    mcc_sel = st.sidebar.multiselect("MCC", sorted(merged["MCC"].unique()))
    band_sel = st.sidebar.multiselect("Bandeira", sorted(merged["_BANDEIRA_KEY"].unique()))
    prod_sel = st.sidebar.multiselect("Produto", sorted(merged["_PRODUTO_KEY"].unique()))
    mes_sel  = st.sidebar.multiselect("Mês", sorted(merged["Mes"].unique()))

    filt = merged.copy()
    if mcc_sel:  filt = filt[filt["MCC"].isin(mcc_sel)]
    if band_sel: filt = filt[filt["_BANDEIRA_KEY"].isin(band_sel)]
    if prod_sel: filt = filt[filt["_PRODUTO_KEY"].isin(prod_sel)]
    if mes_sel:  filt = filt[filt["Mes"].isin(mes_sel)]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Vendas Brutas (R$)", f"{filt['Valor'].sum():,.2f}")
    c2.metric("MDR (R$) cobrado", f"{filt['MDR (R$)'].sum():,.2f}")
    c3.metric("Antecipação (R$)", f"{filt['Tarifa_Ant'].sum():,.2f}")
    c4.metric("MDR Líquido Vegas (R$)", f"{filt['MDR Líquido Vegas (R$)'].sum():,.2f}")

    resumo = filt.groupby("Mes", as_index=False).agg(**{
        "Vendas Brutas (R$)": ("Valor","sum"),
        "MDR Líquido Vegas (R$)": ("MDR Líquido Vegas (R$)","sum")
    })
    st.subheader("Resumo Mensal")
    st.dataframe(resumo)

    st.subheader("Evolução Mensal")
    st.line_chart(resumo.set_index("Mes"))

    st.subheader("Detalhe das Transações")
    cols_to_show = ["Data Transacao","Estabelecimento","MCC","_BANDEIRA_KEY","_PRODUTO_KEY","Valor","MDR (R$)","Tarifa_Ant",
                    "Custo_Entrepay_MDR (R$)","Custo_Entrepay_Ant (R$)",
                    "Receita_Bruta_Vegas (R$)","Imposto_sobre_Receita (R$)","MDR Líquido Vegas (R$)"]
    st.dataframe(filt[cols_to_show])

    from io import BytesIO
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        filt.to_excel(writer, sheet_name="Lancamentos_Filtrados", index=False)
        resumo.to_excel(writer, sheet_name="Resumo_Mensal", index=False)
    st.download_button("📤 Baixar Excel (filtro atual)", data=buffer.getvalue(),
                       file_name="vegas_pay_resumo.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Faça upload das duas planilhas para ver o painel.")
