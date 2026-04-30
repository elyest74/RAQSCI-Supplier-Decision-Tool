import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io, tempfile

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

# ================= CONFIG =================
st.set_page_config(layout="wide")

SECTORES = [
    "Industrial / Manufacturing","Energía / Utilities","Construcción / Ingeniería",
    "Automoción","Aeroespacial / Defensa","Alimentación / FMCG",
    "Retail / Moda / eCommerce","Farmacéutico / Healthcare",
    "Tecnología / Electrónica","Logística / Transporte",
    "Servicios / Outsourcing","Químico / Materias primas","Otros"
]

mapping = {"Bajo (1)":1,"Medio (3)":3,"Alto (5)":5}

# ================= VALIDACIÓN =================
def validate_data(df):
    errores = []
    required_cols = ["Proveedor","Categoria","Subcategoria"]

    for col in required_cols:
        if col not in df.columns:
            errores.append(f"Falta columna obligatoria: {col}")

    for i, row in df.iterrows():
        for col in df.columns[3:]:
            val = row[col]
            if pd.isna(val):
                errores.append(f"Fila {i+2}: vacío en {col}")
            elif val not in mapping:
                errores.append(f"Fila {i+2}: valor inválido '{val}' en {col}")

    return errores

# ================= PESOS =================
def get_weights(sector):
    s = sector.lower()
    if "retail" in s:
        return {"R":0.05,"A":0.20,"Q":0.30,"S":0.20,"C":0.15,"I":0.10}
    if "industrial" in s:
        return {"R":0.10,"A":0.35,"Q":0.25,"S":0.10,"C":0.15,"I":0.05}
    return {"R":0.10,"A":0.25,"Q":0.20,"S":0.15,"C":0.20,"I":0.10}

# ================= EXCEL =================
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "RAQSCI_INPUT"

    headers = ["Proveedor","Categoria","Subcategoria"] + [f"X{i}" for i in range(1,22)]
    ws.append(headers)

    dv = DataValidation(type="list", formula1='"Bajo (1),Medio (3),Alto (5)"')
    ws.add_data_validation(dv)

    for r in range(2,200):
        for c in range(4,len(headers)+1):
            dv.add(ws.cell(row=r,column=c))

    for i in range(1,len(headers)+1):
        ws.column_dimensions[get_column_letter(i)].width = 20

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ================= MODELO =================
def calc_raqsci(vals, sector):
    w = list(get_weights(sector).values())
    score = sum([vals[i]*w[i] for i in range(6)])
    status = "APTO" if min(vals[:3])>=3 else "NO APTO"
    return round(score,2), status

def calc_kraljic(q,c,a):
    impacto=(q+c)/2
    riesgo=a
    if impacto>=4 and riesgo>=4: return impacto,riesgo,"Estratégica"
    elif impacto>=4: return impacto,riesgo,"Apalancada"
    elif riesgo>=4: return impacto,riesgo,"Cuello botella"
    return impacto,riesgo,"No crítico"

# ================= PROCESS =================
def process(df,sector):

    for col in df.columns[3:]:
        df[col] = df[col].map(mapping)

    if df.iloc[:,3:].isnull().any().any():
        st.error("Valores inválidos tras transformación")
        st.stop()

    data=[]
    for _,row in df.iterrows():

        vals=[
            row[3:6].mean(),
            row[6:9].mean(),
            row[9:12].mean(),
            row[12:15].mean(),
            row[15:19].mean(),
            row[19:22].mean()
        ]

        score,status = calc_raqsci(vals,sector)
        impacto,riesgo,k = calc_kraljic(vals[2],vals[4],vals[1])

        data.append({
            "Proveedor":row["Proveedor"],
            "Score":score,
            "Estado":status,
            "Impacto":impacto,
            "Riesgo":riesgo,
            "Kraljic":k
        })

    return pd.DataFrame(data)

# ================= INSIGHTS =================
def generate_insights(df):

    insights=[]

    if df["Riesgo"].mean()>4:
        insights.append("Riesgo de suministro elevado")

    if len(df)<=2:
        insights.append("Dependencia crítica de proveedores")

    if len(df[df["Kraljic"]=="Estratégica"]) / len(df) > 0.5:
        insights.append("Alta concentración en proveedores estratégicos")

    return insights

# ================= NARRATIVA =================
def generate_narrative(df):

    mean_score = round(df["Score"].mean(),2)
    risk = round(df["Riesgo"].mean(),2)

    return f"""
El análisis muestra un nivel medio de desempeño de {mean_score}/5,
con un nivel de riesgo de suministro de {risk}/5.

La estructura de proveedores evidencia una distribución estratégica basada en la matriz Kraljic,
permitiendo identificar oportunidades de optimización y mitigación de riesgos.

Se recomienda priorizar proveedores con mejor equilibrio entre desempeño y riesgo,
alineando la estrategia de compras con los objetivos del negocio.
"""

# ================= DECISIÓN =================
def generate_recommendation(df):

    df_valid = df[df["Estado"]=="APTO"]

    if df_valid.empty:
        return "No existen proveedores aptos para adjudicación"

    df_valid = df_valid.copy()
    df_valid["DecisionScore"] = df_valid["Score"] - (df_valid["Riesgo"]*0.3)

    top = df_valid.sort_values(by="DecisionScore", ascending=False).iloc[0]

    return f"Se recomienda adjudicar al proveedor {top['Proveedor']} por presentar el mejor equilibrio entre desempeño y riesgo."

# ================= MATRIZ =================
def plot_matrix(df):

    fig = go.Figure()

    colors = {
        "Estratégica":"#E74C3C",
        "Apalancada":"#2ECC71",
        "Cuello botella":"#F1C40F",
        "No crítico":"#3498DB"
    }

    for cat in df["Kraljic"].unique():
        sub=df[df["Kraljic"]==cat]

        fig.add_trace(go.Scatter(
            x=sub["Impacto"],
            y=sub["Riesgo"],
            mode="markers",
            marker=dict(size=12,color=colors.get(cat,"grey")),
            text=sub["Proveedor"],
            name=cat,
            hovertemplate="<b>%{text}</b>"
        ))

    fig.add_shape(type="line",x0=3,x1=3,y0=0,y1=5)
    fig.add_shape(type="line",x0=0,x1=5,y0=3,y1=3)

    return fig

# ================= PDF =================
def generate_pdf(df, fig):

    styles = getSampleStyleSheet()
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=A4)

    elems=[]

    # portada
    elems.append(Paragraph("Strategic Supplier Decision Tool", styles['Title']))
    elems.append(Spacer(1,200))
    elems.append(Paragraph("Desarrollado por Elymar Estévez", styles['Normal']))
    elems.append(PageBreak())

    # resumen
    elems.append(Paragraph("Executive Summary", styles['Heading2']))
    elems.append(Paragraph(generate_narrative(df), styles['Normal']))
    elems.append(PageBreak())

    # insights
    elems.append(Paragraph("Key Insights", styles['Heading2']))
    for i in generate_insights(df):
        elems.append(Paragraph(f"- {i}", styles['Normal']))

    elems.append(PageBreak())

    # matriz
    img = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
    fig.write_image(img)
    elems.append(Image(img, width=450, height=300))

    elems.append(PageBreak())

    # tabla
    data = [df.columns.tolist()] + df.values.tolist()
    elems.append(Table(data))

    elems.append(PageBreak())

    # recomendación
    elems.append(Paragraph("Recommendation", styles['Heading2']))
    elems.append(Paragraph(generate_recommendation(df), styles['Normal']))

    doc.build(elems)

    return open(buffer.name,"rb").read()

# ================= UI =================
st.title("Strategic Supplier Decision Tool")

st.info("""
Descarga la plantilla → complétala → súbela → analiza → descarga informe
""")

sector = st.selectbox("Sector",SECTORES)

with st.sidebar:
    st.image("elymar.png", width=120)
    st.download_button("Descargar plantilla", generate_excel(), "plantilla.xlsx")
    file = st.file_uploader("Upload Excel")

if file:

    df_raw = pd.read_excel(file)

    errores = validate_data(df_raw)
    if errores:
        st.error(errores[:5])
        st.stop()

    df = process(df_raw,sector)

    col1,col2 = st.columns([2,1])

    with col1:
        fig = plot_matrix(df)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Ranking proveedores")
        st.dataframe(df.sort_values(by="Score", ascending=False))

    st.subheader("Insights")
    for i in generate_insights(df):
        st.warning(i)

    st.subheader("Narrativa ejecutiva")
    st.write(generate_narrative(df))

    st.subheader("Recomendación")
    st.success(generate_recommendation(df))

    pdf = generate_pdf(df, fig)
    st.download_button("Descargar informe PDF", pdf, "report.pdf")
