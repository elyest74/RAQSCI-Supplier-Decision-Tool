import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import tempfile

from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

# -------------------------------
# CONFIG
# -------------------------------
st.set_page_config(page_title="RAQSCI Tool", layout="wide")

# -------------------------------
# SECTOR (usuario)
# -------------------------------
SECTORES = [
    "Industrial / Manufacturing",
    "Energía / Utilities",
    "Construcción / Ingeniería",
    "Automoción",
    "Aeroespacial / Defensa",
    "Alimentación / FMCG",
    "Retail / Moda / eCommerce",
    "Farmacéutico / Healthcare",
    "Tecnología / Electrónica",
    "Logística / Transporte",
    "Servicios / Outsourcing",
    "Químico / Materias primas",
    "Otros"
]

# -------------------------------
# PESOS RAQSCI POR SECTOR
# -------------------------------
def get_sector_weights(sector):
    s = sector.lower()

    # default equilibrado
    w = {"R":0.10,"A":0.25,"Q":0.20,"S":0.15,"C":0.20,"I":0.10}

    if "retail" in s or "moda" in s:
        w = {"R":0.05,"A":0.20,"Q":0.30,"S":0.20,"C":0.15,"I":0.10}
    elif "industrial" in s:
        w = {"R":0.10,"A":0.35,"Q":0.25,"S":0.10,"C":0.15,"I":0.05}
    elif "farmacéutico" in s:
        w = {"R":0.30,"A":0.25,"Q":0.25,"S":0.10,"C":0.05,"I":0.05}

    return w

# -------------------------------
# RAQSCI
# -------------------------------
def calculate_raqsci_score(r,a,q,s,c,i,sector):

    w = get_sector_weights(sector)

    score = (
        r*w["R"] +
        a*w["A"] +
        q*w["Q"] +
        s*w["S"] +
        c*w["C"] +
        i*w["I"]
    )

    if r<3 or q<3 or a<3:
        return round(score,2),"NO APTO","Fallo crítico"

    if a<=2:
        score *= 0.85

    alert = None
    if c>=4.5 and (q<=2.5 or a<=2.5):
        alert = "Riesgo compra barata"

    return round(score,2),"APTO",alert

# -------------------------------
# IMPACTO / RIESGO (Kraljic)
# -------------------------------
def calculate_impact_risk(q,c,a,sector):

    w = get_sector_weights(sector)

    impacto = (q*w["Q"] + c*w["C"]) / (w["Q"] + w["C"])
    riesgo = a

    return impacto, riesgo

def classify_kraljic(impacto,riesgo):

    if impacto>=4 and riesgo>=4:
        return "Estratégica"
    elif impacto>=4:
        return "Apalancada"
    elif riesgo>=4:
        return "Cuello de botella"
    else:
        return "No crítico"

# -------------------------------
# EXCEL TEMPLATE
# -------------------------------
def generate_excel():
    wb=Workbook()
    ws=wb.active
    ws.title="RAQSCI_INPUT"

    headers=[
        "Proveedor","Categoria","Subcategoria",
        "R_certificaciones","R_cumplimiento","R_ESG",
        "A_capacidad","A_dependencia","A_resiliencia",
        "Q_defectos","Q_consistencia","Q_auditorias",
        "S_leadtime","S_flexibilidad","S_soporte",
        "C_precio","C_logistica","C_inventario","C_TCO",
        "I_mejora","I_ID","I_digitalizacion"
    ]

    ws.append(headers)

    ws.append([
        "Proveedor Demo","Packaging","Caja ecommerce",
        4,5,3,5,4,4,4,5,4,3,4,4,4,3,3,4,3,3,4
    ])

    file=io.BytesIO()
    wb.save(file)
    return file.getvalue()

# -------------------------------
# PROCESAR EXCEL
# -------------------------------
def process_file(file, sector):

    df = pd.read_excel(file)
    data=[]

    for _,row in df.iterrows():

        r=(row["R_certificaciones"]+row["R_cumplimiento"]+row["R_ESG"])/3
        a=(row["A_capacidad"]+row["A_dependencia"]+row["A_resiliencia"])/3
        q=(row["Q_defectos"]+row["Q_consistencia"]+row["Q_auditorias"])/3
        s=(row["S_leadtime"]+row["S_flexibilidad"]+row["S_soporte"])/3
        c=(row["C_precio"]+row["C_logistica"]+row["C_inventario"]+row["C_TCO"])/4
        i=(row["I_mejora"]+row["I_ID"]+row["I_digitalizacion"])/3

        score,status,alert = calculate_raqsci_score(r,a,q,s,c,i,sector)

        impacto,riesgo = calculate_impact_risk(q,c,a,sector)
        krajlic = classify_kraljic(impacto,riesgo)

        data.append({
            "Proveedor":row["Proveedor"],
            "Categoria":row["Categoria"],
            "Subcategoria":row["Subcategoria"],
            "R":round(r,2),"A":round(a,2),"Q":round(q,2),
            "S":round(s,2),"C":round(c,2),"I":round(i,2),
            "Score":score,
            "Estado":status,
            "Kraljic":krajlic,
            "Impacto":round(impacto,2),
            "Riesgo":round(riesgo,2),
            "Alerta":alert
        })

    return pd.DataFrame(data)

# -------------------------------
# DIAGNÓSTICO AVANZADO
# -------------------------------
def generate_insights(df):

    insights=[]
    top=df.iloc[0]

    insights.append(f"Proveedor líder: {top['Proveedor']} con score {top['Score']}")

    if top["Riesgo"]>=4:
        insights.append("Alto riesgo de suministro")

    if top["C"]>=4 and top["Q"]<=3:
        insights.append("Posible incoherencia: bajo coste vs calidad")

    if len(df)<=2:
        insights.append("Dependencia de proveedores limitada")

    return insights

def generate_actions(df):

    top=df.iloc[0]
    actions=[]

    if top["Kraljic"]=="Estratégica":
        actions+=["Desarrollar relación a largo plazo","Colaboración estratégica"]

    elif top["Kraljic"]=="Apalancada":
        actions+=["Negociación activa","Competencia entre proveedores"]

    elif top["Kraljic"]=="Cuello de botella":
        actions+=["Asegurar suministro","Buscar alternativas"]

    else:
        actions+=["Optimizar proceso","Reducir complejidad"]

    return actions

# -------------------------------
# MATRIZ KRALJIC
# -------------------------------
def plot_matrix(df):

    fig=go.Figure()

    colores={
        "Estratégica":"#E74C3C",
        "Apalancada":"#2ECC71",
        "Cuello de botella":"#F1C40F",
        "No crítico":"#3498DB"
    }

    for cat in df["Kraljic"].unique():
        sub=df[df["Kraljic"]==cat]

        fig.add_trace(go.Scatter(
            x=sub["Impacto"],
            y=sub["Riesgo"],
            mode="markers",
            name=cat,
            marker=dict(size=12,color=colores[cat]),
            text=sub["Proveedor"],
            hovertemplate="<b>%{text}</b><br>Impacto:%{x}<br>Riesgo:%{y}<extra></extra>"
        ))

    fig.add_shape(type="line",x0=3,x1=3,y0=0,y1=5,line=dict(dash="dash"))
    fig.add_shape(type="line",x0=0,x1=5,y0=3,y1=3,line=dict(dash="dash"))

    fig.update_layout(
        legend=dict(x=1.05,y=1),
        margin=dict(r=120)
    )

    return fig

# -------------------------------
# PDF
# -------------------------------
def generate_pdf(df,fig):

    styles=getSampleStyleSheet()
    buffer=tempfile.NamedTemporaryFile(delete=False,suffix=".pdf")
    doc=SimpleDocTemplate(buffer.name,pagesize=A4)

    elems=[]
    top=df.iloc[0]

    # portada
    elems.append(Spacer(1,100))
    elems.append(Paragraph("Strategic Supplier Decision Tool",styles['Title']))
    elems.append(Spacer(1,20))
    elems.append(Paragraph("Integrated Kraljic & RAQSCI Analysis",styles['Heading2']))
    elems.append(Spacer(1,200))
    elems.append(Paragraph("Desarrollado por Elymar Estévez",styles['Normal']))
    elems.append(PageBreak())

    # resumen
    elems.append(Paragraph("Executive Summary",styles['Heading2']))
    elems.append(Paragraph(f"Proveedor líder: {top['Proveedor']} | Score: {top['Score']}",styles['Normal']))
    elems.append(Spacer(1,20))

    # insights
    elems.append(Paragraph("Key Insights",styles['Heading2']))
    for i in generate_insights(df):
        elems.append(Paragraph(f"• {i}",styles['Normal']))

    elems.append(Spacer(1,20))

    # recomendación
    elems.append(Paragraph("Recommendation",styles['Heading2']))
    elems.append(Paragraph("Se recomienda adjudicación con control de riesgos.",styles['Normal']))
    elems.append(PageBreak())

    # matriz
    img=tempfile.NamedTemporaryFile(delete=False,suffix=".png").name
    fig.write_image(img)
    elems.append(Image(img,width=450,height=320))
    elems.append(PageBreak())

    # tabla
    data=[df.columns.tolist()]+df.values.tolist()
    table=Table(data)
    elems.append(table)

    elems.append(Spacer(1,30))
    elems.append(Paragraph("Desarrollado por Elymar Estévez",styles['Normal']))

    doc.build(elems)

    return open(buffer.name,"rb").read()

# -------------------------------
# UI
# -------------------------------

# Banner
st.markdown("""
<div style='background-color:#1F3A5F;padding:20px;border-radius:8px'>
<h1 style='color:white'>Strategic Supplier Decision Tool</h1>
<p style='color:white'>Integrated Kraljic & RAQSCI Analysis for Procurement Leaders</p>
</div>
""", unsafe_allow_html=True)

# Sector selector
sector = st.selectbox("Selecciona el sector de tu empresa", SECTORES)

# Instrucciones
st.markdown("""
### Cómo utilizar la herramienta:

1. Descarga la plantilla Excel desde el panel lateral  
2. Completa la información de proveedores siguiendo la estructura definida  
3. Sube el fichero utilizando el botón Upload  
4. La aplicación realizará automáticamente:  
   - Evaluación RAQSCI  
   - Clasificación Kraljic  
   - Identificación de riesgos  
   - Recomendación de decisión  
5. Analiza los resultados y descarga el informe en PDF  
""")

# Sidebar
with st.sidebar:
    st.image("elymar.png", width=120)
    st.download_button("📥 Descargar plantilla", generate_excel(), "plantilla.xlsx")
    uploaded = st.file_uploader("📤 Upload Excel", type=["xlsx"])

# Procesamiento
if uploaded:
    df = process_file(uploaded, sector)
    st.dataframe(df)

    fig = plot_matrix(df)
    st.plotly_chart(fig, use_container_width=True)

    pdf = generate_pdf(df, fig)
    st.download_button("📄 Descargar informe PDF", pdf, "report.pdf")
