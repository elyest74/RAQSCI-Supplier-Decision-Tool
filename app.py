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
# SESSION
# -------------------------------
if "proveedores" not in st.session_state:
    st.session_state.proveedores = []

# -------------------------------
# RAQSCI
# -------------------------------
def calculate_raqsci_score(r,a,q,s,c,i,strategy):

    weights = {
        "Estratégica":[0.05,0.30,0.25,0.10,0.15,0.15],
        "Apalancada":[0.10,0.15,0.20,0.15,0.35,0.05],
        "Cuello de botella":[0.10,0.35,0.20,0.20,0.10,0.05],
        "No crítico":[0.10,0.10,0.15,0.20,0.40,0.05]
    }

    w = weights[strategy]
    score = (r*w[0]+a*w[1]+q*w[2]+s*w[3]+c*w[4]+i*w[5])

    if r<3 or q<3 or a<3:
        return round(score,2),"NO APTO","Fallo crítico"

    if a<=2:
        score*=0.85

    alert=None
    if c==5 and (q<=2 or a<=2):
        alert="Riesgo compra barata"

    return round(score,2),"APTO",alert

# -------------------------------
# KRALJIC
# -------------------------------
def calculate_impact_risk(q,c,a):
    impacto=(q+c)/2
    riesgo=a
    return impacto,riesgo

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

    headers=["Proveedor","Sector","Categoria","Subcategoria",
             "R_certificaciones","R_cumplimiento","R_ESG",
             "A_capacidad","A_dependencia","A_resiliencia",
             "Q_defectos","Q_consistencia","Q_auditorias",
             "S_leadtime","S_flexibilidad","S_soporte",
             "C_precio","C_logistica","C_inventario","C_TCO",
             "I_mejora","I_ID","I_digitalizacion"]

    ws.append(headers)

    ws.append(["Proveedor Demo","Industrial","Metal","Acero",
               4,5,3,5,4,4,4,5,4,3,4,4,4,3,3,4,3,3,4])

    file=io.BytesIO()
    wb.save(file)
    return file.getvalue()

# -------------------------------
# LOAD EXCEL
# -------------------------------
def process_file(file):

    df=pd.read_excel(file)

    data=[]

    for _,row in df.iterrows():

        r=(row["R_certificaciones"]+row["R_cumplimiento"]+row["R_ESG"])/3
        a=(row["A_capacidad"]+row["A_dependencia"]+row["A_resiliencia"])/3
        q=(row["Q_defectos"]+row["Q_consistencia"]+row["Q_auditorias"])/3
        s=(row["S_leadtime"]+row["S_flexibilidad"]+row["S_soporte"])/3
        c=(row["C_precio"]+row["C_logistica"]+row["C_inventario"]+row["C_TCO"])/4
        i=(row["I_mejora"]+row["I_ID"]+row["I_digitalizacion"])/3

        score,status,alert=calculate_raqsci_score(r,a,q,s,c,i,"Estratégica")

        impacto,riesgo=calculate_impact_risk(q,c,a)
        krajlic=classify_kraljic(impacto,riesgo)

        data.append({
            "Proveedor":row["Proveedor"],
            "Impacto":round(impacto,2),
            "Riesgo":round(riesgo,2),
            "Score":score,
            "Estado":status,
            "Kraljic":krajlic,
            "Alerta":alert
        })

    return pd.DataFrame(data)

# -------------------------------
# INSIGHTS
# -------------------------------
def insights(df):
    top=df.iloc[0]
    return [f"Proveedor líder: {top['Proveedor']} con score {top['Score']}"]

def actions(df):
    return ["Desarrollar relación estratégica"]

# -------------------------------
# MATRIX
# -------------------------------
def plot_matrix(df):
    fig=go.Figure()

    colors_map={
        "Estratégica":"red",
        "Apalancada":"green",
        "Cuello de botella":"orange",
        "No crítico":"blue"
    }

    for cat in df["Kraljic"].unique():
        sub=df[df["Kraljic"]==cat]
        fig.add_trace(go.Scatter(
            x=sub["Impacto"],
            y=sub["Riesgo"],
            mode="markers",
            name=cat,
            marker=dict(size=12,color=colors_map[cat]),
            text=sub["Proveedor"],
            hovertemplate="%{text}"
        ))

    fig.add_shape(type="line",x0=3,x1=3,y0=0,y1=5)
    fig.add_shape(type="line",x0=0,x1=5,y0=3,y1=3)

    fig.update_layout(legend=dict(x=1.1,y=1))

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
    elems.append(Paragraph("RAQSCI Supplier Report",styles['Title']))
    elems.append(Spacer(1,200))
    elems.append(Paragraph("Desarrollado por Estevez Procurement Advisor",styles['Normal']))
    elems.append(PageBreak())

    # resumen
    elems.append(Paragraph("Executive Summary",styles['Heading2']))
    elems.append(Paragraph(f"Proveedor líder: {top['Proveedor']}",styles['Normal']))
    elems.append(Spacer(1,20))

    # recomendación
    elems.append(Paragraph("Recommendation",styles['Heading2']))
    elems.append(Paragraph("Se recomienda adjudicación.",styles['Normal']))
    elems.append(PageBreak())

    # matriz
    img=tempfile.NamedTemporaryFile(delete=False,suffix=".png").name
    fig.write_image(img)
    elems.append(Image(img,width=400,height=300))
    elems.append(PageBreak())

    # tabla
    data=[df.columns.tolist()]+df.values.tolist()
    table=Table(data)
    elems.append(table)

    elems.append(Spacer(1,30))
    elems.append(Paragraph("Desarrollado por Estevez Procurement Advisor",styles['Normal']))

    doc.build(elems)

    return open(buffer.name,"rb").read()

# -------------------------------
# UI
# -------------------------------
st.title("Strategic Supplier Decision Tool | RAQSCI")

with st.sidebar:
    st.image("elymar.png",width=120)

    st.download_button("Descargar Excel",generate_excel(),"plantilla.xlsx")

    uploaded=st.file_uploader("Cargar Excel",type=["xlsx"])

if uploaded:
    df=process_file(uploaded)
    st.dataframe(df)

    fig=plot_matrix(df)
    st.plotly_chart(fig)

    pdf=generate_pdf(df,fig)
    st.download_button("Descargar PDF",pdf,"report.pdf")
