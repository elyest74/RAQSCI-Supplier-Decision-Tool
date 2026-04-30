import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import tempfile

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

# -------------------------------
# CONFIG
# -------------------------------
st.set_page_config(layout="wide")

# -------------------------------
# SECTORES
# -------------------------------
SECTORES = [
    "Industrial / Manufacturing","Energía / Utilities","Construcción / Ingeniería",
    "Automoción","Aeroespacial / Defensa","Alimentación / FMCG",
    "Retail / Moda / eCommerce","Farmacéutico / Healthcare",
    "Tecnología / Electrónica","Logística / Transporte",
    "Servicios / Outsourcing","Químico / Materias primas","Otros"
]

# -------------------------------
# MAPPING
# -------------------------------
mapping = {
    "Bajo (1)":1,
    "Medio (3)":3,
    "Alto (5)":5
}

# -------------------------------
# PESOS POR SECTOR
# -------------------------------
def get_sector_weights(sector):
    s = sector.lower()
    if "moda" in s or "retail" in s:
        return {"R":0.05,"A":0.20,"Q":0.30,"S":0.20,"C":0.15,"I":0.10}
    elif "industrial" in s:
        return {"R":0.10,"A":0.35,"Q":0.25,"S":0.10,"C":0.15,"I":0.05}
    elif "farmacéutico" in s:
        return {"R":0.30,"A":0.25,"Q":0.25,"S":0.10,"C":0.05,"I":0.05}
    return {"R":0.10,"A":0.25,"Q":0.20,"S":0.15,"C":0.20,"I":0.10}

# -------------------------------
# RAQSCI
# -------------------------------
def calculate_raqsci(r,a,q,s,c,i,sector):
    w = get_sector_weights(sector)
    score = r*w["R"] + a*w["A"] + q*w["Q"] + s*w["S"] + c*w["C"] + i*w["I"]
    status = "APTO" if r>=3 and q>=3 and a>=3 else "NO APTO"
    return round(score,2), status

# -------------------------------
# KRALJIC
# -------------------------------
def calculate_kraljic(q,c,a):
    impacto = (q+c)/2
    riesgo = a
    if impacto>=4 and riesgo>=4:
        return impacto, riesgo, "Estratégica"
    elif impacto>=4:
        return impacto, riesgo, "Apalancada"
    elif riesgo>=4:
        return impacto, riesgo, "Cuello de botella"
    else:
        return impacto, riesgo, "No crítico"

# -------------------------------
# EXCEL
# -------------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "RAQSCI_INPUT"

    headers = [
        "Proveedor","Categoria","Subcategoria",
        "R_certificaciones","R_cumplimiento","R_ESG",
        "A_capacidad","A_dependencia","A_resiliencia",
        "Q_defectos","Q_consistencia","Q_auditorias",
        "S_leadtime","S_flexibilidad","S_soporte",
        "C_precio","C_logistica","C_inventario","C_TCO",
        "I_mejora","I_ID","I_digitalizacion"
    ]

    ws.append(headers)

    # ancho columnas
    for i in range(1,len(headers)+1):
        ws.column_dimensions[get_column_letter(i)].width = 20

    # validación
    dv = DataValidation(type="list", formula1='"Bajo (1),Medio (3),Alto (5)"')
    ws.add_data_validation(dv)

    for row in range(2,200):
        for col in range(4,len(headers)+1):
            dv.add(ws.cell(row=row,column=col))

    ws.freeze_panes = "D2"

    # hoja guía
    ws2 = wb.create_sheet("COMO_PUNTUAR")

    ws2.append(["CRITERIO","CÓMO MEDIR","BAJO (1)","MEDIO (3)","ALTO (5)"])
    ws2.append(["Defectos","% defectos",">5%","~2%","<0.5%"])
    ws2.append(["Lead time","semanas",">8","~4","<2"])
    ws2.append(["Dependencia","nº proveedores","1","2","3+"])
    ws2.append(["Precio","vs mercado","alto","medio","bajo"])

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# -------------------------------
# PROCESAR
# -------------------------------
def process(df, sector):

    for col in df.columns[3:]:
        df[col] = df[col].map(mapping)

    results=[]

    for _,row in df.iterrows():

        r = row[3:6].mean()
        a = row[6:9].mean()
        q = row[9:12].mean()
        s = row[12:15].mean()
        c = row[15:19].mean()
        i = row[19:22].mean()

        score,status = calculate_raqsci(r,a,q,s,c,i,sector)
        impacto,riesgo,k = calculate_kraljic(q,c,a)

        results.append({
            "Proveedor":row["Proveedor"],
            "Score":score,
            "Estado":status,
            "Impacto":impacto,
            "Riesgo":riesgo,
            "Kraljic":k,
            "C":c,"Q":q
        })

    return pd.DataFrame(results)

# -------------------------------
# MATRIZ
# -------------------------------
def plot_matrix(df):
    fig = go.Figure()
    colors_map = {
        "Estratégica":"red",
        "Apalancada":"green",
        "Cuello de botella":"orange",
        "No crítico":"blue"
    }

    for cat in df["Kraljic"].unique():
        sub = df[df["Kraljic"]==cat]
        fig.add_trace(go.Scatter(
            x=sub["Impacto"], y=sub["Riesgo"],
            mode="markers", name=cat,
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
def generate_pdf(df, fig):

    styles = getSampleStyleSheet()
    buffer = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    doc = SimpleDocTemplate(buffer.name, pagesize=A4)

    elems=[]
    top = df.sort_values(by="Score",ascending=False).iloc[0]

    elems.append(Paragraph("Strategic Supplier Decision Tool",styles["Title"]))
    elems.append(Spacer(1,200))
    elems.append(Paragraph("Desarrollado por Elymar Estévez",styles["Normal"]))
    elems.append(PageBreak())

    elems.append(Paragraph("Executive Summary",styles["Heading2"]))
    elems.append(Paragraph(f"Proveedor recomendado: {top['Proveedor']}",styles["Normal"]))
    elems.append(PageBreak())

    img = tempfile.NamedTemporaryFile(delete=False,suffix=".png").name
    fig.write_image(img)
    elems.append(Image(img,width=400,height=300))
    elems.append(PageBreak())

    data=[df.columns.tolist()]+df.values.tolist()
    elems.append(Table(data))

    doc.build(elems)
    return open(buffer.name,"rb").read()

# -------------------------------
# UI
# -------------------------------
st.markdown("""
<div style='background:#1F3A5F;padding:20px'>
<h1 style='color:white'>Strategic Supplier Decision Tool</h1>
<p style='color:white'>Integrated Kraljic & RAQSCI Analysis</p>
</div>
""", unsafe_allow_html=True)

sector = st.selectbox("Sector",SECTORES)

st.markdown("""
### Instrucciones
1. Descarga plantilla  
2. Rellena  
3. Sube archivo  
4. Analiza resultados  
""")

with st.sidebar:
    st.download_button("Descargar Excel",generate_excel(),"plantilla.xlsx")
    file = st.file_uploader("Upload",type=["xlsx"])

if file:
    df_raw = pd.read_excel(file)
    df = process(df_raw,sector)

    st.dataframe(df)

    fig = plot_matrix(df)
    st.plotly_chart(fig)

    pdf = generate_pdf(df,fig)
    st.download_button("Descargar PDF",pdf,"report.pdf")
