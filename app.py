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

st.set_page_config(layout="wide")

# ================= BANNER =================
st.markdown("""
<div style='background:#1F3A5F;padding:25px'>
<h1 style='color:white'>Strategic Supplier Decision Tool</h1>
<p style='color:white'>Integrated Kraljic & RAQSCI Analysis for Procurement Leaders</p>
</div>
""", unsafe_allow_html=True)

# ================= CONFIG =================
SECTORES = [
"Industrial / Manufacturing","Energía / Utilities","Construcción / Ingeniería",
"Automoción","Aeroespacial / Defensa","Alimentación / FMCG",
"Retail / Moda / eCommerce","Farmacéutico / Healthcare",
"Tecnología / Electrónica","Logística / Transporte",
"Servicios / Outsourcing","Químico / Materias primas","Otros"
]

mapping = {"Bajo (1)":1,"Medio (3)":3,"Alto (5)":5}

# ================= PESOS =================
def get_weights(sector):
    s = sector.lower()
    if "retail" in s: return {"R":0.05,"A":0.20,"Q":0.30,"S":0.20,"C":0.15,"I":0.10}
    if "industrial" in s: return {"R":0.10,"A":0.35,"Q":0.25,"S":0.10,"C":0.15,"I":0.05}
    if "farmacéutico" in s: return {"R":0.35,"A":0.25,"Q":0.25,"S":0.05,"C":0.05,"I":0.05}
    return {"R":0.10,"A":0.25,"Q":0.20,"S":0.15,"C":0.20,"I":0.10}

# ================= EXCEL =================
def generate_excel():
    wb = Workbook()
    ws = wb.active

    headers = [
        "Proveedor","Categoria","Subcategoria",
        "R_certificaciones","R_cumplimiento","R_ESG",
        "A_capacidad","A_dependencia","A_resiliencia",
        "Q_defectos","Q_consistencia","Q_auditorias",
        "S_leadtime","S_flexibilidad","S_soporte",
        "C_precio","C_logistica","C_TCO",
        "I_mejora","I_ID","I_digitalizacion"
    ]

    ws.append(headers)

    for i in range(1,len(headers)+1):
        ws.column_dimensions[get_column_letter(i)].width = 22

    dv = DataValidation(type="list", formula1='"Bajo (1),Medio (3),Alto (5)"')
    ws.add_data_validation(dv)

    for r in range(2,200):
        for c in range(4,len(headers)+1):
            dv.add(ws.cell(row=r,column=c))

    ws2 = wb.create_sheet("GUIA")
    ws2.append(["CRITERIO","BAJO","MEDIO","ALTO"])
    ws2.append(["Calidad",">5% defectos","~2%","<1%"])
    ws2.append(["Leadtime",">8","~4","<2"])
    ws2.append(["Dependencia","1 proveedor","2","3+"])

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ================= VALIDACIÓN =================
def validate(df):
    required = ["Proveedor","Categoria","Subcategoria"]
    for col in required:
        if col not in df.columns:
            st.error(f"Falta columna: {col}")
            st.stop()

# ================= PROCESS =================
def process(df,sector):

    for col in df.columns[3:]:
        df[col] = df[col].map(mapping)

    if df.iloc[:,3:].isnull().any().any():
        st.error("Valores inválidos en Excel")
        st.stop()

    w = get_weights(sector)
    data=[]

    for _,row in df.iterrows():

        r=row[3:6].mean()
        a=row[6:9].mean()
        q=row[9:12].mean()
        s=row[12:15].mean()
        c=row[15:18].mean()
        i=row[18:21].mean()

        score = r*w["R"]+a*w["A"]+q*w["Q"]+s*w["S"]+c*w["C"]+i*w["I"]

        impacto=(q+c)/2
        riesgo=a

        if impacto>=4 and riesgo>=4: k="Estratégica"
        elif impacto>=4: k="Apalancada"
        elif riesgo>=4: k="Cuello botella"
        else: k="No crítico"

        data.append({
            "Proveedor":row["Proveedor"],
            "R":r,"A":a,"Q":q,"S":s,"C":c,"I":i,
            "Score":round(score,2),
            "Estado":"APTO" if r>=3 and q>=3 else "NO APTO",
            "Impacto":impacto,
            "Riesgo":riesgo,
            "Kraljic":k,
            "C_val":c,"Q_val":q
        })

    return pd.DataFrame(data)

# ================= KPIs =================
def render_kpis(df):
    total=len(df)
    aptos=len(df[df["Estado"]=="APTO"])
    no_aptos=total-aptos
    estrategicos=len(df[df["Kraljic"]=="Estratégica"])

    col1,col2,col3 = st.columns(3)
    col4,col5,col6 = st.columns(3)

    col1.metric("Proveedores",total)
    col2.metric("% APTOS",round(aptos/total*100,1))
    col3.metric("% NO APTOS",round(no_aptos/total*100,1))
    col4.metric("Score medio",round(df["Score"].mean(),2))
    col5.metric("Riesgo medio",round(df["Riesgo"].mean(),2))
    col6.metric("% Estratégicos",round(estrategicos/total*100,1))

# ================= MATRIZ =================
def plot_matrix(df):
    fig=go.Figure()

    colors={"Estratégica":"#E74C3C","Apalancada":"#2ECC71","Cuello botella":"#F1C40F","No crítico":"#3498DB"}

    for cat in df["Kraljic"].unique():
        sub=df[df["Kraljic"]==cat]
        fig.add_trace(go.Scatter(
            x=sub["Impacto"],y=sub["Riesgo"],
            mode="markers",
            marker=dict(size=12,color=colors[cat]),
            text=sub["Proveedor"],
            name=cat,
            hovertemplate="<b>%{text}</b>"
        ))

    fig.add_shape(type="line",x0=3,x1=3,y0=0,y1=5)
    fig.add_shape(type="line",x0=0,x1=5,y0=3,y1=3)

    fig.add_annotation(x=4.5,y=4.5,text="Estratégica",showarrow=False)
    fig.add_annotation(x=4.5,y=1,text="Apalancada",showarrow=False)
    fig.add_annotation(x=1,y=4.5,text="Cuello botella",showarrow=False)
    fig.add_annotation(x=1,y=1,text="No crítico",showarrow=False)

    fig.update_layout(xaxis_title="Impacto",yaxis_title="Riesgo")
    return fig

# ================= INSIGHTS =================
def generate_insights(df):

    insights=[]

    if df["Riesgo"].mean()>4:
        insights.append("Riesgo de suministro elevado")

    if len(df)<=2:
        insights.append("Dependencia crítica de proveedores")

    if len(df[df["Kraljic"]=="Estratégica"])/len(df)>0.5:
        insights.append("Alta concentración en proveedores estratégicos")

    if df["Score"].mean()<3:
        insights.append("Desempeño global bajo")

    incoherencias=df[(df["C_val"]>=4)&(df["Q_val"]<=2)]
    if len(incoherencias)>0:
        insights.append("Incoherencias coste-calidad detectadas")

    return insights

# ================= ESTRATEGIA =================
def generate_strategy(df):

    estrategia=[]

    for cat in ["Estratégica","Apalancada","Cuello botella","No crítico"]:
        if len(df[df["Kraljic"]==cat])>0:
            estrategia.append(f"Acción sobre categoría {cat}")

    return estrategia

# ================= PDF =================
def generate_pdf(df,fig):

    styles=getSampleStyleSheet()
    buffer=tempfile.NamedTemporaryFile(delete=False,suffix=".pdf")
    doc=SimpleDocTemplate(buffer.name,pagesize=A4)

    elems=[]

    elems.append(Paragraph("Strategic Supplier Decision Tool",styles['Title']))
    elems.append(PageBreak())

    elems.append(Paragraph(f"Proveedores: {len(df)}",styles['Normal']))
    elems.append(Paragraph(f"Score medio: {round(df['Score'].mean(),2)}",styles['Normal']))
    elems.append(PageBreak())

    img=tempfile.NamedTemporaryFile(delete=False,suffix=".png").name
    fig.write_image(img)
    elems.append(Image(img,width=450,height=300))

    doc.build(elems)

    return open(buffer.name,"rb").read()

# ================= UI =================
st.info("Descarga la plantilla → complétala → súbela → analiza → descarga informe")

sector = st.selectbox("Sector",SECTORES)

with st.sidebar:
    st.download_button("Descargar plantilla", generate_excel(),"plantilla.xlsx")
    file = st.file_uploader("Upload Excel")

if file:

    df_raw = pd.read_excel(file)
    validate(df_raw)

    df = process(df_raw,sector)

    render_kpis(df)

    col1,col2 = st.columns([2,1])

    with col1:
        fig = plot_matrix(df)
        st.plotly_chart(fig,use_container_width=True)

    with col2:
        st.dataframe(df.sort_values(by="Score",ascending=False))

    st.subheader("RAQSCI")
    st.dataframe(df[["Proveedor","R","A","Q","S","C","I","Score"]])

    st.subheader("Insights")
    for i in generate_insights(df):
        st.warning(i)

    st.subheader("Estrategia")
    for s in generate_strategy(df):
        st.write("- "+s)

    pdf=generate_pdf(df,fig)
    st.download_button("Descargar PDF",pdf)

# ================= MODO PRESENTACIÓN =================
modo_presentacion = st.toggle("Modo presentación")

if modo_presentacion and "df" in locals():

    if "slide" not in st.session_state:
        st.session_state.slide = 0

    def next_slide():
        st.session_state.slide += 1

    def prev_slide():
        st.session_state.slide -= 1

    col1, col2, col3 = st.columns([1,2,1])

    with col1:
        if st.button("⬅"):
            prev_slide()

    with col3:
        if st.button("➡"):
            next_slide()

    slide = st.session_state.slide

    if slide == 1:
        render_kpis(df)
    elif slide == 2:
        st.plotly_chart(fig)
