import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io

# -------------------------------
# CONFIGURACIÓN GENERAL
# -------------------------------
st.set_page_config(page_title="RAQSCI Tool", layout="wide")

# -------------------------------
# ESTILOS
# -------------------------------
st.markdown("""
<style>
.main { background-color: #F5F7FA; }
h1 { color: #1F3A5F; }
h2, h3 { color: #2E5C8A; }
</style>
""", unsafe_allow_html=True)

# -------------------------------
# SESSION STATE
# -------------------------------
if "proveedores" not in st.session_state:
    st.session_state.proveedores = []

# -------------------------------
# FUNCIÓN RAQSCI
# -------------------------------
def calculate_raqsci_score(r, a, q, s, c, i, strategy):

    weights = {
        "Estratégica":      [0.05, 0.30, 0.25, 0.10, 0.15, 0.15],
        "Apalancada":       [0.10, 0.15, 0.20, 0.15, 0.35, 0.05],
        "Cuello de botella":[0.10, 0.35, 0.20, 0.20, 0.10, 0.05],
        "No crítico":       [0.10, 0.10, 0.15, 0.20, 0.40, 0.05]
    }

    w = weights[strategy]

    score = (r*w[0] + a*w[1] + q*w[2] + s*w[3] + c*w[4] + i*w[5])

    # Reglas críticas
    if r < 3 or q < 3 or a < 3:
        return round(score, 2), "NO APTO", "Fallo en criterio crítico"

    # Penalización
    if a <= 2:
        score *= 0.85

    # Alertas
    alert = None
    if c == 5 and (q <= 2 or a <= 2):
        alert = "Riesgo de compra barata"

    return round(score, 2), "APTO", alert

# -------------------------------
# EXCEL TEMPLATE
# -------------------------------
def generate_excel_template():
    columns = [
        "Proveedor","Categoria_Kraljic","Evaluador","Fecha","Comentarios",
        "Precio_unitario","Coste_logistico","Coste_inventario","Coste_total_estimado",
        "R_certificaciones","R_cumplimiento","R_ESG",
        "A_capacidad","A_dependencia","A_resiliencia",
        "Q_defectos","Q_consistencia","Q_auditorias",
        "S_leadtime","S_flexibilidad","S_soporte",
        "C_precio","C_logistica","C_inventario","C_TCO",
        "I_mejora","I_ID","I_digitalizacion"
    ]

    df = pd.DataFrame(columns=columns)

    example = {
        "Proveedor": "Proveedor Demo",
        "Categoria_Kraljic": "Estratégica",
        "Evaluador": "Elymar",
        "Fecha": "2026-01-01",
        "Comentarios": "Ejemplo",
        "Precio_unitario": 10,
        "Coste_logistico": 1,
        "Coste_inventario": 0.5,
        "Coste_total_estimado": 11.5,
        "R_certificaciones": 4,
        "R_cumplimiento": 5,
        "R_ESG": 3,
        "A_capacidad": 5,
        "A_dependencia": 4,
        "A_resiliencia": 4,
        "Q_defectos": 4,
        "Q_consistencia": 5,
        "Q_auditorias": 4,
        "S_leadtime": 3,
        "S_flexibilidad": 4,
        "S_soporte": 4,
        "C_precio": 4,
        "C_logistica": 3,
        "C_inventario": 3,
        "C_TCO": 4,
        "I_mejora": 3,
        "I_ID": 3,
        "I_digitalizacion": 4
    }

    df = pd.concat([df, pd.DataFrame([example])])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='RAQSCI_INPUT')

    return output.getvalue()

# -------------------------------
# HEADER
# -------------------------------
st.title("Strategic Supplier Decision Tool | RAQSCI")
st.caption("Structured supplier evaluation for smarter procurement decisions")

# -------------------------------
# SIDEBAR
# -------------------------------
with st.sidebar:

    st.image("elymar.png", width=120)
    st.markdown("**Desarrollado por Elymar Estévez**")

    st.markdown("---")

    tipo_compra = st.selectbox(
        "Tipo de Categoría",
        ["Estratégica", "Apalancada", "Cuello de botella", "No crítico"]
    )

    st.markdown("---")

    st.download_button(
        label="📥 Descargar plantilla Excel",
        data=generate_excel_template(),
        file_name="plantilla_RAQSci.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------------------------------
# LAYOUT
# -------------------------------
col_input, col_result = st.columns([1, 2])

# -------------------------------
# INPUT
# -------------------------------
with col_input:

    st.subheader("Evaluación de Proveedor")

    with st.form("form_proveedor"):
        nombre = st.text_input("Proveedor")

        r = st.slider("Regulación", 1, 5, 3)
        a = st.slider("Aseguramiento", 1, 5, 3)
        q = st.slider("Calidad", 1, 5, 3)
        s = st.slider("Servicio", 1, 5, 3)
        c = st.slider("Coste", 1, 5, 3)
        i = st.slider("Innovación", 1, 5, 3)

        submitted = st.form_submit_button("Añadir")

    if submitted and nombre != "":
        score, status, alert = calculate_raqsci_score(r, a, q, s, c, i, tipo_compra)

        st.session_state.proveedores.append({
            "Proveedor": nombre,
            "R": r, "A": a, "Q": q, "S": s, "C": c, "I": i,
            "Score": score,
            "Estado": status,
            "Alerta": alert
        })

# -------------------------------
# RESULTADOS
# -------------------------------
with col_result:

    st.subheader("Resultados")

    if len(st.session_state.proveedores) > 0:

        df = pd.DataFrame(st.session_state.proveedores)
        df = df.sort_values(by="Score", ascending=False)

        top = df.iloc[0]

        k1, k2, k3 = st.columns(3)
        k1.metric("Mejor proveedor", top["Proveedor"])
        k2.metric("Score", f"{top['Score']}/5")
        k3.metric("Estado", top["Estado"])

        st.markdown("### Comparativa")
        st.dataframe(df, use_container_width=True)

        # Alertas
        for _, row in df.iterrows():
            if row["Estado"] == "NO APTO":
                st.error(f"{row['Proveedor']} no cumple criterios mínimos")
            if row["Alerta"]:
                st.warning(f"{row['Proveedor']}: {row['Alerta']}")

        # Radar
        st.markdown("### Radar comparativo")

        categories = ['R', 'A', 'Q', 'S', 'C', 'I']
        fig = go.Figure()

        for _, row in df.iterrows():
            fig.add_trace(go.Scatterpolar(
                r=[row['R'], row['A'], row['Q'], row['S'], row['C'], row['I']],
                theta=categories,
                fill='toself',
                name=row["Proveedor"]
            ))

        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5]))
        )

        st.plotly_chart(fig, use_container_width=True)

    else:
        st.info("Añade proveedores para comenzar análisis")
