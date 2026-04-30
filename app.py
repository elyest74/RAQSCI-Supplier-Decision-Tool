import streamlit as st
import pandas as pd
import plotly.graph_objects as go

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
.metric-card {
    background-color: white;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)

# -------------------------------
# DATOS INICIALES
# -------------------------------
if "proveedores" not in st.session_state:
    st.session_state.proveedores = []

# -------------------------------
# FUNCIÓN RAQSCI (MEJORADA)
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
# HEADER
# -------------------------------
st.title("Strategic Supplier Decision Tool | RAQSCI")
st.caption("Structured supplier evaluation for smarter procurement decisions")

# -------------------------------
# SIDEBAR
# -------------------------------
with st.sidebar:
    st.header("Configuración")
    tipo_compra = st.selectbox(
        "Tipo de Categoría (Kraljic)",
        ["Estratégica", "Apalancada", "Cuello de botella", "No crítico"]
    )

    st.markdown("---")
    st.image("elymar.png", width=120)
    st.markdown("**Desarrollado por Elymar Estévez**")

# -------------------------------
# LAYOUT
# -------------------------------
col_input, col_result = st.columns([1, 2])

# -------------------------------
# INPUT PROVEEDOR
# -------------------------------
with col_input:
    st.subheader("Evaluación de Proveedor")

    with st.form("form_proveedor"):
        nombre = st.text_input("Nombre del proveedor")

        r = st.slider("Regulación", 1, 5, 3)
        a = st.slider("Aseguramiento", 1, 5, 3)
        q = st.slider("Calidad", 1, 5, 3)
        s = st.slider("Servicio", 1, 5, 3)
        c = st.slider("Coste", 1, 5, 3)
        i = st.slider("Innovación", 1, 5, 3)

        submitted = st.form_submit_button("Añadir proveedor")

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

        # Ranking
        df_sorted = df.sort_values(by="Score", ascending=False)

        # KPIs principales
        top = df_sorted.iloc[0]

        k1, k2, k3 = st.columns(3)

        k1.metric("Mejor proveedor", top["Proveedor"])
        k2.metric("Score", f"{top['Score']}/5")

        if top["Estado"] == "APTO":
            k3.metric("Estado", "APTO", delta="Validado", delta_color="normal")
        else:
            k3.metric("Estado", "NO APTO", delta="Riesgo", delta_color="inverse")

        # Tabla
        st.markdown("### Comparativa de proveedores")
        st.dataframe(df_sorted, use_container_width=True)

        # Alertas
        for _, row in df_sorted.iterrows():
            if row["Alerta"]:
                st.warning(f"{row['Proveedor']}: {row['Alerta']}")
            if row["Estado"] == "NO APTO":
                st.error(f"{row['Proveedor']}: No cumple criterios mínimos")

        # Radar chart
        st.markdown("### Análisis comparativo")

        categories = ['R', 'A', 'Q', 'S', 'C', 'I']
        fig = go.Figure()

        for _, row in df_sorted.iterrows():
            fig.add_trace(go.Scatterpolar(
                r=[row['R'], row['A'], row['Q'], row['S'], row['C'], row['I']],
                theta=categories,
                fill='toself',
                name=row["Proveedor"]
            ))

        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            showlegend=True
        )

        st.plotly_chart(fig, use_container_width=True)

    else:
        st.info("Introduce proveedores para ver resultados.")
