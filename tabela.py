import streamlit as st
import pandas as pd
import io
from PyPDF2 import PdfWriter, PdfReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import altair as alt

def load_data(csv_url):
    @st.cache_data(ttl=3600)
    def fetch_data(url):
        return pd.read_csv(url)
    return fetch_data(csv_url)

def filter_data(df, os_filter, status_filter, tecnico_filter, disciplina_filter, date_range):
    df["DATA RECEBIDO"] = pd.to_datetime(df["DATA RECEBIDO"], format="%d/%m/%Y", errors='coerce')
    
    if os_filter:
        df = df[df["OS"].astype(str).str.contains(os_filter, case=False, na=False)]
    if status_filter:
        df = df[df["STATUS*"].isin(status_filter)]
    if tecnico_filter:
        df = df[df["RESPONSAVEL TÉCNICO"].isin(tecnico_filter)]
    if disciplina_filter:
        df = df[df["DISCIPLINAS"].isin(disciplina_filter)]
    if date_range:
        start_date, end_date = date_range
        df = df[(df["DATA RECEBIDO"] >= start_date) & (df["DATA RECEBIDO"] <= end_date)]
    return df

def create_excel_download(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def generate_filled_pdf(template_path, data):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Helvetica", 8)

    # Function to draw wrapped text
    def draw_wrapped_text(canvas, text, x, y, max_width):
        lines = text.split("\n")
        for line in lines:
            while len(line) > max_width:
                split_index = line.rfind(" ", 0, max_width)
                if split_index == -1:
                    split_index = max_width
                wrapped_line = line[:split_index]
                line = line[split_index:].strip()
                canvas.drawString(x, y, wrapped_line)
                y -= 10
            canvas.drawString(x, y, line)
            y -= 10

    # Insert data at correct coordinates
    can.drawString(90, 670, f"{data['OS']}")
    can.drawString(140, 670, f"{data['NORMAL / URGENTE']}")
    can.drawString(210, 670, f"{data['PRAZO DE ATENDIMENTO']}")
    can.drawString(300, 750, f"{data['ID']}")
    can.drawString(280, 670, f"{data['CONTRATO']}")
    can.drawString(410, 670, f"{data['FISCAL']}")
    can.drawString(70, 640, f"{data['PRÉDIO']}")
    can.drawString(210, 630, f"{data['LOCAL']}")
    can.drawString(100, 610, f"{data['FONE||RAMAL']}")
    can.drawString(250, 610, f"{data['SOLICITANTE']}")
    can.drawString(70, 520, f"{data['OBSERVAÇÃO']}")
    draw_wrapped_text(can, f"{data['DESCRIÇÃO DETALHADA']}", 70, 530, 90)
    can.save()

    packet.seek(0)

    existing_pdf = PdfReader(open(template_path, "rb"))
    output = PdfWriter()
    new_pdf = PdfReader(packet)
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    output_stream = io.BytesIO()
    output.write(output_stream)
    output_stream.seek(0)

    return output_stream

def create_chart(df_filtered):
    df_filtered['MES'] = df_filtered['DATA RECEBIDO'].dt.to_period('M').astype(str)
    chart = alt.Chart(df_filtered).mark_bar().encode(
        x='MES',
        y='count()',
        color='STATUS*',
        tooltip=['MES', 'count()', 'RESPONSAVEL TÉCNICO', 'DISCIPLINAS']
    ).properties(
        width=800,
        height=400
    )
    return chart

def tabela():

    st.markdown(
        """
        <style>
        /* Style to enhance table appearance */
        .dataframe th, .dataframe td {
            padding: 10px;
            border: 1px solid #ddd;
        }
        .dataframe th {
            background-color: #f4f4f4;
            font-weight: bold;
        }
        .dataframe {
            border-collapse: collapse;
            width: 100%;
        }
        /* Page title style */
        h1 {
            text-align: center;
        }
        /* Search filter style */
        .stTextInput, .stMultiSelect, .stDateInput {
            margin-bottom: 20px;
        }
        /* Download button style */
        .stDownloadButton {
            color: white;
            border: none;
            padding: 10px 20px;
            text-align: left;
            font-size: 16px;
            margin-top: 20px;
            cursor: pointer;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTlBXGpJ6j2i-C6edJ-eB4X2DD-7KA7Ys1bIR-tCFeYt6B-7S30bcY_bd0TUtEbttDiMBtexpD-2C4-/pub?gid=1319816246&single=true&output=csv"
    with st.spinner('Carregando dados...'):
        df = load_data(url)

    st.title("Tabela Ordens de Serviços")

    col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 2, 2])

    with col1:
        os_filter = st.text_input("Pesquisar por OS", value="")

    with col2:
        if "STATUS*" in df.columns:
            status_options = df["STATUS*"].unique().tolist()
            status_filter = st.multiselect(
                "Selecionar STATUS*", options=status_options
            )
        else:
            st.error("A coluna 'STATUS*' não foi encontrada no DataFrame.")
            status_filter = []

    with col3:
        if "RESPONSAVEL TÉCNICO" in df.columns:
            tecnico_options = df["RESPONSAVEL TÉCNICO"].unique().tolist()
            tecnico_filter = st.multiselect(
                "Selecionar RESPONSAVEL TÉCNICO", options=tecnico_options
            )
        else:
            st.error("A coluna 'RESPONSAVEL TÉCNICO' não foi encontrada no DataFrame.")
            tecnico_filter = []

    with col4:
        if "DISCIPLINAS" in df.columns:
            disciplina_options = df["DISCIPLINAS"].unique().tolist()
            disciplina_filter = st.multiselect(
                "Selecionar DISCIPLINAS", options=disciplina_options
            )
        else:
            st.error("A coluna 'DISCIPLINAS' não foi encontrada no DataFrame.")
            disciplina_filter = []

    with col5:
        date_range = st.date_input("Selecionar Data Recebido", [])

    df_filtered = filter_data(df, os_filter, status_filter, tecnico_filter, disciplina_filter, date_range)

    st.dataframe(df_filtered, use_container_width=True)

    if not df_filtered.empty:
        st.subheader("Gráficos Interativos")
        st.altair_chart(create_chart(df_filtered), use_container_width=True)

        selected_data = df_filtered.iloc[0]
        template_path = "./mnt/data/RAT.pdf"

        col6, col7, col8 = st.columns([1, 1, 8], gap="small")
        
        with col6:
            st.download_button(
                label="Gerar XLSX",
                data=create_excel_download(df_filtered),
                file_name="dados_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        
        with col7:
            st.download_button(
                label="Gerar PDF",
                data=generate_filled_pdf(template_path, selected_data),
                file_name=f"OS_{selected_data['OS']}.pdf",
                mime="application/pdf",
            )

if __name__ == "__main__":
    tabela()
