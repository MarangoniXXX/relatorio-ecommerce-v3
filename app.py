import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage
from datetime import datetime, timedelta

st.set_page_config(page_title="Relatório de Pedidos", layout="wide")
st.title("\U0001F4C5 Relatório de Pedidos ")

uploaded_file = st.file_uploader(
    "Envie a planilha do relatório:", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    st.write("Abas encontradas:", xls.sheet_names)

    df = pd.read_excel(uploaded_file, sheet_name="Relatório Base", header=1)
    df.columns = df.columns.str.strip()

    st.success("Planilha carregada com sucesso!")
    st.write("Colunas encontradas:", df.columns.tolist())

    required_cols = ['Data Entrada', 'Nome Cliente',
                     'Tipo Cliente', 'Venda Líquida']
    if not all(col in df.columns for col in required_cols):
        st.error(
            "Colunas esperadas não encontradas. Verifique: Data Entrada, Nome Cliente, Tipo Cliente, Venda Líquida.")
        st.stop()

    df['Data Entrada'] = pd.to_datetime(df['Data Entrada'], errors='coerce')

    # Sidebar filtros
    st.sidebar.header("Filtros")
    clientes = df['Nome Cliente'].dropna().unique()
    tipos_cliente = df['Tipo Cliente'].dropna().unique()

    cliente_selecionado = st.sidebar.multiselect("Nome do Cliente", clientes)
    tipo_selecionado = st.sidebar.multiselect("Tipo de Cliente", tipos_cliente)

    ontem = datetime.today() - timedelta(days=1)
    data_inicio, data_fim = st.sidebar.date_input(
        "Período de Data Entrada",
        [ontem, ontem]
    )

    top_n = st.sidebar.selectbox(
        "Quantidade de Clientes para Gráfico:", [5, 10, 20])

    df_filtrado = df.copy()

    if cliente_selecionado:
        df_filtrado = df_filtrado[df_filtrado['Nome Cliente'].isin(
            cliente_selecionado)]
    if tipo_selecionado:
        df_filtrado = df_filtrado[df_filtrado['Tipo Cliente'].isin(
            tipo_selecionado)]
    if data_inicio and data_fim:
        df_filtrado = df_filtrado[(df_filtrado['Data Entrada'] >= pd.to_datetime(data_inicio)) &
                                  (df_filtrado['Data Entrada'] <= pd.to_datetime(data_fim))]

    # Dashboard
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Resultado filtrado:")
        st.dataframe(df_filtrado)

    with col2:
        faturamento_total = df_filtrado['Venda Líquida'].sum()
        st.metric(label="\U0001F4B0 Faturamento Líquido Total",
                  value=f"R$ {faturamento_total:,.2f}")

    # Gráficos
    if not df_filtrado.empty:
        st.subheader(
            "\U0001F4CA Faturamento por Cliente (Top {})".format(top_n))

        top_clientes = df_filtrado.groupby('Nome Cliente')[
            'Venda Líquida'].sum().sort_values(ascending=False).head(top_n)

        fig1, ax1 = plt.subplots()
        top_clientes.plot(kind='barh', ax=ax1, color='mediumseagreen')
        ax1.set_xlabel("Venda Líquida (R$)")
        ax1.set_ylabel("Cliente")
        ax1.set_title("Top {} Clientes por Venda Líquida".format(top_n))
        ax1.invert_yaxis()
        st.pyplot(fig1)

        st.subheader("\U0001F967 Participação no Faturamento")
        fig2, ax2 = plt.subplots()
        top_clientes.plot(kind='pie', ax=ax2, autopct='%1.1f%%', startangle=90)
        ax2.set_ylabel('')
        ax2.set_title(
            "Participação no Faturamento dos Top {} Clientes".format(top_n))
        st.pyplot(fig2)
    else:
        st.warning("Nenhum dado para exibir no gráfico.")

    # Download e envio
    st.subheader("\U0001F4C5 Relatórios e E-mails")

    if st.button("\U0001F4C4 Gerar Relatório PDF"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=14)
        pdf.cell(
            0, 10, txt="Relatório de Faturamento Diário - Consultores", ln=True, align='C')
        pdf.cell(
            0, 10, txt=f"Gerado em: {datetime.today().strftime('%d/%m/%Y')}", ln=True, align='C')
        pdf.ln(10)

        pdf.set_font("Arial", size=12)
        pdf.cell(60, 10, "Nome Cliente", 1)
        pdf.cell(40, 10, "Tipo Cliente", 1)
        pdf.cell(50, 10, "Venda Líquida (R$)", 1)
        pdf.ln()

        for _, row in df_filtrado.iterrows():
            pdf.cell(60, 10, str(row['Nome Cliente'])[:30], 1)
            pdf.cell(40, 10, str(row['Tipo Cliente']), 1)
            pdf.cell(50, 10, f"R$ {row['Venda Líquida']:,.2f}", 1)
            pdf.ln()

        pdf.cell(100, 10, "TOTAL", 1)
        pdf.cell(50, 10, f"R$ {faturamento_total:,.2f}", 1)
        pdf.ln()

        pdf_output = pdf.output(dest='S').encode('latin1')

        st.download_button(
            label="⬇️ Baixar Relatório PDF",
            data=pdf_output,
            file_name="relatorio_pedidos.pdf",
            mime="application/pdf"
        )

        excel_output = io.BytesIO()
        df_filtrado.to_excel(excel_output, index=False)
        excel_output.seek(0)

        st.download_button(
            label="⬇️ Baixar Relatório Excel",
            data=excel_output,
            file_name="relatorio_pedidos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("\U0001F4E7 Enviar Relatório por E-mail (Outlook)")
    destinatarios = st.text_input("Digite os e-mails separados por vírgula:")
    email_outlook = st.text_input(
        "Seu e-mail Outlook", placeholder="voce@outlook.com")
    senha_email = st.text_input("Sua senha Outlook (segura)", type="password")
    corpo_email = st.text_area(
        "Mensagem do e-mail:", "Prezados, segue em anexo o relatório de vendas referente ao dia.")

    if st.button("Enviar E-mails") and destinatarios and email_outlook and senha_email:
        try:
            lista_emails = [email.strip()
                            for email in destinatarios.split(",") if email.strip()]

            msg = EmailMessage()
            msg['Subject'] = 'Relatório de Pedidos'
            msg['From'] = email_outlook
            msg.set_content(corpo_email)

            msg.add_attachment(pdf_output, maintype='application',
                               subtype='pdf', filename='relatorio.pdf')

            with smtplib.SMTP('smtp.office365.com', 587) as smtp:
                smtp.starttls()
                smtp.login(email_outlook, senha_email)
                for destinatario in lista_emails:
                    msg['To'] = destinatario
                    smtp.send_message(msg)
                    del msg['To']

            st.success("E-mails enviados com sucesso!")

        except Exception as e:
            st.error(f"Erro ao enviar e-mails: {e}")

else:
    st.info("Envie uma planilha para começar!")
