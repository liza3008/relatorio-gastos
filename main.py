
import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
from openpyxl import Workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ===============================
# CONFIGURAÇÕES
# ===============================
PASTA_RELATORIOS = "relatorios"
os.makedirs(PASTA_RELATORIOS, exist_ok=True)

# ===============================
# GERAÇÃO DOS DADOS
# ===============================
def gerar_dados():
    dados = {
        "Categoria": ["Alimentação", "Transporte", "Aluguel", "Internet", "Saúde", "Lazer", "Educação"],
        "Valor (R$)": [820.50, 300.00, 1200.00, 120.00, 450.00, 200.00, 380.00]
    }
    df = pd.DataFrame(dados)
    df["Percentual"] = (df["Valor (R$)"] / df["Valor (R$)"].sum() * 100).round(2)
    return df

# ===============================
# GERAÇÃO DE GRÁFICO
# ===============================
def gerar_grafico(df):
    plt.figure(figsize=(7,5))
    plt.bar(df["Categoria"], df["Valor (R$)"], color="#0078D7", alpha=0.85)
    plt.title("Gastos por Categoria", fontsize=14, fontweight="bold")
    plt.xlabel("Categoria")
    plt.ylabel("Valor (R$)")
    plt.grid(axis="y", linestyle="--", alpha=0.4)
    plt.tight_layout()

    grafico_path = os.path.join(PASTA_RELATORIOS, "grafico_gastos.png")
    plt.savefig(grafico_path)
    plt.close()
    return grafico_path

# ===============================
# GERAÇÃO DO PDF
# ===============================
def gerar_pdf(df, grafico):
    pdf_path = os.path.join(PASTA_RELATORIOS, "relatorio_gastos.pdf")
    c = canvas.Canvas(pdf_path, pagesize=A4)

    c.setFont("Helvetica-Bold", 18)
    c.drawString(160, 800, "Relatório de Gastos Mensais")

    y = 760
    total = 0
    c.setFont("Helvetica", 12)
    for i, row in df.iterrows():
        c.drawString(80, y, f"{row['Categoria']}: R$ {row['Valor (R$)']:.2f}")
        y -= 20
        total += row["Valor (R$)"]

    c.setFont("Helvetica-Bold", 13)
    c.drawString(80, y - 10, f"Total Geral: R$ {total:.2f}")

    c.drawImage(ImageReader(grafico), 100, 100, width=380, height=250)
    c.save()
    return pdf_path

# ===============================
# EXPORTAÇÃO DE CSV E EXCEL
# ===============================
def exportar_planilhas(df):
    csv_path = os.path.join(PASTA_RELATORIOS, "relatorio_gastos.csv")
    excel_path = os.path.join(PASTA_RELATORIOS, "relatorio_gastos.xlsx")
    df.to_csv(csv_path, index=False)
    df.to_excel(excel_path, index=False)
    return csv_path, excel_path

# ===============================
# ENVIO DE E-MAIL ESTILIZADO
# ===============================
def enviar_email(anexos):
    remetente = "cristineelizabeth06@gmail.com"
    senha = "xase puak duan ibwr"
    destinatarios = ["eliza300821@gmail.com"]

    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    data = date.today().strftime("%d/%m/%Y")
    msg["Subject"] = f"Relatório de Gastos Mensais — {5/11/2025}"

    corpo_html = f"""
    <html>
      <body style="font-family: 'Segoe UI', Arial; background: linear-gradient(135deg, #0078D7, #00B4D8); margin:0; padding:30px;">
        <div style="background:#fff; border-radius:15px; max-width:650px; margin:auto; padding:30px; box-shadow:0 4px 15px rgba(0,0,0,0.15);">
          <div style="text-align:center; border-bottom:3px solid #0078D7; padding-bottom:15px;">
            <img src="https://cdn-icons-png.flaticon.com/512/3135/3135692.png" width="70" />
            <h2 style="color:#0078D7; margin:10px 0;">Relatório de Gastos — {5/11/2025}</h2>
            <p style="color:#555;">Resumo financeiro automático gerado pelo sistema Python</p>
          </div>

          <div style="margin-top:25px; color:#333;">
            <p>Olá, <strong>Equipe Financeira</strong> </p>
            <p>Segue abaixo um resumo dos seus <b>gastos mensais</b>:</p>

            <table width="100%" style="border-collapse:collapse; margin-top:15px;">
              <tr style="background-color:#0078D7; color:white;">
                <th style="padding:10px;">Categoria</th>
                <th style="padding:10px;">Valor (R$)</th>
                <th style="padding:10px;">%</th>
              </tr>
    """

    # Inserindo tabela com dados
    df = gerar_dados()
    for _, row in df.iterrows():
        corpo_html += f"""
              <tr style="text-align:center; background-color:#f9f9f9;">
                <td style="padding:10px;">{row['Categoria']}</td>
                <td style="padding:10px;">R$ {row['Valor (R$)']:.2f}</td>
                <td style="padding:10px;">{row['Percentual']}%</td>
              </tr>
        """

    total = df["Valor (R$)"].sum()
    corpo_html += f"""
              <tr style="background-color:#e3f2fd; font-weight:bold; text-align:center;">
                <td style="padding:10px;">Total</td>
                <td style="padding:10px;" colspan="2">R$ {total:.2f}</td>
              </tr>
            </table>

            <div style="text-align:center; margin-top:30px;">
              <a href="https://drive.google.com/drive/folders/SEU_LINK_AQUI" target="_blank" 
                 style="background-color:#28a745; color:white; padding:14px 28px; border-radius:8px; 
                        text-decoration:none; font-weight:bold; font-size:15px; transition:0.3s;">
                 Baixar Relatório Completo
              </a>
            </div>

            <p style="margin-top:25px; color:#555;">Você também pode conferir o gráfico completo no PDF anexo.</p>

            <div style="text-align:center; margin-top:40px;">
              <img src="https://cdn-icons-png.flaticon.com/512/1598/1598163.png" width="60" />
              <p style="color:#999; font-size:13px;">Relatório gerado automaticamente por <b>Python</b></p>
            </div>
          </div>
        </div>

        <p style="text-align:center; margin-top:20px; color:white; font-size:12px;">
          © 2025 Sistema de Relatórios Automáticos — Desenvolvido em Python
        </p>
      </body>
    </html>
    """

    msg.attach(MIMEText(corpo_html, "html"))

    # Anexa os arquivos
    for arquivo in anexos:
        with open(arquivo, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(arquivo))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo)}"'
            msg.attach(part)

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(remetente, senha)
            server.send_message(msg)
        print("📨 E-mail estilizado enviado com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")

# ===============================
# EXECUÇÃO PRINCIPAL
# ===============================
if __name__ == "__main__":
    df = gerar_dados()
    grafico = gerar_grafico(df)
    pdf = gerar_pdf(df, grafico)
    csv, excel = exportar_planilhas(df)

    anexos = [pdf, csv, excel]
    enviar_email(anexos)
