# 📊 Relatório de Faturamento Diário - E-commerce

Sistema interativo criado com **Streamlit** para geração de relatórios de faturamento a partir de planilhas Excel, com visualização de métricas, gráficos e envio automático de e-mails em PDF.

## ✅ Funcionalidades

- 📅 Filtros por período, cliente e tipo de cliente
- 📈 Gráficos dinâmicos: Top N Clientes e Pizza de Faturamento
- 💰 Indicador de faturamento total
- 📄 Geração de PDF com tabela e total consolidado
- 📥 Download do relatório em Excel
- 📧 Envio de relatórios automáticos por e-mail (Outlook)
- 🧑‍💼 Suporte a múltiplos destinatários e mensagem personalizada
------------------------------------------------------------------------------------------------
## 🚀 Como usar

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/relatorio-ecommerce-v3.git
cd relatorio-ecommerce-v3

pip install -r requirements.txt
streamlit run app.py
------------------------------------------------------------------------------------------------
📁 Estrutura esperada da planilha Excel
A aba "Relatório Base" deve conter colunas como:

Data Entrada

Nome Cliente

Tipo Cliente

Venda Líquida

O cabeçalho deve começar na segunda linha da planilha (linha 2 no Excel).

👨‍💻 Desenvolvido por
Murilo Morini Marangoni
[murilo.3marangoni@gmail.com]
