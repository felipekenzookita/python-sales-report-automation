# 📊 Sales Report Automation

Projeto desenvolvido em Python para automatizar a geração de relatórios de vendas a partir de uma planilha Excel. O sistema realiza análises dos dados, calcula indicadores de desempenho e envia automaticamente um relatório formatado por e-mail utilizando o Microsoft Outlook.

---

## 🚀 Funcionalidades

- Leitura de dados de uma planilha Excel
- Cálculo do faturamento por loja
- Cálculo da quantidade de produtos vendidos
- Cálculo do ticket médio por loja
- Geração de tabelas em HTML
- Envio automático de relatório por e-mail via Outlook

---

## 🛠️ Tecnologias Utilizadas

- Python
- Pandas
- OpenPyXL
- PyWin32

---

## 📂 Estrutura do Projeto

```
python-sales-report-automation/
│
├── main.py
├── Vendas.xlsx
├── README.md
└── requirements.txt
```

---

## ▶️ Como Executar

1. Clone este repositório

```bash
git clone https://github.com/felipekenzookita/python-sales-report-automation.git
```

2. Instale as dependências

```bash
pip install -r requirements.txt
```

3. Execute o programa

```bash
python main.py
```

> **Observação:** É necessário ter o Microsoft Outlook instalado e configurado no computador para o envio automático dos e-mails.

---

## 📈 Indicadores Gerados

O relatório apresenta:

- Faturamento por loja
- Quantidade total de produtos vendidos
- Ticket médio por loja

---

## 📧 Resultado

Ao final da execução, o sistema envia automaticamente um e-mail contendo os indicadores em formato de tabela HTML.

---

## 📚 Bibliotecas Utilizadas

```python
import pandas as pd
import win32com.client
```

---

## 👨‍💻 Autor

Felipe Kenzo
