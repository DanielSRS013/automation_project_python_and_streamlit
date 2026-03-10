"""def send_email(df_excel):
    email = EmailMessage()
    email['Subject'] = 'FORMULÁRIO_ESTOQUE NEGATIVO'
    email['From'] = 'aprendiz.auditoria@lennyniemeyer.com'
    email['To'] = 'nadaavercomg@gmail.com'

    email.set_content(
Olá,

Segue em anexo o formulário de estoque em formato Excel.

Atenciosamente,
Sistema Automático
)
    
    caminho = Path(df_excel)
    tipo, enconding = mimetypes.guess_type(caminho)

    with open(caminho, 'rb') as f:
        email.add_attachment(
            f.read(),
            maintype = 'application',
            subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename = caminho.name
        )

    with smtplib.SMTP('smtp.office365.com', 587) as smtp:
        smtp.starttls()
        smtp.login()
        smtp.send_message(email)"""


import pandas as pd

# 1. Criar dados de exemplo
data = {'Produto': ['A', 'B', 'C'], 'Vendas': [100, 150, 50]}
df = pd.DataFrame(data)

# 2. Função de estilo (colorir de vermelho se Vendas < 60)
def destacar_vendas(val):
    color = 'red' if val < 60 else 'white'
    return f'background-color: {color}'

# 3. Aplicar estilo e exportar para Excel
styled_df = df.style.applymap(destacar_vendas, subset=['Vendas'])
styled_df.to_excel('relatorio_colorido.xlsx', index=False)



import pandas as pd

# Criando um DataFrame de exemplo
df = pd.DataFrame({
    'Produto': ['A', 'B', 'C', 'D'],
    'Vendas': [30, 75, 45, 90]
})

# Função para aplicar cor
def colorir(valor):
    if valor > 50:
        return 'background-color: lightgreen'
    else:
        return ''

# Aplicando estilo
df_style = df.style.applymap(colorir, subset=['Vendas'])

# Exportando para Excel
df_style.to_excel('relatorio.xlsx', engine='openpyxl', index=False)