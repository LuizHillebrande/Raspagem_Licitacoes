import pandas as pd

# Gerar 2000 linhas com as informações desejadas
dados = {
    'Razao Social': ['teste ltda'] * 2000,  # Coluna com 'teste ltda' em todas as linhas
    'Email': ['luiz.hillebrande1505@gmail.com'] * 2000  # Coluna com o email em todas as linhas
}

# Criar o DataFrame
df = pd.DataFrame(dados)

# Salvar o DataFrame em um arquivo Excel
df.to_excel('empresa_emails.xlsx', index=False)

print("Arquivo 'empresa_emails.xlsx' gerado com sucesso!")
