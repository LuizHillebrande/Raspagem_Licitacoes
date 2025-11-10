import pandas as pd
import os
import sys

nome_arquivo = "emails_vencedores.xlsx"

# Se passar --delete, apenas exclui o arquivo
if len(sys.argv) > 1 and sys.argv[1] == "--delete":
    if os.path.exists(nome_arquivo):
        os.remove(nome_arquivo)
        print(f"✅ Arquivo '{nome_arquivo}' excluído com sucesso!")
    else:
        print(f"❌ Arquivo '{nome_arquivo}' não encontrado.")
    sys.exit(0)

# Criar dados de teste com 6000 linhas
dados = []
for i in range(1, 6001):
    dados.append({
        "Razao Social": f"Empresa Teste {i}",
        "Email": f"teste{i}@exemplo.com",
        "CNPJ": f"{i:014d}"
    })

# Criar DataFrame
df = pd.DataFrame(dados)

# Salvar no arquivo
df.to_excel(nome_arquivo, index=False)

print(f"✅ Arquivo '{nome_arquivo}' criado com {len(df)} linhas!")
print(f"📍 Localização: {os.path.abspath(nome_arquivo)}")
print("\n🎯 Agora você pode testar a função de divisão de arquivos na interface.")
print("\n💡 Para excluir o arquivo, execute: python teste_gerar_arquivo.py --delete")

