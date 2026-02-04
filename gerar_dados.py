import pandas as pd
from datetime import datetime, timedelta
import random

# Configurações
ARQUIVO_OUTPUT = "input_teste.xlsx"
QTD_LINHAS = 20

# Dados Fictícios
descricoes = [
    "Erro ao processar pagamento", "Acesso negado ao sistema", "Solicitação de novo perfil",
    "Relatório não carrega", "Ajuste de permissão", "Integração falhou", 
    "Dúvida sobre funcionalidade", "Tela preta ao iniciar", "Dados inconsistentes",
    "Atualização de cadastro"
]

situacoes = [
    "Programando", "Atendimento pendente (suporte)", "Homologando", 
    "Verificando", "Aguardando testes internos", "Ag. Confirmação de Orçamento",
    "Aguardando liberacao oficial", "Retorno de homologacao", "Aguardando detalhamento",
    "Resolvido", "Liberado para cliente"
]

classificacoes = ["CORRECAO", "MELHORIA", "DUVIDA", "PROJETO"]

usuarios = [
    "usuario.jsilva", "usuario.moliveira", "usuario.psantos", "usuario.ti"
]

dados = []

for i in range(1, QTD_LINHAS + 1):
    # Gera datas aleatórias (algumas atrasadas, algumas no prazo)
    dias_offset = random.randint(-10, 5) # Entre 10 dias atrás e 5 dias no futuro
    data_sla = datetime.now() + timedelta(days=dias_offset)
    
    linha = {
        "Protocolo": f"REQ-{2024000 + i}",
        "Resumo": random.choice(descricoes),
        "Situação": random.choice(situacoes),
        "Classificação": random.choice(classificacoes),
        "Prazo SLA": data_sla.strftime("%d/%m/%Y %H:%M"),
        "Usuário responsável": random.choice(usuarios) if random.random() > 0.2 else None, # 20% vazio
        "Incluído por": random.choice(usuarios)
    }
    dados.append(linha)

df = pd.DataFrame(dados)
df.to_excel(ARQUIVO_OUTPUT, index=False)

print(f"Arquivo '{ARQUIVO_OUTPUT}' gerado com sucesso com {QTD_LINHAS} linhas!")