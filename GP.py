import pandas as pd
import openpyxl
from tkinter import filedialog as tk

bandeira = True

# Faz a verificação dos casos pendentes em inconsistências de GP.
def incon_fase_0():
    num = 0
    for i in usuario_anti:
        analise_x_anti = usuario_anti[num]
        users_anti.append(analise_x_anti)

        num += 1

def incon_fase_1():
    num = 0
    for i in usuario:
        analise_x = usuario.loc[num]
        analise_n = maquina.loc[num]
        analise_t = tagse.loc[num]
        
        
        if isinstance(analise_t, float) :
        
            if analise_x not in users:

                if analise_x in users_anti:
                    users.append(analise_x)
                    final.append([analise_x, 'Pendente Gestão', 'Pendencia Antiga'])
                else:
                    users.append(analise_x)
                    final.append([analise_x, 'Pendente Gestão', 'Nova pendencia'])
        
        num += 1

# Faz a verificação dos casos resolvidos em inconsistências de GP.
def incon_fase_2():
    num = 0
    for i in usuario:
        analise_x = usuario.loc[num]

        if analise_x not in users and analise_x not in fusers:
        
            fusers.append(analise_x)
            final.append([analise_x, 'Resolvido'])
        
        num += 1
    
    final.sort()

# Função Mãe, unifica todas as funções e atualiza a planilha de resposta, das inconsistencia de GP.
def inconsistencias():

    incon_fase_0()
    incon_fase_1()
    incon_fase_2()
    df = pd.DataFrame(final, columns=['Users', 'status', 'Época'])
    df.to_excel(tabela_res, 'Planilha2', index=None)

while bandeira == True:

    print ('\n\nSejá Bem vindo\n\n')
    
    print ("Selecione a planilha que deseja analisar.")
    tabela = tk.askopenfilename()
    print(tabela)
    print ("Selecione a planilha da ultima semana.")
    tabela_anti = tk.askopenfilename()

    

    tabela_res = "Resultado_GP.xlsx"
    lUsermaisPosto = pd.read_excel(tabela, "Extracao_+umposto", header=0, index_col=None)
    lUsermaisPosto_anti = pd.read_excel(tabela_anti, "Extracao_+umposto", header=0, index_col=None)
    usuario = lUsermaisPosto.get('Atribuído a')
    maquina = lUsermaisPosto.get('Nome')
    usuario_anti = lUsermaisPosto_anti.get('Atribuído a')
    tagse = lUsermaisPosto.get('Tag Secundária')

    users_anti = []
    fusers = []
    users = []
    final = []


    print(f'Analisando planilha {tabela}')
    inconsistencias()
    
    resposta = input("\n\nPrcesso finalizado!\nDeseja sair? (s) ou (n):")
    resposta = resposta.lower()
    
    print(resposta)

    
    while resposta != 's' and resposta != 'n':
        resposta = input('\nEscolha não encontrada.\nPor favor tente novamente: (s) ou (n) ')
        resposta = resposta.lower()

    if resposta == 's':
        bandeira = False
        
    else:
        bandeira = True
