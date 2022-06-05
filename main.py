from random import Random
from openpyxl import Workbook, load_workbook

def criarTabuleiro(qtdLinha, qtdColuna ):
    random = Random()
    wb = Workbook()
    bombas = []
    abaJogo = wb.create_sheet('jogo1')
    totalCasas = qtdColuna*qtdLinha
    dificuldade = 2 # 1 = facil | 2 = médio | 3 = dificil

    #NIVEL DE DIFICULDADE
    if dificuldade == 1:
        qtdBombas = int(totalCasas*0.15) #int => para 'truncar o numero'
    elif dificuldade == 2:
        qtdBombas = int(totalCasas*0.25)
    else:
        qtdBombas = int(totalCasas*0.50)
    vazio = 0
    cont = 1
    #ALETORIEDADE DAS BOMBAS
    for i in range(qtdBombas):
        lcBombas = random.randrange(1, totalCasas)
        if lcBombas in bombas:
            lcBombas = random.randrange(1, totalCasas)
        bombas.append(lcBombas)
    #print(bombas, len(bombas), qtdBombas)
    #CRIAÇÃO DO TABULEIRO
    for i in range(1, qtdLinha+1):
        for j in range(1, qtdColuna+1):
            if cont in bombas:
                abaJogo.cell(row=i, column=j, value=-1)
            elif cont not in bombas:
                abaJogo.cell(row=i, column=j, value=vazio)
            cont += 1
    wb.save('banco.xlsx')

wb = load_workbook(filename='banco.xlsx', read_only=True)
jogo = wb['jogo1']
print(jogo.cell(column=2, row=1).value)
