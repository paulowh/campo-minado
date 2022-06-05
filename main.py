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

def lerJogo():
    wb = load_workbook(filename='banco.xlsx', read_only=True)
    jogo = wb['jogo1']
    linha = jogo.max_row
    coluna = jogo.max_column

    for i in range(1, linha+1):
        for j in range(1, coluna + 1):
            print('{:^3}'.format(jogo.cell(column=j, row=i).value), end='|')
            #print('{:^5}'.format('[ ]'), end='')
        print('\n')

def verificarJogada():
    wb = load_workbook(filename='banco.xlsx', read_only=True)
    jogo = wb['jogo1']
    maxLinha = jogo.max_row
    maxColuna = jogo.max_column
    gameOver = False
    jogadas = []

    while True:
        n1 = int(input('linha: '))
        n2 = int(input('coluna: '))
        #teste de jogadas
        jogadas.append(jogo.cell(column=n2, row=n1))

        for i in range(1, maxLinha+1):
            for j in range(1, maxColuna + 1):
                passouK = False
                for k in jogadas:  
                    #print(k, jogo.cell(column=j, row=i))
                    if k == jogo.cell(column=j, row=i) and jogo.cell(column=j, row=i).value == 0:
                        print('{:^5}'.format('[0]'), end='')
                        passouK = True
                    elif k == jogo.cell(column=j, row=i) and jogo.cell(column=j, row=i).value == -1:
                        print('{:^5}'.format('[x]'), end='')
                        passouK = True
                        gameOver = True
                if passouK == False:
                    print('{:^5}'.format('[ ]'), end='')
            print('\n')

        if gameOver == True:
            print('Game Over!!!')
            break
    
verificarJogada()


    



