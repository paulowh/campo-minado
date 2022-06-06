from random import Random
from openpyxl import Workbook, load_workbook

def bombinhas(totalCasas):
    random = Random()
    dificuldade = 3 # 1 = facil | 2 = médio | 3 = dificil
    lcBombas = 0
    bombas = []
    dificuldade, linhas, colunas = configuracoes()

    #NIVEL DE DIFICULDADE
    if dificuldade == 1:
        qtdBombas = int(totalCasas*0.15) #int => para 'truncar o numero'
    elif dificuldade == 2:
        qtdBombas = int(totalCasas*0.25)
    else:
        qtdBombas = int(totalCasas*0.50)

    #ALETORIEDADE DAS BOMBAS
    for i in range(qtdBombas):
        lcBombas = random.randrange(1, totalCasas)
        if lcBombas in bombas:
            lcBombas = random.randrange(1, totalCasas)
        bombas.append(lcBombas)

    return bombas

def criarTabuleiro(qtdLinha, qtdColuna):
    try:
        wb = load_workbook(filename='banco.xlsx')
    except:
        wb = Workbook()
        wb.active.title = 'jogo'
    
    try:
        abaJogo = wb['jogo']
    except:
        abaJogo = wb.create_sheet('jogo')

    vazio = 0
    cont = 1

    bombas = bombinhas(qtdColuna*qtdLinha)

    #CRIAÇÃO DO TABULEIRO
    for i in range(1, qtdLinha+1):
        for j in range(1, qtdColuna+1):

            if cont in bombas:
                abaJogo.cell(row=i, column=j, value=-1)
            elif cont not in bombas:
                abaJogo.cell(row=i, column=j, value=vazio)
            cont += 1
    wb.save('banco.xlsx')

def verificarJogada():
    wb = load_workbook(filename='banco.xlsx', read_only=False)
    jogo = wb['jogo']
    maxLinha = jogo.max_row
    maxColuna = jogo.max_column
    gameOver = False
    ganhou = False
    jogadas = []
    dificuldade, qtdLinha, qtdColuna = configuracoes()
    while True:
        n1 = int(input('linha: '))
        n2 = int(input('coluna: '))
        #teste de jogadas
        
        if jogo.cell(column=n2, row=n1) not in jogadas:
            jogadas.append(jogo.cell(column=n2, row=n1))
            for i in range(1, maxLinha+1):
                for j in range(1, maxColuna + 1):
                    passouK = False
                    for k in jogadas:  
                        #print(k, jogo.cell(column=j, row=i))
                        if k == jogo.cell(column=j, row=i) and jogo.cell(column=j, row=i).value == 0:
                            print('{:^5}'.format(jogo.cell(column=j, row=i).value), end='')
                            passouK = True
                            if len(jogadas) == (qtdColuna*qtdLinha - len(bombinhas(qtdColuna*qtdLinha))):
                                ganhou = True
                                
                        elif k == jogo.cell(column=j, row=i) and jogo.cell(column=j, row=i).value == -1:
                            print('{:^5}'.format('X'), end='')
                            passouK = True
                            gameOver = True
                    if passouK == False:
                        print('{:^5}'.format('[ ]'), end='')
                print('\n')
        else:
            print('Jogada já efetuada, tente novamente')

        if gameOver == True:
            print('Game Over!!!')
            break
        if ganhou == True:
            print('Você Ganhou')
            break

def configuracoes():
    try:
        wb = load_workbook(filename='banco.xlsx')
        config = wb['configurações']
    except:
        wb = Workbook()
        config = wb.create_sheet('configurações')
        config.cell(column=1, row=1, value='Linhas')
        config.cell(column=2, row=1, value=3)
        config.cell(column=1, row=2, value='Colunas')
        config.cell(column=2, row=2, value=3)
        config.cell(column=1, row=3, value='Dificuldade')
        config.cell(column=2, row=3, value=2)
    
    dificuldade = config.cell(column=2, row=3).value
    linhas = config.cell(column=2, row=1).value
    colunas = config.cell(column=2, row=2).value
    wb.save('banco.xlsx')
    
    return int(dificuldade), int(linhas), int(colunas)

def alterarConfiguracao():
    try:
        wb = load_workbook(filename='banco.xlsx')
        config = wb['configurações']
    except:
        wb = Workbook()
        config = wb.create_sheet('configurações')
    dificuldade = input('Qual o nivel de dificuldade(1/2/3): ')
    linhas = input('Quantidade de linhas: ')
    colunas = input('Quantidade de colunas: ')

    config.cell(column=1, row=1, value='Linhas')
    config.cell(column=2, row=1, value=linhas)
    config.cell(column=1, row=2, value='Colunas')
    config.cell(column=2, row=2, value=colunas)
    config.cell(column=1, row=3, value='Dificuldade')
    config.cell(column=2, row=3, value=dificuldade)
    wb.save('banco.xlsx')

#alterarConfiguracao()

dificuldade, linhas, colunas = configuracoes()
criarTabuleiro(linhas, colunas)
verificarJogada()
