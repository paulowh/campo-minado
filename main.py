import random
from openpyxl import Workbook, load_workbook

def configuracoes():
    try:
        wb = load_workbook(filename='banco.xlsx')
        config = wb['configurações']
    except: #CASO NÃO TENHA O ARQUIVO ELE CRIA UMA CONFIGURAÇÃO BASICA
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

def bombinhas():
    bombas = []
    dificuldade, linhas, colunas = configuracoes()
    totalCasas = (linhas*colunas)
    # NIVEL DE DIFICULDADE
    if dificuldade == 1:
        qtdBombas = int(totalCasas*0.15)  # int => para 'truncar' o numero
    elif dificuldade == 2:
        qtdBombas = int(totalCasas*0.25)
    else:
        qtdBombas = int(totalCasas*0.50)

    # ALETORIEDADE DAS BOMBAS
    while True:
        num = random.randint(1, totalCasas)
        if num not in bombas:
            bombas.append(num)
        if len(bombas) == qtdBombas: 
            break

    return bombas

def calcularAdjacente(matriz):
    dificuldade, qtdLinha, qtdColuna = configuracoes()

    for a in range(0, qtdLinha):
        for b in range(0, qtdColuna):
            if matriz[a][b] == '-1': 
                continue
            cont_minas_adj = 0
                
            for c in range(a-1 if a>0 else 0, a+2 if a<(qtdLinha-1) else qtdLinha):
                for d in range(b-1 if b>0 else 0, b+2 if b<(qtdColuna-1) else qtdColuna):
                    #print(c, d)
                    if matriz[c][d] == '-1': 
                        cont_minas_adj += 1

            matriz[a][b] = str(cont_minas_adj)
        #print(matriz)
    return matriz

def gravarTabuleiro(jogo):
    try:
        wb = load_workbook(filename='banco.xlsx', read_only=False)
    except:
        wb = Workbook()

    try:
        abaJogo = wb['jogo']
    
    except:
        abaJogo = wb.create_sheet('jogo')
    #resolvendo o problema de quando muda as config para um tabuleiro menor
    abaJogo.delete_rows(1, abaJogo.max_row)

     # GRAVANDO TABULEIRO
    for i in range(1, len(jogo)+1):
        for j in range(1, len(jogo[0])+1):
            abaJogo.cell(row=i, column=j, value=jogo[i-1][j-1])

    wb.save('banco.xlsx')

def criarTabuleiro(qtdLinha, qtdColuna):
   
    cont = 1
    bombas = bombinhas()
    game = []

    #criando primeiro uma matriz para facilitar minha vida
    for i in range(1, qtdLinha+1):
        jogo = []
        for j in range(1, qtdColuna+1):
            if cont in bombas:
                jogo.append('-1')
            elif cont not in bombas:
                jogo.append('0')
            cont += 1
        game.append(jogo)
    #print(game)

    gravarTabuleiro(calcularAdjacente(game))

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
        # teste de jogadas

        if jogo.cell(column=n2, row=n1) not in jogadas:
            jogadas.append(jogo.cell(column=n2, row=n1))
            for i in range(1, maxLinha+1):
                for j in range(1, maxColuna + 1):
                    passouK = False
                    for k in jogadas:
                        # print(k, jogo.cell(column=j, row=i))
                        if k == jogo.cell(column=j, row=i) and jogo.cell(column=j, row=i).value != '-1':
                            print('{:^5}'.format(
                                jogo.cell(column=j, row=i).value), end='')
                            passouK = True
                            if len(jogadas) == (qtdColuna*qtdLinha - len(bombinhas())):
                                ganhou = True

                        elif k == jogo.cell(column=j, row=i) and jogo.cell(column=j, row=i).value == '-1':
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

alterarConfiguracao()

dificuldade, linhas, colunas = configuracoes()
#criarTabuleiro(linhas, colunas)
criarTabuleiro(linhas,colunas)
#verificarJogada()
