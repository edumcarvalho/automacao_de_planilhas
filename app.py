import openpyxl

workbook = openpyxl.Workbook()
del workbook['Sheet']

def add_pages():
    mais_pagina = 's'
    while mais_pagina == 's':                
        paginas.append(input('Digite o nome da página:'))         
        mais_pagina = input('Criar mais uma página nesta planilha?(s/n)')
    for pagina in paginas:
        workbook.create_sheet(pagina)

def add_line(line):
    mais_coluna = 's'        
    while mais_coluna == 's':                        
        if line == 'header':
            linha.append(input('Digite uma coluna para o cabeçaho:'))         
            mais_coluna = input('Adicionar mais uma coluna?(s/n)')
        if line == 'row':                                                
            register = input('Digite os dados a serem adicionados a uma nova linha, separados por vírgula:')
            lin = register.split(',')                     
            sheet_atual.append(lin)    
            mais_coluna = input('Adicionar mais uma linha?(s/n)')
    if line == 'header':
        sheet_atual.append(linha)    
while True:
    paginas = []        
    linha = []
    print('Bem-vindo ao gerador de planilhas!')
    print('Para começar vamos criar uma nova página dentro de uma planilha')    

    add_pages()
    print(paginas)
    sheet_atual = workbook[input('Digite o nome da página a ser manipulada:')]

    add_line('header')
    print('Colunas:')
    print(linha)

    if input('Adicionar dados a essa planilha?(s/n)') != 's':
        break

    print(f'As páginas disponíveis no momentos são: {paginas}')    
    add_line('row')

    arquive = input('Digite o nome da planilha a ser salva:')
    workbook.save(f'{arquive}.xlsx')
    print('Planilha criada com sucesso.')
    break

    

