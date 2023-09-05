Passos em que o script se baseia:


Arquivo CNPJ: Limpar coluna A da planilha CNPJ (CTRL + SHIFT + L)
Arquivo AJUSTE_BASE: Limpar a planilha TODA (CTRL + SHIFT + T)
Matriz Estudo Mensal: Limpar a coluna A da aba BASE.PF (CTRL + SHIFT + L)
Matriz Estudo Mensal: Limpar a coluna A da aba BASE.Empresas (CTRL + SHIFT + L)
Matriz Estudo Mensal: Limpar a coluna A e B da aba BASE.Socios (CTRL + SHIFT + L)

- TR_USER 1
- Clicar no campo TIPO DE CONSULTA
- Aba DADOS - FILTRO
- Filtrar na COLUNA D apenas PJ
- Copiar o CNPJ de todos os registros da coluna A e colar como valores na planilha CNPJ

- TR_USER 2
- Clicar no campo TIPO DE CONSULTA
- Aba DADOS - FILTRO
- Filtrar na COLUNA D apenas PJ
- Copiar todos os registros da coluna A e colar na planilha CNPJ

- TR_USER 3
- Clicar no campo TIPO DE CONSULTA
- Aba DADOS - FILTRO
- Filtrar na COLUNA D apenas PJ
- Copiar todos os registros da coluna A e colar na planilha CNPJ

- TR_USER 4
- Clicar no campo TIPO DE CONSULTA
- Aba DADOS - FILTRO
- Filtrar na COLUNA D apenas PJ
- Copiar todos os registros da coluna A e colar na planilha CNPJ

- Na planilha CNPJ remover duplicatas
- SALVAR

- Copiar os CNPJs do Arquivo CNPJ (base ajustada) para o arquivo "Matriz Estudo Mensal" na planilha BASE.Empresas

====================

- No Datahub, selecionar LOCALIZE LOTE, Clicar em PJ, carregar o arquivo CNPJ e selecionar QSA. Dar nome à extração e selecionar formato excel. Iniciar processamento.

- No Datahub, na área de administração, em DOWNLOADS, fazer o download do arquivo solicitado

====================

- Selecionar todas as células da planilha CNPJ_QSA e colar na PLANILHA AJUSTE_BASE

- Rodar a MACRO: AJUSTAR_BASE (CTRL + SHIFT + A), que possui o seguinte código:

    Sub AJUSTAR_BASE()
        Columns("B:D").Select
        Selection.Delete Shift:=xlToLeft
        Columns("C:E").Select
        Selection.Delete Shift:=xlToLeft
        Columns("D:G").Select
        Selection.Delete Shift:=xlToLeft
        Columns("E:G").Select
        Selection.Delete Shift:=xlToLeft
        ActiveWindow.SmallScroll ToRight:=2
        ActiveWindow.SmallScroll Down:=0
        Columns("F:I").Select
        Selection.Delete Shift:=xlToLeft
        Columns("G:I").Select
        Selection.Delete Shift:=xlToLeft
        Columns("H:J").Select
        Selection.Delete Shift:=xlToLeft
        ActiveWindow.SmallScroll ToRight:=-2
    End Sub

- Clicar no Menu DADOS > Abrir o PowerQuery (Obter Dados)

- Selecionar arquivo excel e a planilha AJUSTE_BASE, Próximo

- Selecionar a Pasta de Trabalho: AJUSTE_BASE

- Clicar em TRANSFORMAR DADOS

- Selecionar a COLUNA CNPJ com o botão direito > "Transformar Outras Colunas em Linhas"

- Na Coluna "Valor", excluir os "em branco"

- Clicar em "FECHAR E CARREGAR"

====================

- Copiar os CPFs da Base Ajustada (Arquivo > AJUSTE_BASE - COLUNA C) para o arquivo "2023 - Matriz Estudo Mensal" na planilha BASE.SOCIOS


- TR_USER 1
- Clicar no campo TIPO DE CONSULTA
- Aba DADOS - FILTRO
- Filtrar na COLUNA D apenas PF
- Copiar o CPF de todos os registros da coluna A e colar como valores na planilha "2023 - Matriz Estudo Mensal" na planilha BASE.PF

- TR_USER 2
- Clicar no campo TIPO DE CONSULTA
- Aba DADOS - FILTRO
- Filtrar na COLUNA D apenas PJ
- Copiar o CPF de todos os registros da coluna A e colar como valores na planilha "2023 - Matriz Estudo Mensal" na planilha BASE.PF

Substituir "." e "-" por VAZIO (CTRL + SHIFT + D)

- Na planilha BASE.PF remover duplicatas do CPF


====================