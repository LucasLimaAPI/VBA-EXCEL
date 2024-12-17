Attribute VB_Name = "Aula3"
Option Explicit
Sub fnComparaPastaDeTrabalhoExcel()

    'Declara��o das vari�veis de objeto
    Dim wb As Workbook      'Pasta de trabalho
    Dim ws As Worksheet     'Planilha
    
    'Ler a pasta indicada na planilha "Origem", selecionando apenas .xlsm
    Call fnLePasta(Workbooks("VBA 4.xlsm").Sheets("VBA4").Range("B1"))
    
    'Abrir um formul�rio que me permita selecionar as pastas
    'Deve ser possivel indicar uma pasta de trabalho origem ea planilha a ser lida
    'Deve ser possivel indicar uma pasta de trabalho destino e uma pasta origem, uma pasta destino e uma planilha em cada uma delas
    'As diferen�as devem ser apontadas na planilha "Diferen�as" que esta na pasta VBA 4.xlsm
    
    'Abertura da pasta de trabalho escolhida
    If fnAbrePastaDeTrabalho("C:\Users\lucas.oliveira\Desktop\VBA\Materiais\Planilha Modelo.xlsx", wb) Then
        Set ws = wb.ActiveSheet

        Call fnFechaPastaDeTrabalho(wb)
        MsgBox "Pasta de trabalho manipulada com sucesso"
    Else
        MsgBox "Pasta de trabalho n�o pode ser aberta"
    End If
    
   'Limpeza de objetos
    Set wb = Nothing
    Set ws = Nothing

End Sub

'Fun��o que abre a pasta de trabalho
Private Function fnAbrePastaDeTrabalho(pCaminhoCompletoWB As String, pWB As Object) As Boolean
    
    fnAbrePastaDeTrabalho = False
    
    '"Seta" As var�veis na "pasta de trabalho" e na  "Planilha de Trabalho"
    Set pWB = Workbooks.Open(pCaminhoCompletoWB)
    
    'Caso n�o de nenhum erro, responde verdadeiro.
    fnAbrePastaDeTrabalho = True
End Function
    
Private Function fnFechaPastaDeTrabalho(pWB As Object) As Boolean
    'Fecha a pasta de trabalho sem salvar
    pWB.Close False

    'Caso n�o d� nenhum erro, responde verdadeiro
    fnFechaPastaDeTrabalho = True

End Function

Private Function fnComparaPlanilhas() As Boolean

    'Declara��o de vari�veis
    Dim rCelula                 As Range
    Dim iDiferencas             As Integer

    
    'Inicializa a vari�vel de diferen�as
    iDiferencas = 0
     
    'Verifica se as demais planilha destino � igual a planilha origem
    For Each rCelula In Selection
        If rCelula.Value <> Sheets(sNomePlanComparar).Range(rCelula.Address) Then
            'Muda a cor da fonte e do interior da c�lula se ela for diferente da origem
            rCelula.Interior.Color = vbRed
            rCelula.Font.Color = vbYellow
            iDiferencas = iDiferencas + 1
        Else
            'Garante que a c�lula esteja sem preenchimento e com a fonte "automatic"
            rCelula.Interior.Pattern = xlNone
            rCelula.Font.ColorIndex = xlAutomatic
        End If
    
    Next

    'If que define qual mensagem vai ser mostrada
    If iDiferencas = 0 Then
        MsgBox "Nenhuma C�lula Modificada"
    Else
        MsgBox iDiferencas & " C�lulas Modificas no Destino"
    End If

End Function

Private Function fnLePasta(pPastaRaiz As String)

    'Declara��o de vari�veis
    Dim sNomeArquivo    As String
    Dim iContaLinhas    As Integer
    
    'Inicia a contagem das linhas
    iContaLinhas = 2
    Range("B" & iContaLinhas) = sNomeArquivo
    
    'Para cada item na pasta
    sNomeArquivo = Dir(pPastaRaiz, vbDirectory)
    
    'Repete para cada item da pasta
    Do While sNomeArquivo <> vbNullString
    
        'Atribui o �ltimo item lido � linha atual
        Range("B" & iContaLinhas) = sNomeArquivo
        
        'Verifica se o item � uma pasta
        If GetAttr(pPastaRaiz & Range("B" & iContaLinhas)) = vbDirectory Then
            Range("C" & iContaLinhas) = "pasta"
        Else
            Range("C" & iContaLinhas) = "arquivo"
        End If

        'L� o pr�ximo item da pasta
        sNomeArquivo = Dir()
        
        'Incrementa o contador
        iContaLinhas = iContaLinhas + 1
    Loop

End Function
