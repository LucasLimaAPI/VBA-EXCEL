Attribute VB_Name = "Aula1"
'Diretiva para "obrigar" a declaração de variáveis
Option Explicit

Public sNomePlanComparar As String

Sub sbComparaPlanilhas()

    'Declaração de variáveis
    Dim rCelula                 As Range
    Dim iDiferencas             As Integer

    'Verifica se existem pelo menos 2 planilhas na pasta
    If ActiveWorkbook.Sheets.Count <= 1 Then
        MsgBox "Não há planilhas suficientes para comparação"
        Exit Sub
    End If
    
    'Inicializa a variável de diferenças
    iDiferencas = 0
    
    frmEscolhaPlanilha.Show
    
    'Verifica se as demais planilha destino é igual a planilha origem
    For Each rCelula In Selection
        If rCelula.Value <> Sheets(sNomePlanComparar).Range(rCelula.Address) Then
            'Muda a cor da fonte e do interior da célula se ela for diferente da origem
            rCelula.Interior.Color = vbRed
            rCelula.Font.Color = vbYellow
            iDiferencas = iDiferencas + 1
        Else
            'Garante que a célula esteja sem preenchimento e com a fonte "automatic"
            rCelula.Interior.Pattern = xlNone
            rCelula.Font.ColorIndex = xlAutomatic
        End If
    
    Next

    'If que define qual mensagem vai ser mostrada
    If iDiferencas = 0 Then
        MsgBox "Nenhuma Célula Modificada"
    Else
        MsgBox iDiferencas & " Células Modificas no Destino"
    End If

End Sub
