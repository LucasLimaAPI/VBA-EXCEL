Attribute VB_Name = "Aula1"
'Diretiva para "obrigar" a declara��o de vari�veis
Option Explicit

Public sNomePlanComparar As String

Sub sbComparaPlanilhas()

    'Declara��o de vari�veis
    Dim rCelula                 As Range
    Dim iDiferencas             As Integer

    'Verifica se existem pelo menos 2 planilhas na pasta
    If ActiveWorkbook.Sheets.Count <= 1 Then
        MsgBox "N�o h� planilhas suficientes para compara��o"
        Exit Sub
    End If
    
    'Inicializa a vari�vel de diferen�as
    iDiferencas = 0
    
    frmEscolhaPlanilha.Show
    
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

End Sub
