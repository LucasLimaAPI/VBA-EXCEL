Attribute VB_Name = "M�dulo1"
Sub sbManipulaDados()

    'Declara��o de vari�vel como c�lula
    Dim rCelula             As Range
    Dim lContaLinhaDestino  As Long

    'Inicializa a vari�vel
    lContaLinhaDestino = 2
    
    'Estrutura de repeti��o do tipo For Each
    For Each rCelula In Selection

        If rCelula.Column = 4 Then
            Sheets("Vers�o Final").Cells(lContaLinhaDestino, rCelula.Column) = fnAjustaData(rCelula.Value)
            lContaLinhaDestino = lContaLinhaDestino + 1
        Else
            Sheets("Vers�o Final").Cells(lContaLinhaDestino, rCelula.Column) = rCelula
        End If

        
    Next

End Sub

'Fun��o que ajusta a data de formato americano para brasileiro
Function fnAjustaData(pData As String) As Date
    fnAjustaData = Mid(pData, 9, 2) & "/" & Mid(pData, 6, 2) & "/" & Mid(pData, 1, 4)
End Function

Sub sbVerificarOuCriarPlanilha()
    Dim ws As Worksheet
    Dim bExiste As Boolean
    bExiste = False
    
    ' Verifica se a planilha "Vers�o Final" j� existe
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Vers�o Final" Then
            bExiste = True
            Exit For
        End If
    Next ws
    
    ' Se a planilha n�o existir, cria uma nova
    If Not bExiste Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "Vers�o Final"
        MsgBox "A planilha 'Vers�o Final' foi criada."
    Else
        MsgBox "A planilha 'Vers�o Final' j� existe."
    End If
End Sub


