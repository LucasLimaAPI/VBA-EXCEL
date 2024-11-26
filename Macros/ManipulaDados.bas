Attribute VB_Name = "M�dulo2"
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
        Else
            Sheets("Vers�o Final").Cells(lContaLinhaDestino, rCelula.Column) = rCelula
        End If

        lContaLinhaDestino = lContaLinhaDestino + 1
    Next

End Sub

'Fun��o que ajusta a data de formato americano para brasileiro
Function fnAjustaData(pData As String) As Date
    fnAjustaData = Mid(pData, 9, 2) & "/" & Mid(pData, 6, 2) & "/" & Mid(pData, 1, 4)
End Function
