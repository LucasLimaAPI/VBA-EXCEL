Attribute VB_Name = "Módulo2"
Sub sbManipulaDados()

    'Declaração de variável como célula
    Dim rCelula             As Range
    Dim lContaLinhaDestino  As Long

    'Inicializa a variável
    lContaLinhaDestino = 2
    
    'Estrutura de repetição do tipo For Each
    For Each rCelula In Selection

        If rCelula.Column = 4 Then
            Sheets("Versão Final").Cells(lContaLinhaDestino, rCelula.Column) = fnAjustaData(rCelula.Value)
        Else
            Sheets("Versão Final").Cells(lContaLinhaDestino, rCelula.Column) = rCelula
        End If

        lContaLinhaDestino = lContaLinhaDestino + 1
    Next

End Sub

'Função que ajusta a data de formato americano para brasileiro
Function fnAjustaData(pData As String) As Date
    fnAjustaData = Mid(pData, 9, 2) & "/" & Mid(pData, 6, 2) & "/" & Mid(pData, 1, 4)
End Function
