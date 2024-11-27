Attribute VB_Name = "Módulo1"
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
            lContaLinhaDestino = lContaLinhaDestino + 1
        Else
            Sheets("Versão Final").Cells(lContaLinhaDestino, rCelula.Column) = rCelula
        End If

        
    Next

End Sub

'Função que ajusta a data de formato americano para brasileiro
Function fnAjustaData(pData As String) As Date
    fnAjustaData = Mid(pData, 9, 2) & "/" & Mid(pData, 6, 2) & "/" & Mid(pData, 1, 4)
End Function

Sub sbVerificarOuCriarPlanilha()
    Dim ws As Worksheet
    Dim bExiste As Boolean
    bExiste = False
    
    ' Verifica se a planilha "Versão Final" já existe
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Versão Final" Then
            bExiste = True
            Exit For
        End If
    Next ws
    
    ' Se a planilha não existir, cria uma nova
    If Not bExiste Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "Versão Final"
        MsgBox "A planilha 'Versão Final' foi criada."
    Else
        MsgBox "A planilha 'Versão Final' já existe."
    End If
End Sub


