Attribute VB_Name = "Módulo1"
Sub SepararEFormatar()
    Dim ws As Worksheet
    Dim rng As Range
    Dim UltimaLinha As Long

    ' Definindo a planilha ativa
    Set ws = ActiveSheet
    
    ' Encontrando a última linha da Coluna A
    UltimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Definindo o range a ser separado
    Set rng = ws.Range("A1:A" & UltimaLinha)
    
    ' Separando os dados com delimitador "|"
    rng.TextToColumns Destination:=ws.Range("A1"), DataType:=xlDelimited, _
                      TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                      Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
                      Other:=True, OtherChar:="|"
    
    ' Aplicando formatação: negrito no cabeçalho
    ws.Rows(1).Font.Bold = True

    MsgBox "Separação e formatação concluídas!", vbInformation
End Sub


