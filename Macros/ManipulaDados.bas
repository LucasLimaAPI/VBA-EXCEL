Attribute VB_Name = "M�dulo1"
Sub sbManipulaDados()

    'Declara��o de vari�vel como c�lula
    Dim rCelula             As Range
    Dim lContaLinhaDestino  As Long
    Dim sPlanilhaOrigem     As String
    

    'Inicializa a vari�vel
    lContaLinhaDestino = 2
    sPlanilhaOrigem = ActiveSheet.Name
    
    'Verifica se a planilha "Vers�o Final" j� foi criada.
    sbVerificarOuCriarPlanilha
    
    'Selecionar a planilha que cont�m a origem dos dados
    Sheets(sPlanilhaOrigem).Select
    
    'Estrutura de repeti��o do tipo For Each
    For Each rCelula In Selection

        If rCelula.Column = 4 Then
            Sheets("Vers�o Final").Cells(lContaLinhaDestino, rCelula.Column) = fnAjustaData(rCelula.Value)
            lContaLinhaDestino = lContaLinhaDestino + 1
            sbFormataVers�oFinal
        Else
            Sheets("Vers�o Final").Cells(lContaLinhaDestino, rCelula.Column) = rCelula
        End If

        
    Next
    
    'Sub usada para formatar a planilha "Vers�o Final"
    sbFormataVers�oFinal
    

End Sub

'Fun��o que ajusta a data de formato americano para brasileiro
Function fnAjustaData(pData As String) As Date
    fnAjustaData = Mid(pData, 9, 2) & "/" & Mid(pData, 6, 2) & "/" & Mid(pData, 1, 4)
End Function

Private Sub sbVerificarOuCriarPlanilha()

    'Declara��o da vari�vel
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
        'Cria rotulos de dados
        Sheets("Vers�o Final").Cells(1, 1) = "C�digo do Cliente"
        Sheets("Vers�o Final").Cells(1, 2) = "Tipo de Movimenta��o"
        Sheets("Vers�o Final").Cells(1, 3) = "Valor"
        Sheets("Vers�o Final").Cells(1, 4) = "Data"
    Else
        MsgBox "A planilha 'Vers�o Final' j� existe."
    End If
End Sub

Private Sub sbFormataVers�oFinal()
'
' teste_formata Macro
'

'
    Sheets("Vers�o Final").Select
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Columns("C:C").Select
    Selection.Style = "Currency"
    
    Range("A1:D1").Select
    Selection.Font.Bold = True
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ActiveWindow.Zoom = 202
    
        Sheets("Vers�o Final").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    
'Pode passar o erro sem problemas pois n�o vai prejudicar minha planilha
On Error Resume Next
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$3"), , xlYes).Name = _
        "Tabela5"
    Range("Tabela5[#All]").Select
    ActiveSheet.ListObjects("Tabela5").TableStyle = "TableStyleMedium11"
End Sub


