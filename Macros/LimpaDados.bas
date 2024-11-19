Attribute VB_Name = "LimpaDados"
Option Explicit 'Obriga a gente a declarar a variável.
'COLUNA B: Limpando caracteres diferentes no nome do cliente.
'For linha = 2 To 10
    'cells(2,2) = Replace(cells(2,2), "$", "")
    'Cells(linha, 2) = Replace(Cells(linha, 2), "$", "") 'Primeiro a linha e depois a coluna onde a coluna e representada também por um numero
'Next
   
 'Rotina para limpeza de dados dos clientes
Sub sbLimpaDados() 'sb minimonico
Attribute sbLimpaDados.VB_Description = "Limpa Dados do saldo do cliente formata informações"
Attribute sbLimpaDados.VB_ProcData.VB_Invoke_Func = "R\n14"

    'Declaração de Variável
    Dim lContador As Long
    
    'Inicializa variável de linha
    lContador = 2
     
     'Cria uma copia da planilha selecionada.
    ActiveSheet.Copy After:=Sheets(1)
    ActiveSheet.Name = "Revisada-" & Format(Now(), "YY-DD-SS")
    
    'Repetição para cada uma das linhas da planilha
             'Trim ele vai tirar caracteres e linhas em branco, 'vbNullString é o mesmo que ""
    Do While Trim(Cells(lContador, 1)) <> vbNullString
        
           'COLUNA A: Ajustando o ID do cliente
    If Left(Cells(lContador, 1), 7) <> "Zenith_" Then
        Cells(lContador, 1) = "Zenith_" & Cells(lContador, 1)
    End If
    
        'COLUNA B: Limpando caracteres diferentes no nome do cliente.
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "#", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "$", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "*", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "%", "")
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "&", "")
        
        'COLUNA C: Ajustando o valor Moeda.
        Cells(lContador, 3) = Replace(Cells(lContador, 3), "R$", "")
        Cells(lContador, 3) = Replace(Cells(lContador, 3), ",", "")
        Cells(lContador, 3) = Replace(Cells(lContador, 3), ".", ",")
        Cells(lContador, 3).NumberFormat = "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
        
        'COLUNA D: Criando email interno do cliente.
        Cells(lContador, 4) = Cells(lContador, 1) & "@zenithbank.com.br"
    
            
            lContador = lContador + 1
        
    Loop
    
    'Formata como tabela
     Range("A1").Select
     Range(Selection, Selection.End(xlDown)).Select
     Range(Selection, Selection.End(xlToRight)).Select
     Application.CutCopyMode = False
     ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$10"), , xlYes).Name = "Tabela1"
    

End Sub

