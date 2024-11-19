Attribute VB_Name = "TesteRepeticao"
Option Explicit 'Obriga a gente a declarar a variável.
'COLUNA B: Limpando caracteres diferentes no nome do cliente.
'For linha = 2 To 10
    'Range("B2") = Replace(Range("B2"), "$", "")
    'Cells(linha, 2) = Replace(Cells(linha, 2), "$", "") 'Primeiro a linha e depois a coluna onde a coluna e representada também por um numero
'Next
   
 'Rotina para limpeza de dados dos clientes
Sub sbLimpaDados() 'sb minimonico

    'Declaração de Variável
    Dim lContador As Long
    'Inicializa variável de linha
    lContador = 2
     
     'Cria uma copia da planilha selecionada.
    ActiveSheet.Copy After:=Sheets(1)
    ActiveSheet.Name = "Revisada-" & Format(Now(), "YY-DD-SS")
    
    'Repetição para cada uma das linhas da planilha
             'Trim ele vai ignorar as linhas em branco, 'vbNullString é o mesmo que ""
    Do While Trim(Cells(lContador, 1)) <> vbNullString
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "$", "")
           'COLUNA A: Ajustando o ID do cliente
    If Left(Range("A2"), 7) <> "Zenith_" Then
        Range("A2") = "Zenith_" & Range("A2")
    End If
    
        'COLUNA B: Limpando caracteres diferentes no nome do cliente.
        Range("B2") = Replace(Range("B2"), "#", "")
        Range("B2") = Replace(Range("B2"), "$", "")
        Range("B2") = Replace(Range("B2"), "*", "")
        Range("B2") = Replace(Range("B2"), "%", "")
        Range("B2") = Replace(Range("B2"), "&", "")
        
        'COLUNA C: Ajustando o valor Moeda.
        Range("C2") = Replace(Range("C2"), "R$", "")
        Range("C2") = Replace(Range("C2"), ",", "")
        Range("C2") = Replace(Range("C2"), ".", ",")
        Range("C2").NumberFormat = "_-[$R$-pt-BR] * #,##0.00_-;-[$R$-pt-BR] * #,##0.00_-;_-[$R$-pt-BR] * ""-""??_-;_-@_-"
        
        'COLUNA D: Criando email interno do cliente.
        Range("D2") = Range("A2") & "@zenithbank.com.br"
    
            
            lContador = lContador + 1
        
    Loop
    
    

End Sub

