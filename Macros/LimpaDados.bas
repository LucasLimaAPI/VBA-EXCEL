Attribute VB_Name = "LimpaDados"
Sub sbLimpaDados() 'sb minimonico

    'Cria uma copia da planilha selecionada.
    ActiveSheet.Copy After:=Sheets(1)
    ActiveSheet.Name = "Revisada-" & Format(Now(), "YY-DD-SS")


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

End Sub
