Attribute VB_Name = "TesteRepeticao"
Option Explicit 'Obriga a gente a declarar a vari�vel.

    'COLUNA B: Limpando caracteres diferentes no nome do cliente.
   
   'For linha = 2 To 10
        'Range("B2") = Replace(Range("B2"), "$", "")
        'Cells(linha, 2) = Replace(Cells(linha, 2), "$", "") 'Primeiro a linha e depois a coluna onde a coluna e representada tamb�m por um numero
   'Next
   
Sub sbTesteRepeticao()
    'Declara��o de Vari�vel
    Dim lContador As Long
    
     'COLUNA B: Limpando caracteres diferentes no nome do cliente.

             'Trim ele vai ignorar as linhas em branco, 'vbNullString � o mesmo que ""
    Do While Trim(Cells(lContador, 2)) <> vbNullString
        Cells(lContador, 2) = Replace(Cells(lContador, 2), "$", "")
        
        lContador = lContador + 1
    
    Loop
    
    

End Sub

