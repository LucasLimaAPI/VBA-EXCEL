Attribute VB_Name = "Aula3"
Option Explicit
Sub fnComparaPastaDeTrabalhoExcel()
    
    If fnAbrePastaDeTrabalho("C:\Users\lucas.oliveira\Desktop\VBA\Materiais\Planilha Modelo.xlsx") Then
        MsgBox "Pasta de trabalho manipulada com sucesso"
    Else
        MsgBox "Pasta de trabalho não pode ser aberta"
    End If

End Sub


Private Function fnAbrePastaDeTrabalho(pCaminhoCompletoWB As String) As Boolean
    'Declaração das variáveis de objeto
    Dim wb As Workbook      'Pasta de trabalho
    Dim ws As Worksheet     'Planilha
    
    fnAbrePastaDeTrabalho = False
    
    '"Seta" As varáveis na "pasta de trabalho" e na  "Planilha de Trabalho"
    Set wb = Workbooks.Open(pCaminhoCompletoWB)
    Set ws = wb.ActiveSheet
    
    'Mostra o nome da planilha ativa no momento
    MsgBox ws.Name
    
    'Fecha a pasta de trabalho sem salvar
    wb.Close False
    
    'Limpa os Objetos criados
    Set wb = Nothing
    Set ws = Nothing
    
    fnAbrePastaDeTrabalho = True
    
End Function
