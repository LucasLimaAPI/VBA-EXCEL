Attribute VB_Name = "Aula2"
Option Explicit

Sub sbChamaLeituraPasta()
    Call fnLePasta(Range("B1"))
End Sub

Private Function fnLePasta(pPastaRaiz As String)

    'Declaração de variáveis
    Dim sNomeArquivo    As String
    Dim iContaLinhas    As Integer
    
    'Inicia a contagem das linhas
    iContaLinhas = 2
    Range("B" & iContaLinhas) = sNomeArquivo
    
    'Para cada item na pasta
    sNomeArquivo = Dir(pPastaRaiz, vbDirectory)
    
    'Repete para cada item da pasta
    Do While sNomeArquivo <> vbNullString
    
        'Atribui o último item lido à linha atual
        Range("B" & iContaLinhas) = sNomeArquivo
        
        'Verifica se o item é uma pasta
        If GetAttr(pPastaRaiz & Range("B" & iContaLinhas)) = vbDirectory Then
            Range("C" & iContaLinhas) = "pasta"
        Else
            Range("C" & iContaLinhas) = "arquivo"
        End If

        'Lê o próximo item da pasta
        sNomeArquivo = Dir()
        
        'Incrementa o contador
        iContaLinhas = iContaLinhas + 1
    Loop

End Function
Sub sbLePastaComObjeto()
    ' Variáveis para manipulação de arquivos
    Dim fs           As Object
    Dim f            As Object
    Dim sResultado   As String
    Dim filespec     As String

    ' Defina o caminho completo do arquivo que deseja verificar
    ' Altere o caminho de acordo com a sua pasta
    filespec = "C:\Users\lucas.oliveira\Desktop\VBA\Materiais\Planilha Modelo.xlsx"

    ' Crie um objeto FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")

    ' Obtenha o objeto File correspondente ao arquivo
    Set f = fs.GetFile(filespec)

    ' Construa a mensagem com as informações relevantes
    sResultado = "Nome do arquivo: " & f.Name & vbCrLf
    sResultado = sResultado & "Data de criação: " & f.DateCreated & vbCrLf
    sResultado = sResultado & "Data de acesso: " & f.DateLastAccessed & vbCrLf
    sResultado = sResultado & "Última modificação: " & f.DateLastModified

    ' Exiba a mensagem em uma caixa de diálogo
    MsgBox sResultado, 0, "Informações de Acesso ao Arquivo"
End Sub
