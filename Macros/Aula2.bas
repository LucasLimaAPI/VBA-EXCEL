Attribute VB_Name = "Aula2"
Option Explicit

Sub sbChamaLeituraPasta()
    Call fnLePasta(Range("B1"))
End Sub

Private Function fnLePasta(pPastaRaiz As String)

    'Declara��o de vari�veis
    Dim sNomeArquivo    As String
    Dim iContaLinhas    As Integer
    
    'Inicia a contagem das linhas
    iContaLinhas = 2
    Range("B" & iContaLinhas) = sNomeArquivo
    
    'Para cada item na pasta
    sNomeArquivo = Dir(pPastaRaiz, vbDirectory)
    
    'Repete para cada item da pasta
    Do While sNomeArquivo <> vbNullString
    
        'Atribui o �ltimo item lido � linha atual
        Range("B" & iContaLinhas) = sNomeArquivo
        
        'Verifica se o item � uma pasta
        If GetAttr(pPastaRaiz & Range("B" & iContaLinhas)) = vbDirectory Then
            Range("C" & iContaLinhas) = "pasta"
        Else
            Range("C" & iContaLinhas) = "arquivo"
        End If

        'L� o pr�ximo item da pasta
        sNomeArquivo = Dir()
        
        'Incrementa o contador
        iContaLinhas = iContaLinhas + 1
    Loop

End Function
Sub sbLePastaComObjeto()
    ' Vari�veis para manipula��o de arquivos
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

    ' Construa a mensagem com as informa��es relevantes
    sResultado = "Nome do arquivo: " & f.Name & vbCrLf
    sResultado = sResultado & "Data de cria��o: " & f.DateCreated & vbCrLf
    sResultado = sResultado & "Data de acesso: " & f.DateLastAccessed & vbCrLf
    sResultado = sResultado & "�ltima modifica��o: " & f.DateLastModified

    ' Exiba a mensagem em uma caixa de di�logo
    MsgBox sResultado, 0, "Informa��es de Acesso ao Arquivo"
End Sub
