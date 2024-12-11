VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovimentacao 
   Caption         =   "Formulário de Movimentação"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   OleObjectBlob   =   "frmMovimentacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdSair_Click()
    Unload Me
End Sub



Private Sub cmdSalvar_Click()

        'Consistencia de data invalida para o campo "Data"
        If Not IsDate(txtData) Then
        MsgBox "Digite uma data válida!"
            Exit Sub
        End If
        
         If lblTipoForm = "inclusão" Then
            Range("A1").Select
            Selection.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
         End If
         
        Range("A" & Selection.Row) = txtAtivo
        Range("B" & Selection.Row) = CDbl(txtQtd) 'CDBL DOUBLE - MUDA VALOR PARA DOUBLE
        Range("C" & Selection.Row) = cmdTipo
        Range("D" & Selection.Row) = CCur(txtPreco) 'CCUR Currency - moeda
        Range("E" & Selection.Row) = txtCliente
        Range("F" & Selection.Row) = cmdContato
        Range("G" & Selection.Row) = txtData
        Range("H" & Selection.Row) = txtHora
        
        cmdSair_Click
End Sub




Private Sub UserForm_Activate()
    
    'Carrega o combo de compra e venda
    cmdTipo.AddItem "Compra"
    cmdTipo.AddItem "Venda"
    
    'carrega o combo de cotatos da mesa
    For contador = 2 To 6
        cmdContato.AddItem Planilha2.Cells(contador, 1)
    Next
    
    
    'Verifica o tipo de operacao do formulario
    'Só carrega os dados para alteração e não para inclusão
    If lblTipoForm.Caption = "Alteração" Then
    
        txtAtivo = Range("A" & Selection.Row)
        txtQtd = Range("B" & Selection.Row)
        cmdTipo = Range("C" & Selection.Row)
        txtPreco = Range("D" & Selection.Row)
        txtCliente = Range("E" & Selection.Row)
        cmdContato = Range("F" & Selection.Row)
        txtData = Range("G" & Selection.Row)
        txtHora = Format(Range("H" & Selection.Row), "HH:mm")
        
        Else
            txtData = Format(Now(), "dd/mm/yyyy")
            txtHora = Format(Now(), "HH:mm")
            
    End If
    
End Sub

Private Sub txtData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Verifica as teclas de 1 a 9 no teclado alfanumérico
    If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
        
        'Caso o comprimento da data seja 2 ou 5, adiciona as barras
        If Len(Trim(Me.ActiveControl)) = 2 Or Len(Trim(Me.ActiveControl)) = 5 Then
            Me.ActiveControl = Trim(Me.ActiveControl) & "/"
        End If

    'Verifica as teclas de 1 a 9 no teclado numérico
    ElseIf KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
        
        'Caso o comprimento da data seja 2 ou 5, adiciona as barras
        If Len(Trim(Me.ActiveControl)) = 2 Or Len(Trim(Me.ActiveControl)) = 5 Then
            Me.ActiveControl = Trim(Me.ActiveControl) & "/"
        End If

    'Deixa pronto o tratamento para a tecla "delete" e libera seu uso
    ElseIf KeyCode = vbKeyDelete Then

    'Deixa pronto o tratamento para a tecla "backspace" e libera seu uso
    ElseIf KeyCode = vbKeyBack Then

    'Caso não seja nenhum dos casos acima, cancela a digitação
    Else
        KeyCode = vbKeyCancel
    End If

End Sub

Private Sub txtPreco_Keydown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    
    'Verifica as teclas de 1 a 9 no teclado alfanumérico
    If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
    
    'Verifica as teclas de 1 a 9 no teclado numérico
    ElseIf KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
        
    'Deixa pronto o tratamento para a vírgula "," e libera seu uso
    ElseIf KeyCode = 188 Then
    
    'Deixa pronto o tratamento para a tecla "delete" e libera seu uso
    ElseIf KeyCode = vbKeyDelete Then
        
    'Deixa pronto o tratamento para a tecla "backspace" e libera seu uso
    ElseIf KeyCode = vbKeyBack Then
        
    'Caso não seja nenhum dos casos acima, cancela a digitação
    Else
        KeyCode = vbKeyCancel
    
    End If

End Sub


Private Sub txtQtd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Verifica as teclas de 1 a 9 no teclado alfanumérico
    If KeyCode >= vbKey0 And KeyCode <= vbKey9 Then
    
    'Verifica as teclas de 1 a 9 no teclado numérico
    ElseIf KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
        
    'Deixa pronto o tratamento para a tecla "delete" e libera seu uso
    ElseIf KeyCode = vbKeyDelete Then
        
    'Deixa pronto o tratamento para a tecla "backspace" e libera seu uso
    ElseIf KeyCode = vbKeyBack Then
        
    'Caso não seja nenhum dos casos acima, cancela a digitação
    Else
        KeyCode = vbKeyCancel
    
    End If
End Sub







'Initialize é executado antes do activate, carregando antes mesmo das informações carregadas.

'Private Sub Userform_initialize()
 '   MsgBox "Evento initialize"
'End Sub
