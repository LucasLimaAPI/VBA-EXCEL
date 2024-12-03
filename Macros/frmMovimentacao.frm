VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovimentacao 
   Caption         =   "Formul�rio de Movimenta��o"
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
        MsgBox "Digite uma data v�lida!"
            Exit Sub
        End If
        
        Range("A" & Selection.Row) = txtAtivo
        Range("B" & Selection.Row) = txtQtd
        Range("C" & Selection.Row) = txtTipo
        Range("D" & Selection.Row) = txtPreco
        Range("E" & Selection.Row) = txtCliente
        Range("F" & Selection.Row) = txtContato
        Range("G" & Selection.Row) = txtData
        Range("H" & Selection.Row) = txtHora
End Sub



Private Sub txtData_Change()

End Sub

Private Sub UserForm_Activate()
    'Verifica o tipo de operacao do formulario
    'S� carrega os dados para altera��o e n�o para inclus�o
    If lblTipoForm.Caption = "Altera��o" Then
    
        txtAtivo = Range("A" & Selection.Row)
        txtQtd = Range("B" & Selection.Row)
        txtTipo = Range("C" & Selection.Row)
        txtPreco = Range("D" & Selection.Row)
        txtCliente = Range("E" & Selection.Row)
        txtContato = Range("F" & Selection.Row)
        txtData = Range("G" & Selection.Row)
        txtHora = Format(Range("H" & Selection.Row), "HH:mm")
    End If
    
End Sub

'Initialize � executado antes do activate, carregando antes mesmo das informa��es carregadas.

'Private Sub Userform_initialize()
 '   MsgBox "Evento initialize"
'End Sub
