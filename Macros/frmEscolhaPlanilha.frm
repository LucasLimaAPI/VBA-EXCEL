VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEscolhaPlanilha 
   Caption         =   "Planilha Comparar"
   ClientHeight    =   2340
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "frmEscolhaPlanilha.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEscolhaPlanilha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbPlanilhasPasta_Change()
    sNomePlanComparar = cmbPlanilhasPasta
    Unload Me
End Sub

Private Sub UserForm_Activate()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        If ws.Name <> ActiveSheet.Name Then
            cmbPlanilhasPasta.AddItem ws.Name
        End If
    Next
    
    'Só será usado se houver apenas duas planilhas na pasta trabalho
    'Selecionar automaticamente a segunda planilha
    If cmbPlanilhasPasta.ListCount = 1 Then
        sNomePlanComparar = cmbPlanilhasPasta.List(0)
        Unload Me
    End If
    
End Sub

