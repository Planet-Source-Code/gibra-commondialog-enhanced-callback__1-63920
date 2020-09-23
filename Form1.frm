VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CD As clsCommDlg
Private Sub Form_Load()
  Set CD = New clsCommDlg
  
  
  With CD
    .CancelError = True
    .Filter2 = EMF + EXE + JPG
    'Debug.Print .Filter
  End With
    
End Sub


