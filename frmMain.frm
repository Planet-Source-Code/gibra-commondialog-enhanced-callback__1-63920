VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Common Dialogs Enhanced"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optStyle 
      Caption         =   "Open dialog (enhanced with left custom image)"
      Height          =   225
      Index           =   6
      Left            =   420
      TabIndex        =   9
      Top             =   3105
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4845
      TabIndex        =   8
      Top             =   4755
      Width           =   1365
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Colors dialog"
      Height          =   225
      Index           =   5
      Left            =   420
      TabIndex        =   7
      Top             =   3915
      Width           =   2595
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Font dialog"
      Height          =   225
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   3525
      Width           =   2595
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Open dialog (to delete file)"
      Height          =   225
      Index           =   3
      Left            =   420
      TabIndex        =   3
      Top             =   2715
      Width           =   4005
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Open dialog (enhanced with top custom image, with audio preview!)"
      Height          =   225
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   2310
      Width           =   5685
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Open dialog (enhanced, with image preview)"
      Height          =   225
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1920
      Width           =   5000
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Open dialog (normal, default whitout callback)"
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   1515
      Value           =   -1  'True
      Width           =   5000
   End
   Begin VB.CommandButton cmdOpenDialog 
      Caption         =   "Open..."
      Height          =   375
      Left            =   4845
      TabIndex        =   5
      Top             =   4200
      Width           =   1365
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   90
      Top             =   90
      Width           =   4425
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   420
      Picture         =   "frmMain.frx":058A
      Top             =   4425
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a window style and click Open button:"
      Height          =   195
      Left            =   420
      TabIndex        =   6
      Top             =   1155
      Width           =   5000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CD As clsCommDlg

Private Sub cmdExit_Click()
  Unload Me
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : cmdOpenDialog_Click
' DateTime    : 26/09/2005 05.10
' Author      : Giorgio Brausi
' Purpose     :
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Private Sub cmdOpenDialog_Click()

  Dim i As Integer
  Dim sReturn As String
  Dim lRetColor As Long
  Const BTN_ANNULLA = 3
  Const OPEN_FONT_DIALOG = 4
  Const OPEN_COLOR_DIALOG = 5

  If CD.Style = OPEN_FONT_DIALOG Then
    sReturn = CD.ShowFont(Me)
    If sReturn <> "" Then
      MsgBox UCase(sReturn)
    End If
  
  ElseIf CD.Style = OPEN_COLOR_DIALOG Then
    lRetColor = CD.ShowColor(Me)
    If lRetColor = BTN_ANNULLA Then Exit Sub
    If MsgBox("Do you want to apply the selected color to form?", vbYesNo + vbQuestion) = vbYes Then
      Me.BackColor = lRetColor
      For i = optStyle.lBound To optStyle.UBound
        optStyle(i).BackColor = lRetColor
      Next i
    End If
  
  Else
  
    Rem file Open/Save
    sReturn = CD.ShowOpen(Me)
    If sReturn <> "" Then
      MsgBox "Selected file: " & vbCrLf & vbCrLf & UCase(sReturn)
    End If
  End If
  
End Sub

Private Sub Form_Load()

  Set CD = New clsCommDlg
  With CD
    .CancelError = True
    .InitDir = App.Path
  End With
  
  Rem Set 'normal' style
  optStyle_Click (0)
  
  Image2.Picture = frmControls.Image3.Picture
  Unload frmControls
  Set frmControls = Nothing
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set CD = Nothing
End Sub


Private Sub optStyle_Click(Index As Integer)
  
  Rem Set the dialog style
  CD.Style = Index
  
'  Const OPENFILE_PICTURE = 1
'  Const OPENFILE_AUDIO = 2
  
  If CD.Style = OPENFILE_AUDIO Then
    Rem Set filters to audio files
    CD.Filter2 = WAVE + MIDI
    CD.FilterIndex = 2
    
  ElseIf CD.Style = OPENFILE_PICTURE Then
    Rem Set filter to image files
    CD.Filter2 = BMP + EMF + GIF + ICO + JPG + WMF
  Else
    Rem Set filter to All files (*.*)
    CD.Filter2 = TUTTI
  End If
  
End Sub


