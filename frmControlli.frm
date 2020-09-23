VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   198
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   370
      Left            =   2782
      TabIndex        =   3
      Top             =   2340
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   370
      Left            =   982
      TabIndex        =   2
      Top             =   2340
      Width           =   1500
   End
   Begin VB.TextBox txtNewSkin 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Text            =   "txtNewSkin"
      Top             =   1740
      Width           =   2400
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      Height          =   600
      Left            =   5760
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   6
      Top             =   3600
      WhatsThisHelpID =   10263
      Width           =   600
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   600
      Left            =   5100
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   5
      Top             =   3600
      WhatsThisHelpID =   10262
      Width           =   600
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Stretch image"
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   3840
      WhatsThisHelpID =   10261
      Width           =   1500
   End
   Begin VB.Label lblSkinName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Skin name:"
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   1440
      Width           =   5160
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3060
      Left            =   1860
      Picture         =   "frmControlli.frx":0000
      Top             =   3600
      WhatsThisHelpID =   10259
      Width           =   3120
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1230
      Left            =   60
      Picture         =   "frmControlli.frx":1DE62
      Top             =   60
      WhatsThisHelpID =   10260
      Width           =   6480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPreview_Click()
    gAdattaImmagine = chkPreview
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    '/carica in picTemp per vedere se vi è un errore
    'gPBTemp.AutoRedraw = False
    Set gPBTemp.Picture = LoadPicture(szPath)
    Set gPB.Picture = LoadPicture("")
    If gCheckPreview Then
        gPB.PaintPicture gPBTemp.Picture, 0, 0, gPB.ScaleWidth, gPB.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
    Else
        gPB.PaintPicture gPBTemp.Picture, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
    End If
    gPB.Refresh
    Screen.MousePointer = vbDefault
    gPB.ToolTipText = FileLen(szPath)
    If Err <> 0 Then gPB.ToolTipText = ""
    
End Sub


Private Sub cmdCancel_Click()
   cmdCancel.Tag = "CANCEL"
   gsTestoDiComodo = ""
   Unload Me
End Sub

Private Sub cmdOK_Click()
   cmdOK.Tag = "OK"
   gsTestoDiComodo = Me.txtNewSkin
   Unload Me
End Sub


