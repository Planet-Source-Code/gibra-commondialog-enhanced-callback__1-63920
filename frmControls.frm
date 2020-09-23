VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmControls 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmControls"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   6405
      TabIndex        =   7
      Top             =   1185
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Suona"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Ferma il suono"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmControls.frx":0000
      Top             =   2910
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6840
      ScaleHeight     =   375
      ScaleWidth      =   465
      TabIndex        =   8
      Top             =   1680
      Width           =   465
   End
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
      Left            =   3330
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   6
      Top             =   3615
      WhatsThisHelpID =   10263
      Width           =   600
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   600
      Left            =   2670
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   5
      Top             =   3615
      WhatsThisHelpID =   10262
      Width           =   600
   End
   Begin VB.CheckBox chkPreview 
      Caption         =   "Stretch"
      Height          =   210
      Left            =   2670
      TabIndex        =   4
      Top             =   4245
      WhatsThisHelpID =   10261
      Width           =   1500
   End
   Begin VB.Image Image4 
      Height          =   3390
      Left            =   4875
      Picture         =   "frmControls.frx":0006
      Top             =   1920
      Width           =   1800
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6435
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   327682
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   300
      Picture         =   "frmControls.frx":1A9E
      Top             =   5430
      Width           =   6420
   End
   Begin VB.Image imgStop 
      Height          =   240
      Left            =   6765
      Picture         =   "frmControls.frx":8992
      Top             =   825
      Width           =   240
   End
   Begin VB.Image imgPlay 
      Height          =   240
      Left            =   6405
      Picture         =   "frmControls.frx":8F1C
      Top             =   825
      Width           =   240
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
      Height          =   2265
      Left            =   150
      Picture         =   "frmControls.frx":94A6
      Top             =   2910
      WhatsThisHelpID =   10259
      Width           =   2400
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1380
      Left            =   60
      Picture         =   "frmControls.frx":F1DE
      Top             =   60
      WhatsThisHelpID =   10260
      Width           =   6255
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
   ByVal lpszName As String, _
   ByVal hModule As Long, _
   ByVal dwFlags As Long) As Long
   
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000
Private Const SND_PURGE As Long = &H40
Private Const SND_SYNC As Long = &H0

Private Sub chkPreview_Click()
    gAdattaImmagine = chkPreview
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    Rem Load in picTemp for error test
    Set gPBTemp.Picture = LoadPicture(gszPath)
    Set gPB.Picture = LoadPicture("")
    If gCheckPreview Then
        gPB.PaintPicture gPBTemp.Picture, 0, 0, gPB.ScaleWidth, gPB.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
    Else
        gPB.PaintPicture gPBTemp.Picture, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
    End If
    gPB.Refresh
    Screen.MousePointer = vbDefault
    
    gPB.ToolTipText = FileLen(gszPath)
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


Private Sub Form_Load()
  
  With ImageList1
    Picture1.Picture = imgPlay.Picture
    Picture1.Picture = Picture1.Image
    Picture1.Refresh
    .ListImages.Add , , Picture1.Picture
    Picture1.Picture = imgStop.Picture
    .ListImages.Add , , Picture1.Picture
  End With
  
  With Toolbar1
    .ImageList = ImageList1
    With .Buttons(1)
      .ToolTipText = "Play"
      .Image = 1
      .Key = "play"
    End With
    With .Buttons(2)
      .ToolTipText = "Stop"
      .Image = 2
      .Key = "stop"
    End With
  End With
  
End Sub



Public Sub pPlaySound(ByVal sFileName As String)
  PlaySound sFileName, App.hInstance, SND_ASYNC Or SND_FILENAME
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  Dim lRet As Long
  If gszPath = "" Then Exit Sub
  
  Select Case Button.Key
    
    Case "play"
      Rem First example, play audio file with winmm.dll function:
      If Right(UCase(gszPath), 4) = ".WAV" Then
        pPlaySound gszPath  ' play selected audio file
      ElseIf Right(UCase(gszPath), 4) = ".MID" Then
        Rem Second example, play audio file from external program:
        Rem MPlay32.exe will Open, Play and Close the audio file
        Rem automatically. If you don't show the program, use this:
        Rem Shell "mplay32.exe /Play /Close " & gszPath, vbHide
        Rem Caution, in this way you cannot close directly MPlay32.exe
        Rem but you must wait until MPlay32.exe close itself.
        On Error Resume Next  ' If don't find the program...
        Shell "mplay32.exe /Play /Close " & gszPath, vbNormalFocus
        If Err.Number <> 0 Then
          MsgBox "Cannot play the MIDI sequence.", vbExclamation
        End If
      End If
      
    Case "stop"
      If Right(UCase(gszPath), 4) = ".WAV" Then
        Rem Audio file
        pStopSound
      Else
        Rem MIDI sequence
        lRet = StopMPlay32()
      End If
      
  End Select
  

  
End Sub



Public Sub pStopSound()
  PlaySound 0&, 0&, SND_PURGE
End Sub
