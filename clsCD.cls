VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Common Dialog class"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarAction As Integer 'Copia locale.
Private mvarCancelError As Boolean 'Copia locale.
Private mvarDefaultExt As String 'Copia locale.
Private mvarDialogTitle As String 'Copia locale.
Private mvarFileName As String 'Copia locale.
Private mvarFileTitle As String 'Copia locale.
Private mvarFilter As String 'Copia locale.
Private mvarFilterIndex As Integer 'Copia locale.
Private mvarFlags As Long 'Copia locale.
Private mvarHelpCommand As Integer 'Copia locale.
Private mvarHelpContext As Long 'Copia locale.
Private mvarHelpFile As String 'Copia locale.
Private mvarHelpKey As String 'Copia locale.
Private mvarInitDir As String 'Copia locale.
Private mvarMaxFileSize As Integer 'Copia locale.
Private mvarStyle As COMMON_DIALOG_STYLE 'Copia locale.

Public Enum enumFILTER
    TUTTI = 1
    BMP = 2
    DOC = 4
    EMF = 8
    EXE = 16
    GIF = 32
    ICO = 64
    JPG = 128
    MIDI = 256
    TXT = 512
    WAVE = 1024
    WMF = 2048
End Enum
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarFilter2 As enumFILTER 'Copia locale.

Public Enum enumFONTTYPE
  SIMULATED_FONTTYPE = &H8000
  PRINTER_FONTTYPE = &H4000
  SCREEN_FONTTYPE = &H2000
  BOLD_FONTTYPE = &H100
  ITALIC_FONTTYPE = &H200
  REGULAR_FONTTYPE = &H400
End Enum
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private mvarFontType As enumFONTTYPE 'Copia locale.
Public Property Let FontType(ByVal vData As enumFONTTYPE)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.FontType = 5
    mvarFontType = vData
End Property


Public Property Get FontType() As enumFONTTYPE
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.FotnType
    FontType = mvarFontType
End Property




Public Function ShowFont(ByVal myForm As Form) As String
   ShowFont = pShowFont(myForm)
End Function

Public Property Let Filter2(ByVal vData As enumFILTER)
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.

Dim s As String

    If vData = TUTTI Then
      s = Chr(0) & "All files (*.*)" & Chr(0) & "*.*"
    Else
      If vData And BMP Then
        s = s & Chr(0) & "Bitmap images (*.bmp)" & Chr(0) & "*.bmp"
      End If
      
      If vData And DOC Then
      s = s & Chr(0) & "Word documents (*.doc)" & Chr(0) & "*.doc"
      End If
      
      If vData And EMF Then
        s = s & Chr(0) & "Enhanced Metafile (*.emf)" & Chr(0) & "*.emf"
      End If
      
      If vData And EXE Then
        s = s & Chr(0) & "Applications (*.exe)" & Chr(0) & "*.exe"
      End If
      
      If vData And GIF Then
        s = s & Chr(0) & "GIF images (*.gif)" & Chr(0) & "*.gif"
      End If
      
      If vData And ICO Then
        s = s & Chr(0) & "Icon images (*.ico)" & Chr(0) & "*.ico"
      End If
      
      If vData And JPG Then
        s = s & Chr(0) & "JPG images (*.jpg)" & Chr(0) & "*.jpg"
      End If
      
      If vData And MIDI Then
        s = s & Chr(0) & "MIDI sequence (*.mid)" & Chr(0) & "*.mid"
      End If
      
      If vData And TXT Then
        s = s & Chr(0) & "Text documents (*.txt)" & Chr(0) & "*.txt"
      End If
      
      If vData And WAVE Then
        s = s & Chr(0) & "Wave audio files (*.wav)" & Chr(0) & "*.wav"
      End If
            
      If vData And WMF Then
        s = s & Chr(0) & "Windows Metafile (*.wmf)" & Chr(0) & "*.wmf"
      End If
      
    End If
    
    If s = "" Then
      s = Chr(0) & "All files (*.*)" & Chr(0) & "*.*"
    End If
    
    mvarFilter2 = vData
    
    mvarFilter = Mid$(s, 2)
    
End Property


Public Property Get Filter2() As enumFILTER
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Filter2
    'Set Filter2 = mvarFilter2
    Filter2 = mvarFilter2
End Property







Public Property Let Style(ByVal vData As COMMON_DIALOG_STYLE)
Attribute Style.VB_Description = "Imposta lo stile di personalizzazione della finestra."
'Utilizzato per l'assegnazione di un valore alla proprietà, a sinistra di un'assegnazione.
'Syntax: X.Style = 5
    mvarStyle = vData
    giCommonDialogStyle = vData
    
End Property


Public Property Get Style() As COMMON_DIALOG_STYLE
'Utilizzato per recuperare il valore di una proprietà, a destra dell'assegnazione.
'Syntax: Debug.Print X.Style
    Style = mvarStyle
End Property



Public Sub ShowSave()
Attribute ShowSave.VB_Description = "Visualizza la finestra di dialogo Salva con nome."
End Sub

Public Function ShowOpen(myForm As Form) As String
Attribute ShowOpen.VB_Description = "Visualizza la finestra di dialogo Apri."

  Dim FError As Long
  
  If Style = OPENFILE_DEFAULT Then
      Rem If set the Open dialog with new XP style I can't
      Rem use callback... :-(
      Flags = 0
  Else
      Rem All other dialog use callback! :-)
      Flags = OFN_EXPLORER Or _
            OFN_HIDEREADONLY Or _
            OFN_LONGNAMES Or _
            OFN_PATHMUSTEXIST Or _
            OFN_ENABLEHOOK Or _
            OFN_FILEMUSTEXIST
  End If
  
  Rem ------------------------------------------------------------
  Rem Set the title dialog in accord to CD style
  If Style = OPENFILE_AUDIO Then
    DialogTitle = "Listen audio: select the audio file then click PLAY"
  ElseIf Style = OPENFILE_DELETEFILE Then
    DialogTitle = "Select the file to delete"
  End If
  Rem Show Open dialog...
  ShowOpen = pFileOpen(myForm, FError&, mvarFilter, mvarInitDir, mvarDialogTitle, mvarFilterIndex, mvarFlags)
  
End Function

Public Property Let MaxFileSize(ByVal vData As Integer)
Attribute MaxFileSize.VB_Description = "Restituisce o imposta la dimensione massima dei nomi dei file che è possibile aprire."
'Syntax: X.MaxFileSize = 5
    mvarMaxFileSize = vData
End Property


Public Property Get MaxFileSize() As Integer

'Syntax: Debug.Print X.MaxFileSize
    MaxFileSize = mvarMaxFileSize
End Property



Public Property Let InitDir(ByVal vData As String)
Attribute InitDir.VB_Description = "Restituisce o imposta la directory iniziale per i file."

'Syntax: X.InitDir = 5
    mvarInitDir = vData
End Property


Public Property Get InitDir() As String
'Syntax: Debug.Print X.InitDir
    InitDir = mvarInitDir
End Property



Public Property Let HelpKey(ByVal vData As String)
Attribute HelpKey.VB_Description = "Restituisce o imposta la parola chiave che identifica l'argomento della Guida richiesto."
'Syntax: X.HelpKey = 5
    mvarHelpKey = vData
End Property


Public Property Get HelpKey() As String
'Syntax: Debug.Print X.HelpKey
    HelpKey = mvarHelpKey
End Property



Public Property Let HelpFile(ByVal vData As String)
Attribute HelpFile.VB_Description = "Specifica il percorso e il nome di un file della Guida utilizzato dall'applicazione in esecuzione per visualizzare la Guida o la documentazione in linea."
'Syntax: X.HelpFile = 5
    mvarHelpFile = vData
End Property


Public Property Get HelpFile() As String
'Syntax: Debug.Print X.HelpFile
    HelpFile = mvarHelpFile
End Property



Public Property Let HelpContext(ByVal vData As Long)
Attribute HelpContext.VB_Description = "Restituisce o imposta l'ID di contesto dell'argomento della Guida in linea richiesto."
'Syntax: X.HelpContext = 5
    mvarHelpContext = vData
End Property


Public Property Get HelpContext() As Long
'Syntax: Debug.Print X.HelpContext
    HelpContext = mvarHelpContext
End Property



Public Property Let HelpCommand(ByVal vData As Integer)
Attribute HelpCommand.VB_Description = "Restituisce o imposta il tipo di Guida in linea richiesto."
'Syntax: X.HelpCommand = 5
    mvarHelpCommand = vData
End Property


Public Property Get HelpCommand() As Integer
'Syntax: Debug.Print X.HelpCommand
    HelpCommand = mvarHelpCommand
End Property



Public Property Let Flags(ByVal vData As Long)
Attribute Flags.VB_Description = "Restituisce o imposta le opzioni relative alle finestre di dialogo Apri e Salva con nome."
'Syntax: X.Flags = 5
    mvarFlags = vData
End Property


Public Property Get Flags() As Long
'Syntax: Debug.Print X.Flags
    Flags = mvarFlags
End Property



Public Property Let FilterIndex(ByVal vData As Integer)
Attribute FilterIndex.VB_Description = "Restituisce o imposta un filtro predefinito per una finestra di dialogo Apri o Salva con nome."
'Syntax: X.FilterIndex = 5
    mvarFilterIndex = vData
End Property


Public Property Get FilterIndex() As Integer
'Syntax: Debug.Print X.FilterIndex
    FilterIndex = mvarFilterIndex
End Property



Public Property Let Filter(ByVal vData As String)
Attribute Filter.VB_Description = "Restituisce o imposta i filtri visualizzati nellacasella di riepilogo Tipo file di una finestra di dialogo."
'Syntax: X.Filter = 5
    mvarFilter = vData
End Property


Public Property Get Filter() As String
'Syntax: Debug.Print X.Filter
    Filter = mvarFilter
End Property



Public Property Let FileTitle(ByVal vData As String)
Attribute FileTitle.VB_Description = "Restituisce il nome (ma non il percorso) del file da aprire o salvare."
'Syntax: X.FileTitle = 5
    mvarFileTitle = vData
End Property


Public Property Get FileTitle() As String
'Syntax: Debug.Print X.FileTitle
    FileTitle = mvarFileTitle
End Property



Public Property Let FileName(ByVal vData As String)
Attribute FileName.VB_Description = "Restituisce o imposta il percorso e il nome di un file selezionato."
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



Public Property Let DialogTitle(ByVal vData As String)
Attribute DialogTitle.VB_Description = "Restituisce o imposta la stringa visualizzata nellabarra del titolo della finestra di dialogo."
'Syntax: X.DialogTitle = 5
    mvarDialogTitle = vData
End Property


Public Property Get DialogTitle() As String
'Syntax: Debug.Print X.DialogTitle
    DialogTitle = mvarDialogTitle
End Property



Public Property Let DefaultExt(ByVal vData As String)
Attribute DefaultExt.VB_Description = "Restituisce o imposta l'estensione predefinita per i nomi di file visualizzati nella finestra di dialogo."
'Syntax: X.DefaultExt = 5
    mvarDefaultExt = vData
End Property


Public Property Get DefaultExt() As String
'Syntax: Debug.Print X.DefaultExt
    DefaultExt = mvarDefaultExt
End Property



Public Property Let CancelError(ByVal vData As Boolean)
Attribute CancelError.VB_Description = "Restituisce o imposta un valore che indica se la scelta del pulsante Annulla da parte dell'utente genera o meno un errore."
'Syntax: X.CancelError = 5
    mvarCancelError = vData
End Property


Public Property Get CancelError() As Boolean
'Syntax: Debug.Print X.CancelError
    CancelError = mvarCancelError
End Property



Public Property Let Action(ByVal vData As Integer)
Attribute Action.VB_Description = "Restituisce o imposta il tipo di finestra di dialogo da visualizzare."
'Syntax: X.Action = 5
    mvarAction = vData
End Property


Public Property Get Action() As Integer
'Syntax: Debug.Print X.Action
    Action = mvarAction
End Property




Public Function ShowColor(ByVal myForm As Form) As Long
  Dim CError As Long  ' error code returned
  ShowColor = pShowColors(myForm, CError, myForm.BackColor, 0&)
  
End Function
