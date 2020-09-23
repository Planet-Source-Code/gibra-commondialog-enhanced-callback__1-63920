Attribute VB_Name = "modCMDialog"
'---------------------------------------------------------------------------------------
' Module      : modCMDialog
' DateTime    : 26/09/2005 04.57
' Author      : Giorgio Brausi
' Project     : CommonDialog_Class
' Purpose     : Windows Common Dialogs access  (32bit)
' Descritpion : Enhanced Custom for common dialog
' Comments    : use CallBack
'---------------------------------------------------------------------------------------
Option Explicit

Rem Costants
Public gsTestoDiComodo As String
#If False Then
  Public OPENFILE_DEFAULT
  Public OPENFILE_PICTURE
  Public OPENFILE_AUDIO
  Public OPENFILE_DELETEFILE
  Public OPEN_FONT_DIALOG
  Public OPEN_COLOR_DIALOG
  Public OPENFILE_LIST
#End If
Public Enum COMMON_DIALOG_STYLE
  OPENFILE_DEFAULT = 0
  OPENFILE_PICTURE = 1        ' FileOpen with IMAGE PREVIEW
  OPENFILE_AUDIO = 2          ' FileOpen with TOP custom image + Play & Stop buttons
  OPENFILE_DELETEFILE = 3     ' FileOpen to delete selected file
  OPEN_FONT_DIALOG = 4        ' Font dialog
  OPEN_COLOR_DIALOG = 5       ' Color dialog
  OPENFILE_LIST = 6           ' FileOpen with LEFT custom image
End Enum
Public giCommonDialogStyle As COMMON_DIALOG_STYLE

Private Const WM_USER = &H400
Private Const WM_INITDIALOG As Long = &H110
Private Const WM_COMMAND As Long = &H111
Private Const WM_DESTROY As Long = &H2
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_CHOOSEFONT_GETLOGFONT As Long = (WM_USER + 1)
Private Const WM_CHOOSEFONT_SETFLAGS As Long = (WM_USER + 102)
Private Const WM_CHOOSEFONT_SETLOGFONT As Long = (WM_USER + 101)
Private Const WM_CLOSE As Long = &H10
'Private Const WM_GETFONT As Long = &H31
'Private Const WM_SETFONT As Long = &H30
'Private Const WM_PAINT As Long = &HF&


Declare Function SetParent Lib "user32" (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const CLR_INVALID As Long = &HFFFF&
Private Const CLR_NONE As Long = &HFFFFFFFF
Private Const CLR_DEFAULT As Long = &HFF000000

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Rem ----------------------------------------------------------------
Rem  SUBCLASSING PER COMMONDIALOG
Rem ----------------------------------------------------------------
'Public Const GWL_WNDPROC = (-4) 'already in SubClass.bas
Public lpPrevCDWndProc As Long
Public gHW_CD As Long
Public gForm1 As Form
Public gPB As PictureBox
Public gPBTemp As PictureBox
Public gTB As Toolbar
Public gCheckPreview As CheckBox
Public gImage1 As Image
Public gTBox As TextBox
Public gAdattaImmagine As Integer
Public gszPath As String


Rem ----------------------------------------------------------------
Rem      Begin of Common Dialogs callback
Rem ----------------------------------------------------------------
Public Const WM_NOTIFY = &H4E

Type OPENFILENAME2              ' used in OFNOTITY
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As Long 'String
        lpstrCustomFilter As Long 'String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As Long 'String
        nMaxFile As Long
        lpstrFileTitle As Long 'String
        nMaxFileTitle As Long
        lpstrInitialDir As Long 'String
        lpstrTitle As Long 'String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As Long 'String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As Long 'String
End Type

Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME2
        pszFile As Long 'String '  May be NULL
End Type


Rem CommonDialogNotifications
Public Const CDN_FIRST = (-601)
Public Const CDN_LAST = (-699)

Rem Open/Save dialogs
Public Const CDN_INITDONE = (CDN_FIRST - &H0)
Public Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Public Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Public Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
Public Const CDN_HELP = (CDN_FIRST - &H4)
Public Const CDN_FILEOK = (CDN_FIRST - &H5)
Public Const CDN_TYPECHANGE = (CDN_FIRST - &H6)

Public Const CDM_FIRST = (WM_USER + 100)
Public Const CDM_LAST = (WM_USER + 200)

' lParam = pointer to text buffer that gets filled in
' wParam = max number of characters of the text buffer (including NULL)
' return = < 0 if error; number of characters needed (including NULL)
Public Const CDM_GETSPEC = (CDM_FIRST + &H0)

' lParam = pointer to text buffer that gets filled in
' wParam = max number of characters of the text buffer (including NULL)
' return = < 0 if error; number of characters needed (including NULL)
Public Const CDM_GETFILEPATH = (CDM_FIRST + &H1)

' lParam = pointer to text buffer that gets filled in
' wParam = max number of characters of the text buffer (including NULL)
' return = < 0 if error; number of characters needed (including NULL)
Public Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)

' lParam = pointer to ITEMIDLIST buffer that gets filled in
' wParam = size of the ITEMIDLIST buffer
' return = < 0 if error; length of buffer needed
Public Const CDM_GETFOLDERIDLIST = (CDM_FIRST + &H3)

' lParam = pointer to a string
' wParam = ID of control to change
' return = not used
Public Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)

' lParam = not used
' wParam = ID of control to change
' return = not used
Public Const CDM_HIDECONTROL = (CDM_FIRST + &H5)

' lParam = pointer to default extension (no dot)
' wParam = not used
' return = not used
Public Const CDM_SETDEFEXT = (CDM_FIRST + &H6)

Public Const DWL_MSGRESULT = 0

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" ( _
   ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long

Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare Function SendDlgItemMessage Lib "user32.dll" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Const GWL_WNDPROC As Long = -4
Private Const GWL_STYLE As Long = -16

Public Const LVM_GETITEMCOUNT = &H1000 + 4
Private Const CB_GETCURSEL As Long = &H147
Private Const CB_GETITEMDATA As Long = &H150
Private Const CB_ERR As Long = (-1)
Private Const CB_RESETCONTENT As Long = &H14B


'Declare Function VarPtr Lib "VBA5.DLL" (Ptr As Any) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal Lenght As Long)
   
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Public Const SWP_SHOWWINDOW = &H40
'Public Const SWP_NOREDRAW = &H8
'Public Const SWP_FRAMECHANGED = &H20
'Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED

Rem -----------------------------------------------------------------------
Rem              end of Common Dialog callback
Rem -----------------------------------------------------------------------


Rem --------------------------------------------------------
Rem Get any errors during execution of common dialogs
Rem --------------------------------------------------------
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Rem --------------------------------------------------------
Rem File Open/Save structures and declarations
Rem --------------------------------------------------------
Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Rem new
Public Const OFN_EXPLORER = &H80000
Public Const OFN_LONGNAMES = &H200000


Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Rem --------------------------------------------------------


Rem --------------------------------------------------------
Rem ChooseColor structure and function declarations
Rem --------------------------------------------------------
Type CHOOSECOLOR_TYPE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR_TYPE) As Long
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Rem --------------------------------------------------------


Rem --------------------------------------------------------
Rem FONT STUFF
Rem --------------------------------------------------------
Public Const LF_FACESIZE = 32
Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
        'lfFaceName As String * LF_FACESIZE
End Type
Public lpLF As LOGFONT

Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Rem --------------------------------------------------------

Rem --------------------------------------------------------
Rem ChooseFont structure and function declarations
Rem --------------------------------------------------------
Type ChooseFont
    lStructSize As Long
    hwndOwner As Long           '  caller's window handle
    hdc As Long                 '  printer DC/IC or NULL
    lpLogFont As Long           '  ptr. to a LOGFONT struct - changed from old "lpLogFont As LOGFONT"
    iPointSize As Long          '  10 * size in points of selected font
    Flags As Long               '  enum. type flags
    rgbColors As Long           '  returned text color
    lCustData As Long           '  data passed to hook fn.
    lpfnHook As Long            '  ptr. to hook function
    lpTemplateName As String    '  custom template name
    hInstance As Long           '  instance handle of.EXE that contains cust. dlg. template
    lpszStyle As String         '  return the style field here must be LF_FACESIZE or bigger
    nFontType As Integer        '  same value reported to the EnumFonts call back with the extra FONTTYPE bits added
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long            '  minimum pt size allowed &
    nSizeMax As Long            '  max pt size allowed if CF_LIMITSIZE is used
End Type

Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long


Public Const CF_SCREENFONTS = &H1&
Public Const CF_PRINTERFONTS = &H2&
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000&           'Must also have CF_SCREENFONTS and CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H1000&
Public Const CF_SCALABLEONLY = &H2000&
Public Const CF_TTONLY = &H4000&
Public Const CF_NOFACESEL = &H8000&
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

'Public Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1) 'WM_USER + 1

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const FINDMSGSTRING = "commdlg_FindReplace"
Public Const HELPMSGSTRING = "commdlg_help"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2
Rem --------------------------------------------------------


Rem --------------------------------------------------------
Rem Printer related structures and function declarations
Rem --------------------------------------------------------
Type PrintDlg
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hdc As Long
        Flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long
        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type

Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000

Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Public Const DN_DEFAULTPRN = &H1
Rem --------------------------------------------------------

'************** end of Common Dialogs Declares ***********


Rem --------------------------------------------------------
Rem Public MEMORY Stuff
Rem --------------------------------------------------------
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

'Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal Lenght As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function lstrcpyANY Lib "kernel32" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
Rem --------------------------------------------------------


Rem --------------------------------------------------------
Rem PRINTER stuff
Rem --------------------------------------------------------
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type
' --------------------------------------------------------


Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function SelectObject Lib "gdi32.dll" ( _
   ByVal hdc As Long, _
   ByVal hObject As Long) As Long


Rem ===================================================================================================================================================================================
Rem COMDLG32.DLL RESOURCES ! ! !
Rem ===================================================================================================================================================================================

Rem ===================================================================================================================================================================================
Rem CHOOSE COLOR (SELEZIONA COLORE)
Rem ===================================================================================================================================================================================
'CHOOSECOLOR DIALOGEX 2, 0, 323, 190
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_CAPTION | WS_SYSMENU
'Caption "Colore"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "&Colori di base:", -1, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 4, 4, 140, 9
'   CONTROL "", 720, STATIC, SS_SIMPLE | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 4, 14, 140, 86
'   CONTROL "Colori &personalizzati:", -1, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 4, 106, 140, 9
'   CONTROL "", 721, STATIC, SS_SIMPLE | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 4, 116, 140, 28
'   CONTROL "&Definisci colori personalizzati >>", 719, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 4, 150, 140, 14
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 4, 166, 44, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 52, 166, 44, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 100, 166, 44, 14
'   CONTROL "", 710, STATIC, SS_SIMPLE | SS_SUNKEN | WS_CHILD | WS_VISIBLE, 152, 4, 118, 116
'   CONTROL "", 702, STATIC, SS_SIMPLE | SS_SUNKEN | WS_CHILD | WS_VISIBLE, 298, 4, 8, 116
'   CONTROL "", 709, STATIC, SS_SIMPLE | SS_SUNKEN | WS_CHILD | WS_VISIBLE, 152, 124, 41, 21
'   CONTROL "&u", 713, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 300, 200, 4, 14
'   CONTROL "Colore", 730, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 148, 147, 21, 18
'   CONTROL "|Tinta u&nita", 731, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 173, 147, 20, 18
'   CONTROL "&Tonalità:", 723, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 199, 126, 40, 8
'   CONTROL "", 703, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 241, 124, 18, 12
'   CONTROL "&Saturazione:", 724, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 199, 140, 40, 8
'   CONTROL "", 704, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 241, 138, 18, 12
'   CONTROL "&Luminosità:", 725, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 199, 154, 40, 8
'   CONTROL "", 705, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 241, 152, 18, 12
'   CONTROL "&Rosso:", 726, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 263, 126, 23, 9
'   CONTROL "", 706, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 288, 124, 18, 12
'   CONTROL "&Verde:", 727, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 263, 140, 23, 9
'   CONTROL "", 707, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 288, 138, 18, 12
'   CONTROL "&Blu:", 728, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 263, 154, 23, 9
'   CONTROL "", 708, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 288, 152, 18, 12
'   CONTROL "&Aggiungi ai colori personalizzati", 712, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 198, 166, 107, 14
'}
Private Enum enumCOLOR_DLG_CTL
  stc_BaseColor = -1
  stc_Simple1 = 720
  stc_CustomColor = -1
  stc_Simple2 = 721
  btn_DefineColor = 719
  btn_Ok = -1
  BTN_CANCEL = 2
  btn_Help = 1038
  stc_Colors = 710
  stc_Gradient = 702
  stc_Color = 709
  stc_ColorLeft = 730
  stc_ColorRight = 731
  stc_Shade = 731
  edt_Shade = 703
  stc_Saturation = 724
  edt_Saturation = 704
  stc_Brightness = 725
  edt_Brightness = 705
  stc_R = 726         ' red
  edt_R = 706
  stc_G = 727         ' green
  edt_G = 707
  stc_B = 728         ' blue
  edt_B = 708
  btn_AddCustom = 712
  btn_X = 713         ' hide button
End Enum

Rem===================================================================================================================================================================================
Rem PRINT GENERAL (GENERALE STAMPANTE)
Rem===================================================================================================================================================================================
'100 DIALOGEX 0, 0, 292, 198
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU
'Caption "Generale"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Seleziona stampante", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE, 7, 7, 278, 104
'   CONTROL "", 1073, BUTTON, BS_GROUPBOX | WS_CHILD, 7, 118, 278, 73
'   CONTROL "", 1000, LISTBOX, LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | WS_CHILD | WS_BORDER | WS_HSCROLL, 14, 17, 264, 53
'   CONTROL "Stato:", 1004, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 76, 35, 8
'   CONTROL "", 1005, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE, 50, 76, 111, 8
'   CONTROL "Percorso:", 1006, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 87, 35, 8
'   CONTROL "", 1007, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE, 50, 87, 158, 8
'   CONTROL "Commento:", 1008, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 98, 35, 8
'   CONTROL "", 1009, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE, 50, 98, 158, 8
'   CONTROL "Stampa su fi&le", 1002, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 166, 77, 58, 8
'   CONTROL "Pre&ferenze", 1010, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 228, 74, 50, 14
'   CONTROL "Tro&va stampante...", 1003, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 215, 92, 63, 14
'   CONTROL "Seleziona stampante", 1011, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_GROUP, 7, 65, 69, 8
'}


Rem ===================================================================================================================================================================================
Rem PRINT GENERAL 2 (GENERALE STAMPANTE 2)
Rem ===================================================================================================================================================================================
'101 DIALOGEX 0, 0, 292, 218
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU
'Caption "Generale"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Seleziona stampante", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE, 7, 7, 278, 104
'   CONTROL "", 1073, BUTTON, BS_GROUPBOX | WS_CHILD, 7, 118, 278, 93
'   CONTROL "", 1000, LISTBOX, LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | WS_CHILD | WS_BORDER | WS_HSCROLL, 14, 17, 264, 53
'   CONTROL "Stato:", 1004, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 76, 35, 8
'   CONTROL "", 1005, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 50, 76, 114, 8
'   CONTROL "Percorso:", 1006, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 87, 35, 8
'   CONTROL "", 1007, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 53, 86, 152, 8
'   CONTROL "Commento:", 1008, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 98, 35, 8
'   CONTROL "", 1009, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 53, 96, 152, 8
'   CONTROL "Stampa su fi&le", 1002, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 168, 77, 57, 8
'   CONTROL "Pre&ferenze", 1010, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 228, 74, 50, 14
'   CONTROL "Tro&va stampante...", 1003, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 208, 92, 70, 14
'   CONTROL "Seleziona stampante", 1011, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_GROUP, 7, 65, 73, 8
'}


Rem===================================================================================================================================================================================
Rem FONT (CARATTERE)
Rem===================================================================================================================================================================================
'401 DIALOGEX 13, 54, 287, 196
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_CAPTION | WS_SYSMENU
'Caption "Carattere"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Caratt&ere:", 1088, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 7, 40, 9
'   CONTROL "", 1136, COMBOBOX, CBS_SIMPLE | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_SORT | CBS_HASSTRINGS | CBS_DISABLENOSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 7, 16, 98, 76
'   CONTROL "&Stile:", 1089, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 7, 44, 9
'   CONTROL "", 1137, COMBOBOX, CBS_SIMPLE | CBS_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 110, 16, 74, 76
'   CONTROL "&Dimensioni:", 1090, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 189, 7, 40, 8
'   CONTROL "", 1138, COMBOBOX, CBS_SIMPLE | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_SORT | CBS_HASSTRINGS | CBS_DISABLENOSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 190, 16, 36, 76
'   CONTROL "Effetti", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 97, 98, 72
'   CONTROL "&Barrato", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 13, 110, 49, 10
'   CONTROL "S&ottolineato", 1041, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 13, 123, 51, 10
'   CONTROL "&Colore:", 1091, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 13, 136, 30, 9
'   CONTROL "", 1139, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_TABSTOP, 13, 146, 82, 100
'   CONTROL "Esempio", 1073, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 97, 116, 43
'   CONTROL "AaBbYyZz", 1092, STATIC, SS_CENTER | SS_NOPREFIX | WS_CHILD | WS_GROUP, 118, 111, 100, 23
'   CONTROL "", 1093, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE, 7, 172, 219, 20
'   CONTROL "Sc&rittura:", 1094, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 146, 50, 9
'   CONTROL "", 1140, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_TABSTOP, 110, 157, 116, 30
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 16, 45, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 32, 45, 14
'   CONTROL "&Applica", 1026, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 48, 45, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 64, 45, 14
'   CONTROL "Assi:", 1074, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 200, 200, 210
'   CONTROL "", 1168, SCROLLBAR, SBS_HORZ | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 70, 217, 100, 10
'   CONTROL "", 1169, SCROLLBAR, SBS_HORZ | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 70, 237, 100, 10
'   CONTROL "", 1170, SCROLLBAR, SBS_HORZ | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 70, 257, 100, 10
'   CONTROL "", 1171, SCROLLBAR, SBS_HORZ | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 70, 277, 100, 10
'   CONTROL "", 1172, SCROLLBAR, SBS_HORZ | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 70, 297, 100, 10
'   CONTROL "", 1173, SCROLLBAR, SBS_HORZ | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 70, 317, 100, 10
'   CONTROL "", 1098, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 25, 217, 40, 10
'   CONTROL "", 1099, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 25, 237, 40, 10
'   CONTROL "", 1100, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 25, 257, 40, 10
'   CONTROL "", 1101, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 25, 277, 40, 10
'   CONTROL "", 1102, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 25, 297, 40, 10
'   CONTROL "", 1103, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 25, 317, 40, 10
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 180, 217, 20, 10 , 0x00010000
'   CONTROL "", 1153, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 180, 237, 20, 10 , 0x00010000
'   CONTROL "", 1154, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 180, 257, 20, 10 , 0x00010000
'   CONTROL "", 1155, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 180, 277, 20, 10 , 0x00010000
'   CONTROL "", 1156, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 180, 297, 20, 10 , 0x00010000
'   CONTROL "", 1157, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_READONLY | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 180, 317, 20, 10 , 0x00010000
'   CONTROL "", 1105, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 70, 207, 40, 10
'   CONTROL "", 1106, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 70, 227, 40, 10
'   CONTROL "", 1107, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 70, 247, 40, 10
'   CONTROL "", 1108, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 70, 267, 40, 10
'   CONTROL "", 1109, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 70, 287, 40, 10
'   CONTROL "", 1110, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 70, 307, 40, 10
'   CONTROL "", 1112, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 130, 207, 40, 10
'   CONTROL "", 1113, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 130, 227, 40, 10
'   CONTROL "", 1114, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 130, 247, 40, 10
'   CONTROL "", 1115, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 130, 267, 40, 10
'   CONTROL "", 1116, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 130, 287, 40, 10
'   CONTROL "", 1118, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 130, 307, 40, 10
'}

Rem ------------------------------------------------------------
Rem per maggior praticità ho enumerato tutti i controlli della
Rem finestra Carattere
Rem ------------------------------------------------------------
Private Enum enumFONT_CTL
  stc_Font = 1088
  cbo_Font = 1136
  stc_Style = 1089
  cbo_Style = 1137
  stc_Size = 1090
  cbo_Size = 1138
  btn_Effects = 1072  ' groupbox
  btn_Strike = 1040   ' checkbox
  btn_Under = 1041    ' checkbox
  stc_Color = 1091
  cbo_Color = 1139
  btn_Sample = 1073   ' groupbox
  stc_Sample = 1092   ' (AaBbYyZz)
  stc_info = 1093
  stc_Writing = 1094
  cbo_Writing = 1140
  btn_Ok = 1
  BTN_CANCEL = 2
  btn_Apply = 1026
  btn_Help = 1038
  ' Note: 'Axis' is a invisible groupbox with some controls
  btn_Axis = 1074     ' groupbox
  hsb_1 = 1168        ' horizontal scrollbar
  hsb_2 = 1169
  hsb_3 = 1170
  hsb_4 = 1171
  hsb_5 = 1172
  hsb_6 = 1173
  stc_1 = 1098        ' static
  stc_2 = 1099
  stc_3 = 1100
  stc_4 = 1101
  stc_5 = 1102
  stc_6 = 1103
  stc_7 = 1105
  stc_8 = 1106
  stc_9 = 1107
  stc_10 = 1108
  stc_11 = 1109
  stc_12 = 1110
  stc_13 = 1112
  stc_14 = 1113
  stc_15 = 1114
  stc_16 = 1115
  stc_17 = 1116
  stc_18 = 1118
  edt_1 = 1152        ' edit
  edt_2 = 1153
  edt_3 = 1154
  edt_4 = 1155
  edt_5 = 1156
  edt_6 = 1157
End Enum

Rem ===================================================================================================================================================================================
Rem  FILE OPEN (APRI - stile vecchio Windows 3.1 - N.B. le dialog #1536 e #1537 sono uguali)
Rem ===================================================================================================================================================================================
'1536 DIALOGEX 36, 24, 268, 134
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_CAPTION | WS_SYSMENU
'Caption "Apri"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "&Nome file:", 1090, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 6, 6, 76, 9
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | ES_OEMCONVERT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 6, 16, 90, 12
'   CONTROL "", 1120, LISTBOX, LBS_STANDARD | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | LBS_DISABLENOSCROLL | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 6, 32, 90, 68
'   CONTROL "&Cartelle:", -1, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 6, 96, 9
'   CONTROL "", 1088, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 16, 96, 9
'   CONTROL "", 1121, LISTBOX, LBS_STANDARD | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | LBS_DISABLENOSCROLL | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 110, 32, 96, 68
'   CONTROL "&Tipo file:", 1089, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 6, 104, 90, 9
'   CONTROL "", 1136, COMBOBOX, CBS_DROPDOWNLIST | CBS_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 6, 114, 90, 96
'   CONTROL "&Unità:", 1091, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 104, 96, 9
'   CONTROL "", 1137, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_SORT | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 110, 114, 96, 68
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 212, 6, 50, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 212, 24, 50, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 212, 46, 50, 14
'   CONTROL "&Sola lettura", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 212, 68, 50, 12
'}


Rem ===================================================================================================================================================================================
Rem  PRINT (STAMPA)
Rem ===================================================================================================================================================================================
'1538 DIALOGEX 32, 32, 288, 186
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU
'Caption "Stampa"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Stampante", 1075, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 8, 4, 274, 84
'   CONTROL "&Nome:", 1093, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 20, 29, 8
'   CONTROL "", 1139, COMBOBOX, CBS_DROPDOWNLIST | CBS_SORT | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_GROUP | WS_TABSTOP, 56, 18, 148, 152
'   CONTROL "&Proprietà...", 1025, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 212, 17, 60, 14
'   CONTROL "Stato:", 1095, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 36, 27, 10
'   CONTROL "", 1099, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 56, 36, 218, 10
'   CONTROL "Tipo:", 1094, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 48, 26, 10
'   CONTROL "", 1098, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 56, 48, 216, 10
'   CONTROL "Percorso:", 1097, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 60, 30, 10
'   CONTROL "", 1101, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 56, 60, 217, 10
'   CONTROL "Commento:", 1096, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 72, 36, 10
'   CONTROL "", 1100, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 56, 72, 152, 10
'   CONTROL "Stampa su fi&le", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 212, 72, 58, 12
'   CONTROL "Intervallo di stampa", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 8, 92, 144, 64
'   CONTROL "&Tutte", 1056, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 106, 64, 12
'   CONTROL "Pag&ine", 1058, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 16, 122, 38, 12
'   CONTROL "S&elezione", 1057, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 16, 138, 49, 12
'   CONTROL "&da:", 1089, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 55, 124, 17, 8
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_NUMBER | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 74, 122, 26, 12
'   CONTROL "&a:", 1090, STATIC, SS_RIGHT | WS_CHILD | WS_VISIBLE | WS_GROUP, 102, 124, 16, 8
'   CONTROL "", 1153, EDIT, ES_LEFT | ES_NUMBER | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 120, 122, 26, 12
'   CONTROL "Copie", 1073, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 160, 92, 122, 64
'   CONTROL "N&umero di copie:", 1092, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 168, 108, 56, 10
'   CONTROL "", 1154, EDIT, ES_LEFT | ES_NUMBER | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 240, 106, 32, 12
'   CONTROL "", 1086, STATIC, SS_ICON | SS_CENTERIMAGE | WS_CHILD | WS_VISIBLE | WS_GROUP, 162, 124, 76, 24
'   CONTROL "&Fascic.", 1041, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 240, 130, 36, 12
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 181, 164, 48, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 234, 164, 48, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 8, 164, 48, 14
'}
Private Enum enumPRINT_CTL
  btn_Printer = 1075      ' groupbox
  stc_Name = 1093
  cbo_Name = 1139
  btn_Properties = 1025
  stc_State = 1095
  stc_State2 = 1099
  stc_Type = 1094
  stc_Type2 = 1098
  stc_Path = 1097
  stc_Path2 = 1101
  stc_Comment = 1096
  stc_Comment2 = 1100
  btn_PrintToFile = 1040  ' checkbox
  btn_Range = 1072        ' groupbox
  btn_All = 1056          ' radiobutton
  btn_Pages = 1058        ' ""
  btn_Selection = 1057    ' ""
  stc_From = 1089
  edt_From = 1152
  stc_To = 1090
  edt_To = 1153
  btn_Copies = 1073       ' groupbox
  stc_Number = 1092
  edt_Number = 1154
  stc_Icon = 1086
  btn_Collapse = 1041     ' Fascicola
  btn_Ok = 1
  BTN_CANCEL = 2
  btn_Help = 1038
End Enum

Rem ===================================================================================================================================================================================
Rem  PRINTER SETUP (IMPOSTA STAMPANTE)
Rem ===================================================================================================================================================================================
'1539 DIALOGEX 32, 32, 297, 178
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU
'Caption "Imposta stampante"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Stampante", 1075, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 8, 4, 281, 84
'   CONTROL "&Nome:", 1093, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 20, 29, 8
'   CONTROL "", 1136, COMBOBOX, CBS_DROPDOWNLIST | CBS_SORT | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_GROUP | WS_TABSTOP, 54, 18, 152, 152
'   CONTROL "&Proprietà...", 1025, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 218, 17, 60, 14
'   CONTROL "Stato:", 1095, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 36, 35, 10
'   CONTROL "", 1099, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 55, 36, 231, 10
'   CONTROL "Tipo:", 1094, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 48, 35, 10
'   CONTROL "", 1098, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 55, 48, 231, 10
'   CONTROL "Percorso:", 1097, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 60, 35, 10
'   CONTROL "", 1101, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 55, 60, 231, 10
'   CONTROL "Commento:", 1096, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 72, 35, 10
'   CONTROL "", 1100, STATIC, SS_LEFTNOWORDWRAP | SS_NOPREFIX | WS_CHILD | WS_VISIBLE | WS_GROUP, 55, 72, 231, 10
'   CONTROL "Foglio", 1073, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 8, 92, 170, 56
'   CONTROL "&Formato:", 1089, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 108, 32, 8
'   CONTROL "", 1137, COMBOBOX, CBS_DROPDOWNLIST | CBS_SORT | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_GROUP | WS_TABSTOP, 62, 106, 112, 112
'   CONTROL "&Alimentazione:", 1090, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 14, 128, 45, 8
'   CONTROL "", 1138, COMBOBOX, CBS_DROPDOWNLIST | CBS_SORT | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_GROUP | WS_TABSTOP, 62, 126, 112, 112
'   CONTROL "Orientamento", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 189, 92, 100, 56
'   CONTROL "", 1084, STATIC, SS_ICON | WS_CHILD | WS_VISIBLE | WS_GROUP, 196, 112, 25, 20
'   CONTROL "&Verticale", 1056, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 230, 106, 48, 12
'   CONTROL "&Orizzontale", 1057, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 230, 126, 52, 12
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 183, 156, 51, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 238, 156, 51, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 8, 156, 48, 14
'}
Private Enum enumPRINTER_SETUP_CTL
  btn_Printer = 1075      ' groupbox
  stc_Name = 1093
  cbo_Name = 1139
  btn_Properties = 1025
  stc_State = 1095
  stc_State2 = 1099
  stc_Type = 1094
  stc_Type2 = 1098
  stc_Path = 1097
  stc_Path2 = 1101
  stc_Comment = 1096
  stc_Comment2 = 1100
  btn_Sheet = 1073        ' groupbox
  stc_Format = 1073
  cbo_Format = 1137
  stc_Supply = 1090       '
  cbo_Supply = 1138
  btn_Orientation = 1072  ' groupbox
  stc_Icon = 1084
  btn_Portrait = 1056
  btn_Landscape = 1057
  btn_Ok = 1
  BTN_CANCEL = 2
  btn_Help = 1038
End Enum

Rem ===================================================================================================================================================================================
Rem  FIND (TROVA)
Rem ===================================================================================================================================================================================
'1540 DIALOGEX 30, 73, 242, 62
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_CAPTION | WS_SYSMENU
'Caption "Trova"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Tr&ova:", -1, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 4, 8, 23, 8
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 34, 7, 128, 12
'   CONTROL "Solo parole i&ntere", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 4, 26, 69, 12
'   CONTROL "&Maiuscole/minuscole", 1041, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 4, 42, 79, 12
'   CONTROL "Direzione", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 94, 26, 68, 28
'   CONTROL "&Su", 1056, BUTTON, BS_AUTORADIOBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP, 98, 38, 25, 12
'   CONTROL "&Giù", 1057, BUTTON, BS_AUTORADIOBUTTON | WS_CHILD | WS_VISIBLE, 125, 38, 35, 12
'   CONTROL "Trova succ&essivo", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 168, 5, 70, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 168, 23, 70, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 168, 45, 70, 14
'}
Private Enum enumFIND_CTL
  stc_Find = -1
  edt_Find = 1152
  btn_WholeWord = 1040        ' checkbox
  btn_CaseSensitive = 1041
  btn_SearchDirection = 1072  ' groupbox (direzione)
  btn_Up = 1056               ' radiobutton
  btn_Down = 1057             ' radiobutton
  btn_FindNext = 1
  BTN_CANCEL = 2
  btn_Help = 1038
End Enum


Rem ===================================================================================================================================================================================
Rem  REPLACE (SOSTITUISCI)
Rem ===================================================================================================================================================================================
'1541 DIALOGEX 36, 44, 244, 95
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_CAPTION | WS_SYSMENU
'Caption "Sostituisci"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "&Trova:", -1, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 4, 9, 21, 8
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 55, 7, 113, 12
'   CONTROL "&Sostituisci con:", -1, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 4, 26, 48, 8
'   CONTROL "", 1153, EDIT, ES_LEFT | ES_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 55, 24, 113, 12
'   CONTROL "Solo parole i&ntere", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 5, 46, 71, 10
'   CONTROL "&Maiuscole/minuscole", 1041, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 5, 62, 81, 12
'   CONTROL "Trova succ&essivo", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 171, 4, 70, 14
'   CONTROL "S&ostituisci", 1024, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 171, 21, 70, 14
'   CONTROL "Sostituisci t&utto", 1025, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 171, 38, 70, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 171, 55, 70, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 171, 75, 70, 14
'}
Private Enum enumREPLACE_CTL
  stc_Find = -1
  edt_Find = 1152
  stc_ReplaceWith = -1
  edt_ReplaceWith = 1153
  btn_Whole = 1040  ' checkbox
  btn_CaseSensitive ' checbox
  btn_FindNext = 1  ' like OK
  btn_Replace = 1024
  btn_ReplaceAll = 1025
  BTN_CANCEL = 2
  btn_Help = 1038
End Enum

Rem ===================================================================================================================================================================================
Rem  TIPO DI CARATTERE
Rem ===================================================================================================================================================================================
'1543 DIALOGEX 13, 54, 287, 196
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_CAPTION | WS_SYSMENU
'Caption "Tipo di carattere"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Tipo di cara&ttere:", 1088, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 7, 54, 8
'   CONTROL "", 1136, COMBOBOX, CBS_SIMPLE | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_SORT | CBS_HASSTRINGS | CBS_DISABLENOSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 7, 16, 98, 76
'   CONTROL "&Stile:", 1089, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 7, 44, 9
'   CONTROL "", 1137, COMBOBOX, CBS_SIMPLE | CBS_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 110, 16, 74, 76
'   CONTROL "P&unti:", 1090, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 189, 7, 30, 9
'   CONTROL "", 1138, COMBOBOX, CBS_SIMPLE | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_SORT | CBS_HASSTRINGS | CBS_DISABLENOSCROLL | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 190, 16, 36, 76
'   CONTROL "Effetti", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 97, 98, 72
'   CONTROL "&Barrato", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 13, 110, 49, 10
'   CONTROL "S&ottolineato", 1041, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 13, 123, 51, 10
'   CONTROL "&Colore:", 1091, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 13, 136, 30, 9
'   CONTROL "", 1139, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_TABSTOP, 13, 146, 82, 100
'   CONTROL "Esempio", 1073, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 97, 116, 43
'   CONTROL "AaBbYyZz", 1092, STATIC, SS_CENTER | SS_NOPREFIX | WS_CHILD | WS_GROUP, 118, 111, 100, 23
'   CONTROL "", 1093, STATIC, SS_LEFT | SS_NOPREFIX | WS_CHILD | WS_VISIBLE, 7, 172, 219, 20
'   CONTROL "Sc&rittura:", 1094, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 110, 146, 30, 9
'   CONTROL "", 1140, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_AUTOHSCROLL | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_TABSTOP, 110, 157, 116, 30
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 16, 45, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 32, 45, 14
'   CONTROL "&Applica", 1026, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 48, 45, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 231, 64, 45, 14
'}
Private Enum enumFONT_DLG_CTL
  stc_Font = 1088
  cbo_Font = 1136
  stc_Style = 1089
  cbo_Style = 1137
  stc_Size = 1090
  cbo_Size = 1138
  btn_Effects = 1072  ' groupbox
  btn_Strike = 1040   ' checkbox
  btn_Under = 1041    ' checkbox
  stc_Color = 1091
  cbo_Color = 1139
  btn_Sample = 1073   ' groupbox
  stc_Sample = 1092   ' (AaBbYyZz)
  stc_info = 1093
  stc_Writing = 1094  ' language writing
  cbo_Writing = 1140
  btn_Ok = 1
  BTN_CANCEL = 2
  btn_Apply = 1026
  btn_Help = 1038
End Enum


Rem ===================================================================================================================================================================================
Rem  PAGE SETUP  (IMPOSTA PAGINA)
Rem ===================================================================================================================================================================================
'1546 DIALOGEX 32, 32, 240, 240
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU
'Caption "Imposta pagina"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "", 1080, STATIC, SS_WHITERECT | WS_CHILD | WS_VISIBLE | WS_GROUP, 80, 8, 80, 80
'   CONTROL "", 1081, STATIC, SS_GRAYRECT | WS_CHILD | WS_VISIBLE | WS_GROUP, 160, 12, 4, 80
'   CONTROL "", 1082, STATIC, SS_GRAYRECT | WS_CHILD | WS_VISIBLE | WS_GROUP, 84, 88, 80, 4
'   CONTROL "Foglio", 1073, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 8, 96, 224, 56
'   CONTROL "&Formato:", 1089, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 112, 28, 8
'   CONTROL "", 1137, COMBOBOX, CBS_DROPDOWNLIST | CBS_SORT | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_GROUP | WS_TABSTOP, 63, 110, 160, 160
'   CONTROL "&Alimentazione:", 1090, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 16, 132, 45, 8
'   CONTROL "", 1138, COMBOBOX, CBS_DROPDOWNLIST | CBS_SORT | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_GROUP | WS_TABSTOP, 63, 130, 160, 160
'   CONTROL "Orientamento", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 8, 156, 61, 56
'   CONTROL "&Verticale", 1056, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 16, 170, 49, 12
'   CONTROL "&Orizzontale", 1057, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 16, 190, 49, 12
'   CONTROL "Margini", 1075, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 75, 156, 157, 56
'   CONTROL "Si&nistro:", 1102, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 79, 172, 32, 8
'   CONTROL "", 1155, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 114, 170, 39, 12
'   CONTROL "D&estro:", 1103, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 158, 172, 27, 8
'   CONTROL "", 1157, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 188, 170, 39, 12
'   CONTROL "S&uperiore:", 1104, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 79, 192, 32, 8
'   CONTROL "", 1156, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 114, 190, 39, 12
'   CONTROL "&Inferiore:", 1105, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 158, 192, 28, 8
'   CONTROL "", 1158, EDIT, ES_LEFT | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 188, 190, 39, 12
'   CONTROL "OK", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 80, 220, 48, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 132, 220, 48, 14
'   CONTROL "&Stampante...", 1026, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 184, 220, 48, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 8, 220, 48, 14
'}
Private Enum enumPAGESETUP_CTL
  stc_WhiteRect = 1080
  stc_GrayRectRight = 1081
  stc_GrayRectBottom = 1082
  btn_Sheet = 1073          ' groupbox
  stc_Format = 1089
  cbo_Format = 1137
  stc_Supply = 1090       '
  cbo_Supply = 1138
  btn_Orientation = 1072  ' groupbox
  stc_Icon = 1084
  btn_Portrait = 1056
  btn_Landscape = 1057
  btn_Margin = 1075       ' groupbox
  stc_MLeft = 1102
  edt_MLeft = 1155
  stc_MRight = 1103
  edt_MRight = 1157
  stc_MTop = 1104
  edt_MTop = 1156
  stc_MBottom = 1105
  edt_MBottom = 1158
  btn_Ok = 1
  BTN_CANCEL = 2
  btn_Printer = 1026
  btn_Help = 1038
End Enum

Rem ===================================================================================================================================================================================
Rem  FILE APRI - stile normale
Rem ===================================================================================================================================================================================
'1547 DIALOGEX 0, 0, 295, 165
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CLIPCHILDREN | WS_CAPTION | WS_SYSMENU
'Caption "Apri"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Cerca &in:", 1091, STATIC, SS_LEFT | SS_NOTIFY | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 6, 29, 9
'   CONTROL "", 1137, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 41, 3, 141, 300
'   CONTROL "", 1088, STATIC, SS_LEFT | WS_CHILD, 186, 3, 103, 16
'   CONTROL "", 1120, LISTBOX, LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_MULTICOLUMN | WS_CHILD | WS_BORDER | WS_HSCROLL, 4, 20, 287, 85
'   CONTROL "&Nome file:", 1090, STATIC, SS_LEFT | SS_NOTIFY | WS_CHILD | WS_VISIBLE | WS_GROUP, 5, 112, 40, 8
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 49, 111, 193, 12
'   CONTROL "", 1148, "ComboBoxEx32", 0x50210042, 49, 111, 193, 150
'   CONTROL "&Tipo file:", 1089, STATIC, SS_LEFT | SS_NOTIFY | WS_CHILD | WS_VISIBLE | WS_GROUP, 5, 131, 41, 11
'   CONTROL "", 1136, COMBOBOX, CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 49, 129, 193, 100
'   CONTROL "Ap&ri in sola lettura", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 49, 148, 74, 10
'   CONTROL "&Apri", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 246, 110, 45, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 246, 128, 45, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 246, 145, 45, 14
'}
'Private enumFILEOPEN_CTL
'  stc_SearchIn = 1091
'  cbo_SearchIn = 1137
'  'stc_
'End Enum

Rem ===================================================================================================================================================================================
Rem  PAGINE DA STAMPARE
Rem ===================================================================================================================================================================================
'1549 DIALOGEX 0, 0, 280, 75
'STYLE DS_CONTROL | WS_CHILD
'Caption ""
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Pagine da stampare", 1072, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 0, 0, 149, 73
'   CONTROL "&Tutte", 1056, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP, 7, 11, 58, 10
'   CONTROL "S&elezione", 1057, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 7, 24, 49, 10
'   CONTROL "&Pagina corrente", 1058, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 58, 24, 66, 10
'   CONTROL "Pag&ine:", 1059, BUTTON, BS_RADIOBUTTON | WS_CHILD | WS_VISIBLE, 7, 38, 38, 10
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 58, 37, 86, 12
'   CONTROL "Immettere numeri di pagina e/o intervalli di pagine separati da virgole. Esempio: 1,5-12", 1088, STATIC, SS_LEFT | WS_CHILD | WS_DISABLED | WS_GROUP, 6, 52, 139, 18
'   CONTROL "Immettere un unico numero di pagina o un intervallo di pagine. Ad esempio: 5-12", 1089, STATIC, SS_LEFT | WS_CHILD | WS_DISABLED | WS_GROUP, 6, 52, 139, 18
'   CONTROL "", 1073, BUTTON, BS_GROUPBOX | WS_CHILD | WS_VISIBLE | WS_GROUP, 155, 0, 122, 73
'   CONTROL "N&umero di copie:", 1090, STATIC, SS_LEFT | WS_CHILD | WS_VISIBLE | WS_GROUP, 161, 13, 61, 10
'   CONTROL "", 1153, EDIT, ES_LEFT | ES_NUMBER | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | WS_TABSTOP, 230, 11, 25, 12
'   CONTROL "Fas&cic.", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 161, 38, 35, 10
'   CONTROL "", 1084, STATIC, SS_ICON | SS_CENTERIMAGE | WS_CHILD | WS_VISIBLE | WS_GROUP, 200, 33, 76, 24
'}


Rem ===================================================================================================================================================================================
Rem  FILE APRI - Stile con Toolbar laterale
Rem ===================================================================================================================================================================================
'1552 DIALOGEX 0, 0, 370, 237
'STYLE DS_MODALFRAME | DS_CONTEXTHELP | WS_POPUP | WS_VISIBLE | WS_CLIPCHILDREN | WS_CAPTION | WS_SYSMENU
'Caption "Apri"
'LANGUAGE LANG_ITALIAN, SUBLANG_ITALIAN
'Font 8, "MS Shell Dlg"
'{
'   CONTROL "Cerca &in:", 1091, STATIC, SS_LEFT | SS_NOTIFY | WS_CHILD | WS_VISIBLE | WS_GROUP, 4, 7, 57, 8 , 0x00001000
'   CONTROL "", 1137, COMBOBOX, CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 66, 4, 174, 300
'   CONTROL "", 1088, STATIC, SS_LEFT | WS_CHILD, 248, 4, 80, 14
'   CONTROL "", 1184, "ToolbarWindow32", 0x50012B4C, 4, 22, 58, 208 , 0x00000200
'   CONTROL "", 1120, LISTBOX, LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_MULTICOLUMN | WS_CHILD | WS_BORDER | WS_HSCROLL, 66, 22, 300, 156
'   CONTROL "&Nome file:", 1090, STATIC, SS_LEFT | SS_NOTIFY | WS_CHILD | WS_VISIBLE | WS_GROUP, 67, 187, 58, 8
'   CONTROL "", 1152, EDIT, ES_LEFT | ES_AUTOHSCROLL | WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP, 130, 184, 182, 12
'   CONTROL "", 1148, "ComboBoxEx32", 0x50210042, 130, 184, 182, 150
'   CONTROL "&Tipo file:", 1089, STATIC, SS_LEFT | SS_NOTIFY | WS_CHILD | WS_VISIBLE | WS_GROUP, 67, 203, 58, 8
'   CONTROL "", 1136, COMBOBOX, CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP, 130, 201, 182, 100
'   CONTROL "Ap&ri in sola lettura", 1040, BUTTON, BS_AUTOCHECKBOX | WS_CHILD | WS_VISIBLE | WS_TABSTOP, 130, 217, 180, 8
'   CONTROL "&Apri", 1, BUTTON, BS_DEFPUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 316, 184, 50, 14
'   CONTROL "Annulla", 2, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 316, 200, 50, 14
'   CONTROL "&?", 1038, BUTTON, BS_PUSHBUTTON | WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP, 316, 218, 50, 14
'}

Rem --------------------------------------------------------------------------
Rem  IDENTIFICATIVI DEI CONTROLLI NELLE COMMON DIALOG
Rem --------------------------------------------------------------------------
Rem --- Pulsanti --- /'
Public Const IDOK = 1         ' Apri                  1
Public Const IDCANCEL = 2     ' Annulla               2

Public Const ctlFirst = &H400
Public Const ctlLast = &H4FF
Rem --- Pulsante (16 + Help) --- /'
Public Const psh1 = &H400
Public Const psh2 = &H401     '
Public Const psh3 = &H402     '
Public Const psh4 = &H403     '
Public Const pshHelp = &H40E  'Pulsante di HELP       1038
Rem --- CheckBox (16) --- /'
Public Const chx1 = &H410     'Apri in sola lettura   1040
Public Const chx2 = &H411
Public Const chx3 = &H412
Rem --- RadioButton (16)
Public Const rad1 = &H420   ' Tutte
Public Const rad2 = &H421   ' Selezione
Public Const rad3 = &H422   ' Pagina corrente
Public Const rad4 = &H423   ' Pagine

Rem --- Groups (4) --- /'
Public Const grp1 = &H430
Public Const grp2 = &H431
Rem --- Frames (4) ---/'
Public Const frm1 = &H434
Public Const frm2 = &H435
Rem --- rectangles (4) --- /'
Public Const rct1 = &H438
Public Const rct2 = &H439
Rem --- icons (4) --- /'
Public Const ico1 = &H43C
Public Const ico2 = &H43D
Public Const ico3 = &H43E
Public Const ico4 = &H43F
Rem --- static text (32) --- /' sono le Etichette
Public Const stc1 = &H440
Public Const stc2 = &H441   ' Tipo file               1089
Public Const stc3 = &H442   ' Nome File               1090
Public Const stc4 = &H443   ' Cerca in                1091
Public Const stc5 = &H444   '
Public Const stc6 = &H445
Public Const stc7 = &H446
Public Const stc8 = &H447
Public Const stc9 = &H448
Public Const stc10 = &H449
Rem --- ListBox del Drive/Directory corrente --- /'
Public Const lst1 = &H460   ' Listbox directory       1120  (LISTBOX)
Public Const lst2 = &H461
Public Const lst3 = &H462
Rem --- Combo Boxes (16) --- /'
Public Const cmb1 = &H470   ' Tipo file               1136  (COMBOBOX)
Public Const cmb2 = &H471   ' Cerca in                1137  (COMBOBOX)
Public Const cmb3 = &H472   '
Public Const cmb4 = &H47C   ' Nome file               1148  (ComboBoxEx32)
Rem --- Edit controls (16) --- /'
Public Const edt1 = &H480
Public Const edt2 = &H481
Public Const edt3 = &H482
Public Const edt4 = &H483
Public Const scr1 = &H490   'Scroll Bars (8)
Public Const scr2 = &H491
Public Const scr3 = &H492
Rem --- toolbar --- /'
Public Const tlb1 = &H4A0




Rem /* The GetBValue function returns the B-component of an RGB value. */
Public Function GetBValue(ByVal Color As Long) As Integer
  Let GetBValue = (Color \ (&H100 ^ 2) And &HFF)
End Function
Rem /* The GetGValue function returns the G-component of an RGB value. */
Public Function GetGValue(ByVal Color As Long) As Integer
  Let GetGValue = (Color \ (&H100 ^ 1) And &HFF)
End Function

Rem /* The GetRValue function returns the R-component of an RGB value. */
Public Function GetRValue(ByVal Color As Long) As Integer
  Let GetRValue = (Color \ (&H100 ^ 0) And &HFF)
End Function
'---------------------------------------------------------------------------------------
' Procedure   : CDCallBack
' DateTime    : 20/07/2004 10.56
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     :
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Function CDCallBack(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Dim wID As Long
Dim wNotifyCode As Long
Dim hwndCtl As Long
Dim retV As Long, lRet As Long
Dim hdlgParent As Long
Dim rc As RECT, rcDesk As RECT, rL As RECT, rcDE As RECT
Dim pt As POINTAPI
Rem FONT DIALOG
Dim lpLF As LOGFONT
Dim CF As ChooseFont
Dim hFontToUse As Long, sFontName As String, RetValue As Long
Dim tBuf As String * 80
Dim hdc As Long
Dim iIndex As Long, dwRGB As Long

  On Error GoTo CDCallBack_Error

    retV = False
    
    'Debug.Print "CDCallback-> Msg=" & Msg, "wp=" & wp, "lp=" & lp
    Select Case msg
    
        Case WM_NOTIFY  ' FileOpen/FileSave
            retV = CDNotify(hWnd, lp)
        
        Case WM_INITDIALOG
            Rem =====================================================================
            Rem hWnd is related to FONT dialog , NOT to calling form!!!
            Rem =====================================================================
            If giCommonDialogStyle = OPEN_FONT_DIALOG Then
              CustomizeFontDialog hWnd
            End If
            
            
        Rem --------------------------------------------------------------------
        Case WM_COMMAND
            Rem Receive messagge send by control
            Rem (Font & Colors dialog only)
            Rem --------------------------------------------------------------------
            'Debug.Print "msg=" & msg, "wpH=" & HIWORD(wp), "wpL=" & LOWORD(wp), "lp=" & lp, "lpH=" & HIWORD(lp), "lpL=" & LOWORD(lp)
            
            Rem lParam=Handle del controllo
            Rem wParam=ID del controllo (es: 1=OK, 2=Annulla,
            Rem       719=Definisci colori... 712=Aggiungi a colori...
            Rem ------------------------- ESEMPIO -------------------------------
            Rem Select Case wp
            Rem    Case 1  ' OK
            Rem        MsgBox "Pulsante OK"
            Rem    Case 2  ' ANNULLA
            Rem        MsgBox "Pulsante ANNULLA"
            Rem    Case 719
            Rem        MsgBox "Pulsante DEFINISCI COLORI PERSONALIZZATI"
            Rem End Select
            

            If giCommonDialogStyle = OPEN_FONT_DIALOG Then
              Rem Click to Apply update the text on the add TextBox
              Rem ----------------------------------------------------------------------
              If wp = enumFONT_CTL.btn_Apply Then ' Apply pressed
                Rem Obtain info for selected font
                SendMessage hWnd, WM_CHOOSEFONT_GETLOGFONT, 0&, lpLF
                
                hFontToUse = CreateFontIndirect(lpLF)
                hdc = GetDC(hWnd)
                If hFontToUse = 0 Then Exit Function
                SelectObject hdc, hFontToUse
                RetValue = GetTextFace(hdc, 79, tBuf)
                sFontName = Mid$(tBuf, 1, RetValue)
                
                Rem --------------------------------------------------------------------
                Rem Update TextBox
                Rem --------------------------------------------------------------------
                With gTBox
                  Rem ------------------------------------------
                  Rem Retrieve selected color
                  Rem -----------------------------------------
                  iIndex = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETCURSEL, 0&, 0&)    ' cmb4
                  If iIndex <> CB_ERR Then
                    dwRGB = SendDlgItemMessage(hWnd, enumFONT_CTL.cbo_Color, CB_GETITEMDATA, iIndex, 0&)
                    'Debug.Print COLORREF_to_RGB(dwRGB)
                  End If
                  .ForeColor = dwRGB  ' <- DON'T WORK! IT'S A BUG!!!
                  With .Font
                    .Name = sFontName
                    .Size = Abs(lpLF.lfHeight * (72 / GetDeviceCaps(hdc, LOGPIXELSY)))
                    ' .Bold = lpLF.lfWeight > 500
                    .Weight = lpLF.lfWeight
                    .Italic = lpLF.lfItalic
                    .Underline = lpLF.lfUnderline
                    .Strikethrough = lpLF.lfStrikeOut
                  End With
                End With
                ReleaseDC hWnd, hdc
              End If
            End If
                        
         Case WM_DESTROY
            Rem --------------------------------------------------------------------
            Rem Release controls to original form
            Rem --------------------------------------------------------------------
            If giCommonDialogStyle = OPENFILE_PICTURE Then
                SetParent gPB.hWnd, frmControls.hWnd
                SetParent gCheckPreview, frmControls.hWnd
                
            ElseIf giCommonDialogStyle = OPENFILE_AUDIO Then
                SetParent gPB.hWnd, frmControls.hWnd
                SetParent gTB.hWnd, gTB.Tag
                
            ElseIf giCommonDialogStyle = OPEN_FONT_DIALOG Then
                SetParent gTBox.hWnd, frmControls.hWnd
            
            ElseIf giCommonDialogStyle = OPENFILE_LIST Then
                'Set gPB = Nothing
                'Set gPBTemp = Nothing
            End If
            
    End Select
    CDCallBack = retV

  On Error GoTo 0
  Exit Function

CDCallBack_Error:
  Debug.Print "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & "in procedure CDCallBack of Modulo modCMDialog"
  Resume Next
    
End Function

Function CDNotify(ByVal hWnd As Long, ByVal lp As Long) As Long
Rem ------------------------------------------------------------
Rem Form FILE OPEN / FILE SAVE only !
Rem ------------------------------------------------------------

Dim hdlgParent As Long
Dim rc As RECT, rcDesk As RECT, rL As RECT, rcDE As RECT
Dim lpon As OFNOTIFY
Const MAX_PATH = 255
'Dim gszPath As String
Dim hLV As Long
Dim oldParent As Long
Dim hPic As Long
Dim pt As POINTAPI
Dim lRet As Long
Static X As Long, Y As Long, H As Long, W As Long
Dim api As Long, hCaption As Long
Dim hButtonOK As Long
Dim hCtrl As Long, rCtrl As RECT, hToolBar As Long, xPos As Long, yPos As Long
Dim rcTB As RECT
    
    CopyMemory2 lpon, lp&, Len(lpon)
    ' Debug.Print "CDNotify:  code=" & lpon.hdr.code, "id=" & lpon.hdr.idfrom, "hWnd=" & lpon.hdr.hwndFrom    ', "msg=" & msg
    
    Select Case lpon.hdr.code
          
        Case CDN_INITDONE:  ' The CD dialog is being to show
        
            hdlgParent = GetParent(hWnd)    ' Handle e size della
            GetWindowRect hdlgParent, rc    ' Common Dialog
            
            Rem----------------------------------------------------
            Rem if CD is open for Image or Audio file
            Rem----------------------------------------------------
            If giCommonDialogStyle = OPENFILE_PICTURE Then
                Rem Find the ListView (LST1) position to calculate
                Rem where insert the PictureBox
                hLV = GetDlgItem(hdlgParent, lst1)  ' LST1 handle
                GetWindowRect hLV, rL               ' LST1 rectangle
                Rem In <pt> I will set the Left & Top of ListView LST1
                pt.X = rL.Left
                pt.Y = rL.Top
                ScreenToClient hdlgParent, pt
                
                gPB.Height = (rL.Bottom - rL.Top)
                gPB.Width = gPB.Height * 1.2    ' square
                gPB.Left = (pt.X * 3) + (rL.Right - rL.Left)
                gPB.Top = pt.Y
                rc.Right = rL.Right + gPB.Width + (pt.X * 4)
                
                Rem move CheckBox 'chkPreview'
                gCheckPreview.Left = (pt.X * 3) + (rL.Right - rL.Left) + (gPB.Width - gCheckPreview.Width) / 2
                gCheckPreview.Top = gPB.Top + gPB.Height + 15
                
                Rem Use PictureBox e CheckBox in this CD
                SetParent gPB.hWnd, hdlgParent
                SetParent gCheckPreview.hWnd, hdlgParent
                gPB.Visible = True
                gCheckPreview.Visible = True

            ElseIf giCommonDialogStyle = OPENFILE_AUDIO Then
                Const OffSetY = 75
                rc.Bottom = rc.Bottom + OffSetY
                Rem Move buttons down to show image
                CDMoveControl hdlgParent, IDOK, , OffSetY
                CDMoveControl hdlgParent, IDCANCEL, , OffSetY
                
                Rem Left position
                hCtrl = GetDlgItem(hdlgParent, lst1)    ' lst1 handle
                GetWindowRect hCtrl, rCtrl              ' lst1 rectangle
                pt.X = rCtrl.Left
                ScreenToClient hdlgParent, pt
                xPos = pt.X
                
                Rem Obtain width
                W = xPos + (rCtrl.Right - rCtrl.Left) - 5
                
                Rem Obtain top
                hCtrl = GetDlgItem(hdlgParent, cmb2)    ' handle
                GetWindowRect hCtrl, rCtrl              ' rectangle
                pt.Y = rCtrl.Top
                ScreenToClient hdlgParent, pt
                yPos = pt.Y
                
                
                On Error Resume Next
                Rem Load image on TOP
                Set gPB.Picture = LoadPicture("")
                gPB.Move xPos, yPos, W, gPBTemp.Height
                gPB.PaintPicture gPBTemp.Picture, 0, 0, W, gPBTemp.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
                SetParent gPB.hWnd, hdlgParent
                
                Rem Move all controls down to show image
                CDMoveControl hdlgParent, lst1, , OffSetY  ', 100, 100
                CDMoveControl hdlgParent, stc4, , OffSetY
                CDMoveControl hdlgParent, stc3, , OffSetY
                CDMoveControl hdlgParent, stc2, , OffSetY
                CDMoveControl hdlgParent, stc1, , OffSetY
                CDMoveControl hdlgParent, cmb1, , OffSetY   ', 100
                CDMoveControl hdlgParent, cmb2, , OffSetY   ', 100
                CDMoveControl hdlgParent, edt1, , OffSetY   ', 100
                Rem CDMoveControl hdlgParent, IDOK, , OffSetY
                Rem CDMoveControl hdlgParent, IDCANCEL, , OffSetY
                
                Rem move ToolBar32
                hToolBar = CDGetToolBarHandle(hdlgParent)
                GetWindowRect hToolBar, rcTB
                pt.X = rcTB.Left
                pt.Y = rcTB.Top
                ScreenToClient hdlgParent, pt
                api = MoveWindow(hToolBar, pt.X, pt.Y + OffSetY, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True)
                
                Rem Add my toolbar
                gTB.Tag = SetParent(gTB.hWnd, hdlgParent)
                gTB.Move pt.X + 100, pt.Y + OffSetY
                
            ElseIf giCommonDialogStyle = OPENFILE_LIST Then
                Const sp& = 130
                
                Rem ------------------------------------------------------------
                Rem To load the custom image from disk in gPBTemp:
                Rem ------------------------------------------------------------
                '/ On Error Resume Next
                '/ Set gPBTemp.Picture = LoadPicture(<PathFilename>)
                '/ If Err.Number <> 0 then
                '/     ' Error!
                '/     Err.Clear
                '/     Exit Function
                '/ End If
                Rem ------------------------------------------------------------
                
                gPB.Move xPos, yPos, 120, 230
                gPB.PaintPicture gPBTemp.Picture, 0, 0, 120, 230, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
                gPB.Refresh
                SetParent gPB.hWnd, hdlgParent
                
                Rem Move buttons right to show image
                CDMoveControl hdlgParent, IDOK, sp&
                CDMoveControl hdlgParent, IDCANCEL, sp&
                Rem LEFT
                hCtrl = GetDlgItem(hdlgParent, lst1)    ' handle
                GetWindowRect hCtrl, rCtrl              ' rettangolo
                pt.X = rCtrl.Left
                ScreenToClient hdlgParent, pt
                xPos = pt.X
                Rem TOP
                hCtrl = GetDlgItem(hdlgParent, cmb2)    ' handle
                GetWindowRect hCtrl, rCtrl              ' rettangolo
                pt.Y = rCtrl.Top
                ScreenToClient hdlgParent, pt
                yPos = pt.Y
                
                
                Rem Move all controls down to show image
                CDMoveControl hdlgParent, lst1, sp&
                CDMoveControl hdlgParent, stc4, sp&
                CDMoveControl hdlgParent, stc3, sp&
                CDMoveControl hdlgParent, stc2, sp&
                CDMoveControl hdlgParent, stc1, sp&
                CDMoveControl hdlgParent, cmb1, sp&
                CDMoveControl hdlgParent, cmb2, sp&
                CDMoveControl hdlgParent, edt1, sp&
                Rem move ToolBar32
                hToolBar = CDGetToolBarHandle(hdlgParent)
                GetWindowRect hToolBar, rcTB
                pt.X = rcTB.Left
                pt.Y = rcTB.Top
                ScreenToClient hdlgParent, pt
                api = MoveWindow(hToolBar, pt.X + sp&, pt.Y, rcTB.Right - rcTB.Left, rcTB.Bottom - rcTB.Top, True)
                rc.Right = rc.Right + sp&
                
            ElseIf giCommonDialogStyle = OPENFILE_DELETEFILE Then
                Rem CD is open to DELETE a file
                Rem modify text button 'Save' to "Delete"
                hCtrl = GetDlgItem(hdlgParent, IDOK)    ' handle
                SetWindowText& hCtrl, "Delete"
                
            ElseIf giCommonDialogStyle = OPENFILE_DEFAULT Then
               Debug.Print "OPENFILE_DEFAULT"
            End If
            
            
            Rem ----------------------------------------------------
            Rem Screen center
            Rem ----------------------------------------------------
            rcDesk.Left = 0
            rcDesk.Top = 0
            rcDesk.Right = Screen.Width / Screen.TwipsPerPixelX
            rcDesk.Bottom = Screen.Height / Screen.TwipsPerPixelY
            SetWindowPos hdlgParent, 0, _
                         (rcDesk.Right - (rc.Right - rc.Left)) / 2, _
                         (rcDesk.Bottom - (rc.Bottom - rc.Top)) / 2, _
                          rc.Right - rc.Left, _
                          rc.Bottom - rc.Top, _
                          SWP_SHOWWINDOW
            
            
        Case CDN_FILEOK:
            Rem Doubli-Click in filename or Open button pressed
            hdlgParent = GetParent(hWnd)
            gszPath = String$(MAX_PATH, 0)
            SendMessageByString hdlgParent, CDM_GETFILEPATH, MAX_PATH, gszPath
            szTrimNull gszPath
            If giCommonDialogStyle = OPENFILE_DELETEFILE Then
              If MessageBox(hdlgParent, "Sicuri di voler eliminare il file " & vbCrLf & UCase(gszPath) & "?" & vbCrLf & vbCrLf & "(Tranquilli. E' solo un esempio!)", _
                      "Finestra di Eliminazione", vbQuestion Or vbYesNo Or vbCritical) <> vbYes Then
                  MsgBox "Il file NON è stato eliminato!"
                  CDNotify = -1
                  SetWindowLong hWnd, DWL_MSGRESULT, -1
              Else
                  MsgBox "Ora cancello il file!", vbCritical
              End If
            Else
              MsgBox "Hai premuto Apri...", vbInformation
            End If
             
        Case CDN_FOLDERCHANGE:  ' change PATH
            hdlgParent = GetParent(hWnd)
            'gszPath = String$(MAX_PATH, 0)
            'SendMessageByString hdlgParent, CDM_GETFOLDERPATH, MAX_PATH, gszPath
            'szTrimNull gszPath
            'MessageBox hdlgParent, "Nuovo Path: " & gszPath, "Custom Open Dialog", 0
            '
            Dim lpClassName As String
            lpClassName = String$(100, 0)
            
            hLV = GetDlgItem(hdlgParent, lst1)
            GetClassName hLV, lpClassName, Len(lpClassName)
            szTrimNull lpClassName

        Case CDN_SELCHANGE:
            Rem Click in ListBox (file or folder)
            Rem ----------------------------------------------------
            Rem NOTE: this event raise automatically after
            Rem a CDN_FOLDERCHANGE event, also.
            hdlgParent = GetParent(hWnd)
            gszPath = String$(MAX_PATH, 0)
            SendMessageByString hdlgParent, CDM_GETFILEPATH, MAX_PATH, gszPath
            szTrimNull gszPath
            If gszPath = "" Then Exit Function
            
            Screen.MousePointer = vbHourglass
            
            On Error Resume Next
            If giCommonDialogStyle = OPENFILE_PICTURE Then
                Rem Load in picTemp
                Set gPBTemp.Picture = LoadPicture(gszPath)
                
                Set gPB.Picture = LoadPicture("")
                If gCheckPreview Then
                    gPB.PaintPicture gPBTemp.Picture, 0, 0, gPB.ScaleWidth, gPB.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
                Else
                    gPB.PaintPicture gPBTemp.Picture, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight, 0, 0, gPBTemp.ScaleWidth, gPBTemp.ScaleHeight
                End If
                gPB.Refresh
                gPB.ToolTipText = FileLen(gszPath)
            
            ElseIf giCommonDialogStyle = OPENFILE_AUDIO Then
                If Right(UCase(gszPath), 3) = "MID" Or Right(UCase(gszPath), 3) = "WAV" Then
                  gTB.Visible = True
                Else
                  gTB.Visible = False
                End If
            End If
            Screen.MousePointer = vbDefault
            
            
            'hLV = GetDlgItem(hdlgParent, psh1)
            'GetClassName hLV, lpClassName, Len(lpClassName)
            'szTrimNull lpClassName
            'Debug.Print "OK", hLV, lpClassName
            'MessageBox hdlgParent, "Nuovo file selezionato", "Clic su file", 0
            'Debug.Print "OK"
            
        'Case CDN_HELP          ' Help button pressed
        
        Case CDN_TYPECHANGE    ' Files Filter change
            hdlgParent = GetParent(hWnd)
            If giCommonDialogStyle = OPENFILE_AUDIO Then
              Rem When filter change clear the Edit control,
              Rem empty the global variable gszPath and hide
              Rem the toolbar...
              DoEvents
              gszPath = String$(MAX_PATH, 0)
              szTrimNull gszPath
              SetWindowText GetDlgItem(hdlgParent, 1152), ""
              lRet = SendMessage(GetDlgItem(hdlgParent, 1148), CB_RESETCONTENT, 0&, 0&)
            End If
        Case CDN_SHAREVIOLATION  'Error share violation
            MsgBox "Error of shared violation!", vbCritical
    End Select

    Exit Function
    
Salta:
   gPB.ToolTipText = ""

End Function



Function CDGetToolBarHandle(ByVal hDialog As Long)
Dim hTB As Long

    hTB = FindWindowEx(hDialog, 0, "ToolBarWindow32", vbNullString)
    CDGetToolBarHandle = hTB
End Function

Sub CDMoveFontDlgControl(ByVal hCD As Long, ByVal ID As Long, Optional X As Long = -1, Optional Y As Long = -1, Optional CX As Long = -1, Optional CY As Long = -1)
Dim hControl As Long
Dim rcL As RECT
Dim pt As POINTAPI

    Rem Retrieve handle of control to resize
    hControl = GetDlgItem(hCD, ID)
    
    Rem Get control size
    GetWindowRect hControl, rcL
                    
    Rem GetWindowRect return the size relative to screen,
    Rem then convert them to CD dialg
    pt.X = rcL.Left
    pt.Y = rcL.Top
    
    ScreenToClient hCD, pt
    
    Rem Set new size
    X = pt.X
    Y = pt.Y
    CX = rcL.Right - rcL.Left
    CY = rcL.Bottom - rcL.Top
    MoveWindow hControl, X, Y, CX, CY, True

'  Debug.Print X; Y; CX; CY
  
End Sub

Sub CDMoveControl(ByVal hCD As Long, ByVal ID As Long, Optional X As Long = -1, Optional Y As Long = -1, Optional CX As Long = -1, Optional CY As Long = -1)
Dim hControl As Long
Dim rcL As RECT
Dim pt As POINTAPI

    Rem Retrieve handle of control to resize
    hControl = GetDlgItem(hCD, ID)
    
    Rem Get control size
    GetWindowRect hControl, rcL
                    
    Rem GetWindowRect return the size relative to screen,
    Rem then convert them to CD dialg
    Rem If X = -1 then control will moved
    Rem otherwise will resized, too
    pt.X = rcL.Left + IIf(X <> -1, X, 0)
    pt.Y = rcL.Top + IIf(Y <> -1, Y, 0)
    
    ScreenToClient hCD, pt
    
    Rem Set new size
    X = pt.X
    Y = pt.Y
    CX = rcL.Right - rcL.Left + IIf(CX <> -1, CX, 0)
    CY = rcL.Bottom - rcL.Top + IIf(CY <> -1, CY, 0)
    MoveWindow hControl, X, Y, CX, CY, True

End Sub

Sub szTrimNull(st As String)
Dim pos As Long

    pos = InStr(st, vbNullChar)
    If pos > 0 Then
        st = Left$(st, pos - 1)
    End If
    

End Sub

Function pShowColors(myForm As Form, CError As Long, ByVal lInitColor As Long, ByVal Flags As Long) As Long

    pShowColors = 0: CError& = 0

    If lInitColor < 0 Then lInitColor = 0

    Dim C As CHOOSECOLOR_TYPE
    Dim MemHandle As Long, OK As Long
    Dim Address As Long
    Dim wSize As Long
    Dim i As Long
    Dim result As Long

    ReDim ClrArray(15) As Long    ' for 16 custom colors
    wSize = Len(ClrArray(0)) * 16 ' block memory size

    Rem ----------------------------------------------------
    Rem  I prepare a block memory size to keep
    Rem  custom colors
    Rem ----------------------------------------------------
    MemHandle = GlobalAlloc(GHND, wSize)
    If MemHandle = 0 Then
        pShowColors = 1 ' return error code
        Exit Function
    End If

    Address = GlobalLock(MemHandle)
    If Address = 0 Then
        pShowColors = 2 ' return error code
        Exit Function
    End If

    Rem ----------------------------------------------------
    Rem Setall custom colors WHITE
    Rem ----------------------------------------------------
    For i& = 0 To UBound(ClrArray)
        ClrArray(i&) = &HFFFFFF
    Next

    Rem ----------------------------------------------------
    Rem copy custom colors to block memory
    Rem ----------------------------------------------------
    Call CopyMemory(ByVal Address, ClrArray(0), wSize)

    Rem ----------------------------------------------------
    Rem fill CHOOSECOLOR structure to open the
    Rem Colors dialog
    Rem ----------------------------------------------------
    C.lStructSize = Len(C)
    C.hwndOwner = myForm.hWnd
    C.lpCustColors = Address
    C.rgbResult = lInitColor
    C.Flags = Flags& Or CC_ENABLEHOOK Or CC_RGBINIT Or CC_FULLOPEN
    C.lpfnHook = VBGetProcAddress(AddressOf CDCallBack)
    
    result = ChooseColor(C)
    CError = CommDlgExtendedError()

    Rem Cancel button, do nothing
    If result = 0 Then
        pShowColors = 3 '  return error code
        Exit Function
    End If

    Rem ----------------------------------------------------
    Rem copy custom colors
    Rem ----------------------------------------------------
    Call CopyMemory(ClrArray(0), ByVal Address, wSize)
    Rem relelase resource
    OK = GlobalUnlock(MemHandle)
    OK = GlobalFree(MemHandle)

    Rem ----------------------------------------------------
    Rem Return color code selected
    Rem ----------------------------------------------------
    'retChooseColor& = C.rgbResult
    pShowColors = C.rgbResult
    
    Rem ----------------------------------------------------
    Rem This is custom colors (not used here)
    Rem ----------------------------------------------------
    'For i& = 0 To UBound(ClrArray)
    '    Debug.Print "Custom Color"; Str$(i&); ":", Hex$(ClrArray(i&))
    'Next
    

End Function

Function pFileOpen(ByVal myForm As Form, FError&, Filter$, IDir$, Title$, Index%, Flags&, Optional sFileName$) As String

    pFileOpen = 0: FError = 0

    Dim O As OPENFILENAME
    Dim Address As Long
    Dim szFile$, szFilter$, szInitialDir$, szTitle$
    Dim result As Long
    Dim File$, FullPath$

    szFile$ = sFileName & String$(256 - Len(sFileName), 0)
    szFilter$ = Filter$
    szInitialDir$ = IDir$
    szTitle$ = Title$
    
    O.lStructSize = Len(O)
    O.hwndOwner = myForm.hWnd
    O.Flags = Flags&
    O.lpstrFilter = szFilter$ & vbNullChar
    O.nFilterIndex = Index%
    O.lpstrFile = szFile
    O.nMaxFile = Len(szFile$)
    O.lpstrFileTitle = szFile$ & vbNullChar
    O.lpstrInitialDir = szInitialDir$ & vbNullChar
    O.lpstrTitle = szTitle$ & vbNullChar
    O.lpfnHook = VBGetProcAddress(AddressOf CDCallBack)
    
    
    Rem Prepare the controls for CommonDialog
    If giCommonDialogStyle = OPENFILE_PICTURE Then
        Set gForm1 = New frmControls
        Set gPB = frmControls.picPreview                  ' Preview
        Set gPBTemp = frmControls.picTemp
        Set gCheckPreview = frmControls.chkPreview        ' Stretch Y/N
    ElseIf giCommonDialogStyle = OPENFILE_AUDIO Then
        Set gForm1 = New frmControls
        Set gPB = frmControls.picPreview
        Set gPBTemp = frmControls.picTemp
        Set gPBTemp.Picture = frmControls.Image3.Picture  ' Image
        Set gTB = frmControls.Toolbar1
    ElseIf giCommonDialogStyle = OPENFILE_LIST Then
        Set gForm1 = New frmControls
        Set gPB = frmControls.picPreview
        Set gPBTemp = frmControls.picTemp
        Set gPBTemp.Picture = frmControls.Image4.Picture  ' Image
    End If
    
    Rem Open the CD dialog
    result = GetOpenFileName(O)
    FError& = CommDlgExtendedError()
    
    Rem Release resources
    If giCommonDialogStyle = OPENFILE_PICTURE Then
        Set gPB = Nothing
        Set gPBTemp = Nothing
        Set gCheckPreview = Nothing
        Unload frmControls
        Set gForm1 = Nothing
    ElseIf giCommonDialogStyle = OPENFILE_AUDIO Then
        Set gPB = Nothing
        Set gPBTemp = Nothing
        Set gTB = Nothing
        Unload frmControls
        Set gForm1 = Nothing
    ElseIf giCommonDialogStyle = OPENFILE_LIST Then
        Set gPB = Nothing
        Set gPBTemp = Nothing
        Unload frmControls
        Set gForm1 = Nothing
    End If
    
    Rem Cancel button pressed
    If result = 0 Then
        pFileOpen = 3
    End If
    
    If (InStr(O.lpstrFileTitle, Chr$(0)) - 1) = 0 Then
        FullPath$ = Left$(O.lpstrFile, InStr(O.lpstrFile, Chr(0)) - 1)
        File$ = szFile$
    Else
        File$ = Left$(O.lpstrFileTitle, InStr(O.lpstrFileTitle, Chr$(0)) - 1)
        FullPath$ = Left$(O.lpstrFile, O.nFileOffset) & File$
    End If
    
    Rem Retrieve selected file name
    Dim Buffer As String
    Buffer = String(255, 0)
    GetFileTitle FullPath$, Buffer, Len(Buffer)
        
    pFileOpen = FullPath$
    
End Function

Function pFileSave(myForm As Form, FError As Long, Filter As String, IDir As String, FileMask As String, Index As Integer, Title As String, Flags As String, DefExt As String, Optional sFileName As String) As Long
    
    pFileSave = 0: FError = 0

    Dim s As OPENFILENAME
    Dim Address As Long
    
    Dim szFile As String, szFilter As String, szInitialDir As String, szTitle As String, NoTitle As String
    Dim result As Long
    Dim File As String, FullPath As String
    
    NoTitle = FileMask
    szFile = NoTitle + String(256 - Len(NoTitle), 0)
    szFilter = Filter
    szInitialDir = IDir
    szTitle = Title

    s.lStructSize = Len(s)
    s.hwndOwner = myForm.hWnd
    s.Flags = Flags
    s.nFilterIndex = 0
    s.lpstrFile = szFile
    s.nMaxFile = Len(szFile$)
    s.lpstrFileTitle = szFile & vbNullChar
    s.lpstrFilter = szFilter & vbNullChar
    s.lpstrInitialDir = szInitialDir & vbNullChar
    s.lpstrTitle = szTitle & vbNullChar
    s.lpstrDefExt = DefExt
    s.lpfnHook = VBGetProcAddress(AddressOf CDCallBack)
    's.lStructSize = Len(s)

    Rem Open CD dialog
    result = GetSaveFileName(s)
    FError = CommDlgExtendedError()

    Rem Cancel button pressed
    If result = 0 Then
        pFileSave = 3
        Exit Function
    End If

    File$ = Left$(s.lpstrFileTitle, InStr(s.lpstrFileTitle, Chr$(0)) - 1)
    FullPath = Left$(s.lpstrFile, s.nFileOffset) & File$
    

End Function

Public Function pShowFont(myForm As Form) As String

  Dim CF As ChooseFont
  Dim LF As LOGFONT
  Dim TempByteArray() As Byte
  Dim ByteArrayLimit As Long
  Dim OldhDC As Long
  Dim FontToUse As Long
  Dim tBuf As String * 80
  Dim X As Long
  Dim uFlag As Long
    
  'Const CF_LIMITSIZE As Long = &H2000&
  'Const CF_USESTYLE As Long = &H80&
  'Const CF_NOSTYLESEL As Long = &H100000
    
  Dim mRGBResult As Long
  Dim mCancelError As Boolean
  Dim RetValue As Long
  Const FW_BOLD = 700
  Const cdlCFScreenFonts = &H1
  'Const cdlCFWYSIWYG = &H8000
  Dim mhOwner As Long
  Dim mFontName As String
  Dim mItalic As Boolean
  Dim mUnderline As Boolean
  Dim mStrikethru As Boolean
  Dim mFontSize As Long
  Dim mBold As Boolean

  Set gForm1 = New frmControls
  Set gTBox = frmControls.Text1

  mCancelError = True
  mhOwner = myForm.hWnd
  
  Rem Set font attribute for dialog
  mFontName = myForm.Font.Name
  mFontSize = 16
  mBold = True
  
  TempByteArray = StrConv(mFontName & vbNullChar, vbFromUnicode)
  ByteArrayLimit = UBound(TempByteArray)
  With LF
     For X = 0 To ByteArrayLimit
        .lfFaceName(X) = TempByteArray(X)
     Next
    .lfHeight = mFontSize / 72 * GetDeviceCaps(GetDC(mhOwner), LOGPIXELSY)
    .lfItalic = mItalic * -1
    .lfUnderline = mUnderline * -1
    .lfStrikeOut = mStrikethru * -1
    If mBold Then .lfWeight = FW_BOLD
  End With
  With CF
    .lStructSize = Len(CF)
    .hwndOwner = mhOwner
    .hdc = GetDC(mhOwner)
    .lpLogFont = lstrcpyANY(LF, LF)
    If Not uFlag Then
       .Flags = CF_BOTH Or CF_WYSIWYG
    Else
       .Flags = uFlag Or CF_BOTH Or CF_WYSIWYG
    End If
   .Flags = .Flags Or CF_INITTOLOGFONTSTRUCT Or CF_ENABLEHOOK Or CF_EFFECTS Or CF_APPLY
   ' .flags = .flags Or CF_LIMITSIZE
   ' .nSizeMax = 10  'mMax
   ' .nSizeMin = 10  ' mMin
   .rgbColors = mRGBResult
   .lCustData = 0
   .lpfnHook = VBGetProcAddress(AddressOf CDCallBack)
   .lpTemplateName = 0
   .hInstance = 0
   .lpszStyle = 0
   .nFontType = SCREEN_FONTTYPE
   .nSizeMin = 0 '14
   .nSizeMax = 0 '14
   .iPointSize = 14 'mFontSize '* 10
  End With
    
  Rem show Font dialog
  RetValue = ChooseFont(CF)
    
  Unload frmControls
  Set gTBox = Nothing
  Set gForm1 = Nothing
  
  If RetValue = 0 Then
    If mCancelError Then Exit Function 'Err.Raise (RetValue)
  Else
    With LF
      mItalic = .lfItalic * -1
      mUnderline = .lfUnderline * -1
      mStrikethru = .lfStrikeOut * -1
    End With
    With CF
      mFontSize = .iPointSize \ 10
      mRGBResult = .rgbColors
      If .nFontType And BOLD_FONTTYPE Then
        mBold = True
      Else
        mBold = False
      End If
    End With
    
    FontToUse = CreateFontIndirect(LF)
    If FontToUse = 0 Then Exit Function
    OldhDC = SelectObject(CF.hdc, FontToUse)
    RetValue = GetTextFace(CF.hdc, 79, tBuf)
    mFontName = Mid$(tBuf, 1, RetValue)
    pShowFont = mFontName
  End If
   
End Function

Function pPrinter(myForm As Form, pError As Long, Flags As Long, FPage As Integer, TPage As Integer, Min As Integer, Max As Integer, Copies As Integer) As Long
    
    pPrinter = 0: pError = 0


    ' ----------------------------------------------------
    'This is similar to Printer Setup
    ' ----------------------------------------------------
    Dim Address As Long
    Dim P As PrintDlg
    Dim D As DEVMODE
    Dim result As Long, OK As Long
    
    Dim szDriver As String, szDevice As String, szOutPut As String
    Dim l As Long
    ' ----------------------------------------------------


    ' ----------------------------------------------------
    ' Set up structure, then call print dialog.  Exit
    '   function set to 1 if get error at this point
    ' ----------------------------------------------------
    P.lStructSize = Len(P)
    P.hwndOwner = myForm.hWnd
    P.Flags = Flags
    P.nFromPage = FPage
    P.nToPage = TPage
    P.nMinPage = Min
    P.nMaxPage = Max
    P.nCopies = Copies
    
    result = PrintDlg(P)
    pError = CommDlgExtendedError()
    
    If result = 0 Then
        myForm.Print "Pulsante Annulla"
        pPrinter = 1
        Exit Function
    End If
    ' ----------------------------------------------------
    
    
    ' ----------------------------------------------------
    ' Delete the handle
    ' ----------------------------------------------------
    If P.hdc <> 0 Then
        OK = DeleteDC(P.hdc)
    End If
    ' ----------------------------------------------------
    
    ' ----------------------------------------------------
    ' Free the memory
    ' ----------------------------------------------------


    If P.hDevNames = 0 Then
        pPrinter = 2
        Exit Function
    Else
        Dim N As DEVNAMES                       ' this works for old drivers as well
        Address = GlobalLock(P.hDevNames)
        Call CopyMemory(N, ByVal Address, Len(N))
        szDriver = String$(15, 0) 'filename buffer
        szDevice = String$(32, 0) 'device buffer
        szOutPut = String$(80, 0) 'port buffer

        l& = lstrcpy(szDriver, Address + N.wDriverOffset)
        szDriver = Left$(szDriver, InStr(szDriver, Chr$(0)) - 1)

        l& = lstrcpy(szDevice, Address + N.wDeviceOffset)
        szDevice = Left$(szDevice, InStr(szDevice, Chr$(0)) - 1)

        l& = lstrcpy(szOutPut, Address + N.wOutputOffset)
        szOutPut = Left$(szOutPut, InStr(szOutPut, Chr$(0)) - 1)

        myForm.Print szDriver, szDevice, szOutPut

        OK = GlobalUnlock(P.hDevNames)
        OK = GlobalFree(P.hDevNames)
    End If
    ' ----------------------------------------------------
    

    ' ----------------------------------------------------
    ' Lock the address, then make a local copy of the
    '   Public block (hDevMode)
    ' ----------------------------------------------------
    If P.hDevMode = 0 Then 'nothing to lock if hDevMode is NULL
        pPrinter = 3
    Else
        Address = GlobalLock(P.hDevMode)    'hDevMode is returned when the driver supports ExtDeviceMode
        Call CopyMemory(D, ByVal Address, Len(D))
        OK = GlobalUnlock(P.hDevMode)
        OK = GlobalFree(P.hDevMode)
    End If
    
    
    If P.hDevMode <> 0 Then myForm.Print "Printer:", Left$(D.dmDeviceName, InStr(D.dmDeviceName, Chr$(0)) - 1)
    myForm.Print "From Page:", Str$(P.nFromPage)
    myForm.Print "To Page:", Str$(P.nToPage)
    myForm.Print "Copies:", Str$(P.nCopies)

End Function

Public Function CmdError(X As Long) As String
Dim pError As String

    If X = 32765 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Common dialog function failed during initialization (not enough memory?)."
    ElseIf X = 32761 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Common dialog function failed to load a specified string."
    ElseIf X = 32760 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Common dialog function failed to load a specified resource."
    ElseIf X = 32759 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Common dialog function failed to lock a specified resource."
    ElseIf X = 32758 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Common dialog function unable to allocate memory for internal data structures."
    ElseIf X = 32757 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Common dialog function unable to lock memory associated with a handle."
    ElseIf X = 32755 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Cancel was selected."
    ElseIf X = 32752 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Couldn't allocate memory for FileName or Filter."
    ElseIf X = 32751 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The call to WinHelp failed.  Check the Help property values."
    ElseIf X = 28671 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The PD_RETURNDEFAULT flag was set in the Flags member of PRINTDLG data structure, but either hDevMode or hDevNames field were nonzero."
    ElseIf X = 28670 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "Load of the required resources failed."
    ElseIf X = 28669 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The common dialog function failed to parse the strings in the [devices] section of the WIN.INI file."
    ElseIf X = 28668 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The PD_RETURNDEFAULT flag was set in the Flags member of PRINTDLG data structure, but either hDevMode or hDevNames field were nonzero."
    ElseIf X = 28667 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The PRINTDLG function failed to load the specified printer's device driver."
    ElseIf X = 28666 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The printer device-driver failed to initialize a DEVMODE data structure (print driver written for WIN 3.0 or later)."
    ElseIf X = 28665 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The PRINTDLG function failed during initialization."
    ElseIf X = 28664 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "No printer device-drivers were found."
    ElseIf X = 28663 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "A default printer does not exist."
    ElseIf X = 28662 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The data in the DEVMODE and DEVNAMES data structrues describes two different printers."
    ElseIf X = 28661 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The PRINTDLG function failed when it attempted to create an information context."
    ElseIf X = 28660 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The [devices] section of the WIN.INI file does not contain an entry for requested printer."
    ElseIf X = 24574 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "No fonts exist.  Must set internally to CF_BOTH, CF_PRINTERFONTS or CF_SCREENFONTS."
    ElseIf X = 20478 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "An attempt to subclass a listbox failed due to insufficient memory."
    ElseIf X = 20477 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "File name is invalid."
    ElseIf X = 20476 Then
        pError = "#" + LTrim$(Str$(X)) + ",  " + "The buffer at which the member lpstrFile points to is too small."
    Else
        pError = "Unknow Printer Error:  #" & Str(X)
    End If
    
    CmdError = pError
    
End Function


Function pSetup(myForm As Form, pError As Long, Flags As Long) As Long
    
    myForm.Cls
    pSetup = 0: pError = 0

    Dim Address As Long
    Dim P As PrintDlg
    Dim D As DEVMODE

    Dim result As Long, OK As Long
    Dim szDriver As String, szDevice As String, szOutPut As String
    Dim l&
    
    P.lStructSize = Len(P)
    P.hwndOwner = myForm.hWnd
    P.Flags = Flags
    
    result = PrintDlg(P)

    pError = CommDlgExtendedError()
    
    If result = 0 Then
        myForm.Print "Pulsante Annulla"
        pSetup = 1
        Exit Function
    End If
    
    ' ----------------------------------------------------
    ' PrintDlg() returns an hDC, a Public handle to
    '   hDevNames and another to hDevMode.  Delete the
    '   ones we don't need
    ' ----------------------------------------------------
    If P.hdc <> 0 Then
        OK = DeleteDC(P.hdc)
    End If

    If P.hDevNames = 0 Then
        pSetup = 2
        Exit Function
    Else
        Dim N As DEVNAMES                       ' this works for old drivers as well
        Address = GlobalLock(P.hDevNames)
        
        Call CopyMemory(N, ByVal Address, Len(N))
        
        szDriver = String$(15, 0) 'filename buffer
        szDevice = String$(32, 0) 'device buffer
        szOutPut = String$(80, 0) 'port buffer

        l& = lstrcpy(szDriver, Address + N.wDriverOffset)
        szDriver = Left(szDriver, InStr(szDriver, Chr$(0)) - 1)

        l& = lstrcpy(szDevice, Address + N.wDeviceOffset)
        szDevice = Left$(szDevice, InStr(szDevice, Chr$(0)) - 1)

        l& = lstrcpy(szOutPut, Address + N.wOutputOffset)
        szOutPut = Left$(szOutPut, InStr(szOutPut, Chr$(0)) - 1)

        myForm.Print szDriver$, szDevice, szOutPut

        OK = GlobalUnlock(P.hDevNames)
        OK = GlobalFree(P.hDevNames)
    End If
    ' ----------------------------------------------------
    

    ' ----------------------------------------------------
    ' Lock the address, then make a local copy of the
    '   Public block (hDevMode)
    ' ----------------------------------------------------
    If P.hDevMode = 0 Then 'nothing to lock if hDevMode is NULL
        pSetup = 3
    Else
        Address = GlobalLock(P.hDevMode)    'hDevMode is returned when the driver supports ExtDeviceMode
        Call CopyMemory(D, ByVal Address, Len(D))
        OK = GlobalUnlock(P.hDevMode)
        OK = GlobalFree(P.hDevMode)
    End If
    
    
    If P.hDevMode <> 0 Then
        myForm.Print "Printer:", Left$(D.dmDeviceName, InStr(D.dmDeviceName, Chr$(0)) - 1)
        myForm.Print "Orientation:", Str$(D.dmOrientation)
    End If

End Function


Public Function VBGetProcAddress(ByVal lpfn As Long) As Long
    
    VBGetProcAddress = lpfn
    
End Function

Public Sub CDHook()

   lpPrevCDWndProc = SetWindowLong(gHW_CD, GWL_WNDPROC, AddressOf CDWindowProc)

End Sub

Public Sub CDUnHook()
   Dim tmp As Long
   tmp = SetWindowLong(gHW_CD, GWL_WNDPROC, lpPrevCDWndProc)
End Sub




Public Function CDWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim LOWORD As Long
    Dim HIWORD As Long
    Dim wParamLOWORD As Long
    Dim wParamHIWORD As Long

    LOWORD = lParam And &HFFFF&
    HIWORD = lParam \ &H10000 And &HFFFF&
    wParamLOWORD = wParam And &HFFFF&
    wParamHIWORD = wParam \ &H10000 And &HFFFF&

    Select Case uMsg
        Case WM_MOUSEMOVE
        Case WM_LBUTTONDOWN
            MsgBox "Clic"
        
    End Select
    
    
    CDWindowProc = CallWindowProc(lpPrevCDWndProc, hw, _
        uMsg, wParam, lParam)
    
End Function

Public Function LOWORD(Param As Long) As Long

    LOWORD = Param And &HFFFF&
    
End Function

Public Function HIWORD(Param As Long) As Long

    HIWORD = Param \ &H10000 And &HFFFF&
    
End Function

Private Sub CustomizeFontDialog(ByVal hWnd As Long)
Dim wID As Long
Dim hwndCtl As Long
Dim retV As Long
Dim hdlgParent As Long
Dim rc As RECT, rcDesk As RECT, rL As RECT, rc2 As RECT
Dim pt As POINTAPI, pt2 As POINTAPI

Dim hctlStcSample As Long
Dim hctlBtnSample As Long
Dim hctlInfo As Long

    retV = False

    Rem ------------------------------------------------------------
    Rem Change Font dialog size & center to the screen
    Rem ------------------------------------------------------------
    GetWindowRect hWnd, rc   ' dialog rectangle
    rcDesk.Left = 0
    rcDesk.Top = 0
    rcDesk.Right = Screen.Width / Screen.TwipsPerPixelX
    rcDesk.Bottom = Screen.Height / Screen.TwipsPerPixelY
    SetWindowPos hWnd, 0, _
                 (rcDesk.Right - (rc.Right - rc.Left)) / 2, _
                 (rcDesk.Bottom - (rc.Bottom - rc.Top) - 100) / 2, _
                   rc.Right - rc.Left, _
                  rc.Bottom - rc.Top + 100, _
                  SWP_SHOWWINDOW
                  
    Rem Move my textbox to Font dialog and set the text
    Rem with a custom string
    With gTBox
      '.BackColor = vbRed <- DON'T WORK... IT'S A BUG!
      .Tag = SetParent(.hWnd, hWnd)
      .Move 180, 4750, 6100, 1400
      .Text = "ABCDEFGHILMNOPQRSTUVWXYZ" & vbCrLf & "abcdefghilmnopqrstuvwxyz" & vbCrLf & "0123456789"
    End With
    

End Sub

Private Function GetRectangle(ByVal hWnd As Long, ByVal ID As Long) As RECT
Dim hCtl As Long, rc As RECT, lRet As Long
Dim pt1 As POINTAPI, pt2 As POINTAPI

  hCtl = GetDlgItem(hWnd, ID)
  lRet = GetWindowRect(hCtl, rc)

  GetRectangle = rc
  
End Function

Private Function COLORREF_to_RGB(ByVal lColor As Long) As Long
  COLORREF_to_RGB = RGB(GetRValue(lColor), GetGValue(lColor), GetBValue(lColor))
End Function

Public Function StopMPlay32() As Long
  Rem ------------------------------------------------------------
  Rem
  Rem Exercise for you: find how to close
  Rem the MPlay32.exe when started by us...
  Rem
  Rem ------------------------------------------------------------
  
End Function
