VERSION 5.00
Begin VB.UserControl czControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0038281B&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ControlContainer=   -1  'True
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   Begin VB.TextBox txtEmbed 
      Appearance      =   0  'Flat
      BackColor       =   &H0038281B&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrTrack 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "czControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' czUI.ctl - Modern UI Control for VB6
' Version: 0.1
' Single-file, self-contained, GDI+ powered
' Drop multiple instances on a form, set ControlType via Property Window
'
' Usage:
'   1. Project -> Add User Control -> Existing -> czUI.ctl
'   2. Drop czControl instances on your form
'   3. Set ControlType property (czButton, czLabel, czPanel, czTitleBar, etc.)
'   4. Configure properties as needed
'
' GitHub: https://github.com/cyberzilla/czUI
' (c) 2026 - GDI+ Modern UI Framework for VB6
'==============================================================================
Option Explicit

'==============================================================================
' SECTION 1: GDI+ FLAT API DECLARATIONS
'==============================================================================
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As Any, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal Mode As Long) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal graphics As Long, ByVal hint As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal Mode As Long) As Long

' Brushes
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipCreateLineBrushFromRect Lib "gdiplus" (ByRef rect As RECTF, ByVal color1 As Long, ByVal color2 As Long, ByVal Mode As Long, ByVal wrapMode As Long, ByRef lineGradient As Long) As Long

' Pens
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal argb As Long, ByVal Width As Single, ByVal unit As Long, ByRef pen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long

' Paths
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushMode As Long, ByRef path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipAddPathArc Lib "gdiplus" (ByVal path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathLine Lib "gdiplus" (ByVal path As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal path As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As Long

' Primitives
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long

' Text
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, ByRef fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As Long, ByVal unit As Long, ByRef font As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal font As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Long, ByRef StringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As Long) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, ByVal trimming As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As Long, ByVal length As Long, ByVal font As Long, ByRef layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal graphics As Long, ByVal str As Long, ByVal length As Long, ByVal font As Long, ByRef layoutRect As RECTF, ByVal StringFormat As Long, ByRef boundingBox As RECTF, ByRef codepointsFitted As Long, ByRef linesFilled As Long) As Long

'==============================================================================
' SECTION 2: WINDOWS API DECLARATIONS
'==============================================================================
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_NORMAL As Long = &H3
Private Const IMAGE_ICON As Long = 1
Private Const LR_COPYFROMRESOURCE As Long = &H4000
Private Const GWL_STYLE As Long = -16
Private Const WS_SYSMENU As Long = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const GWL_WNDPROC As Long = -4
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Src As Any, ByVal cb As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddr As Long, ByVal dwSize As Long, ByVal flType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddr As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal hImage As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpColorRef As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

' Form rounding APIs
Private Declare Function DwmSetWindowAttribute Lib "dwmapi" (ByVal hWnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As WINRECT) As Long

'==============================================================================
' SECTION 3: TYPES
'==============================================================================
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type RECTF
    X      As Single
    Y      As Single
    Width  As Single
    Height As Single
End Type



Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type WINRECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

'==============================================================================
' SECTION 4: PUBLIC ENUMS
'==============================================================================
Public Enum czControlType
    czTitleBar = 0
    czButton = 1
    czLabel = 2
    czPanel = 3
    czTextBox = 4
    czToggle = 5
    czProgressBar = 6
End Enum

Public Enum czButtonStyle
    bsPrimary = 0
    bsSecondary = 1
    bsDanger = 2
    bsText = 3
End Enum

Public Enum czAlignment
    alLeft = 0
    alCenter = 1
    alRight = 2
End Enum

'==============================================================================
' SECTION 5: PRIVATE CONSTANTS
'==============================================================================
' GDI+ constants
Private Const GP_SMOOTH_ANTIALIAS  As Long = 4
Private Const GP_TEXT_CLEARTYPE    As Long = 5
Private Const GP_PIXELOFFSET_HQ   As Long = 4
Private Const GP_UNIT_PIXEL       As Long = 2
Private Const GP_UNIT_POINT       As Long = 3
Private Const GP_FILL_ALTERNATE   As Long = 0
Private Const GP_FONT_REGULAR     As Long = 0
Private Const GP_FONT_BOLD        As Long = 1
Private Const GP_FONT_ITALIC      As Long = 2
Private Const GP_ALIGN_NEAR       As Long = 0
Private Const GP_ALIGN_CENTER     As Long = 1
Private Const GP_ALIGN_FAR        As Long = 2
Private Const GP_TRIM_ELLIPSIS    As Long = 3
Private Const GP_GRADIENT_VERT    As Long = 1
Private Const GP_WRAP_CLAMP       As Long = 4

' Windows messages
Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Private Const WM_SYSCOMMAND       As Long = &H112
Private Const WM_CLOSE            As Long = &H10
Private Const HTCAPTION           As Long = 2
Private Const SC_MINIMIZE         As Long = &HF020&
Private Const SC_MAXIMIZE         As Long = &HF030&
Private Const SC_RESTORE          As Long = &HF120&

' Theme colors (OLE_COLOR / COLORREF format &H00BBGGRR)
Private Const CLR_FORM_BG         As Long = &H38281B   ' #1B2838
Private Const CLR_PANEL_BG        As Long = &H473424   ' #243447
Private Const CLR_BORDER          As Long = &H58412C   ' #2C4158
Private Const CLR_AMBER           As Long = &H23A6F5   ' #F5A623
Private Const CLR_WHITE           As Long = &HFFFFFF   ' #FFFFFF
Private Const CLR_TEXT_SEC        As Long = &HAA9988   ' #8899AA
Private Const CLR_ACCENT          As Long = &HD99F4A   ' #4A9FD9
Private Const CLR_DANGER          As Long = &H2323E8   ' #E82323
Private Const CLR_GRAY            As Long = &H5A5A5A   ' #5A5A5A
Private Const CLR_PLACEHOLDER     As Long = &H999999   ' #999999

' TitleBar button sizing
Private Const TB_BTN_WIDTH        As Long = 46
Private Const TB_ICON_WIDTH       As Long = 40
Private Const TB_GLYPH_SIZE       As Long = 10

' DWM (Desktop Window Manager) - Windows 11 rounded corners
Private Const DWMWA_WINDOW_CORNER_PREFERENCE As Long = 33
Private Const DWMWCP_ROUND       As Long = 2
Private Const FORM_CORNER_RADIUS  As Long = 10

'==============================================================================
' SECTION 6: EVENTS
'==============================================================================
Public Event Click()
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event CloseClick()
Public Event MinimizeClick()
Public Event MaximizeClick()
Public Event FullScreenClick()
Public Event IconClick()
Public Event ValueChanged(ByVal NewValue As Boolean)
Public Event TextChanged(ByVal NewText As String)
Public Event EditGotFocus()
Public Event EditLostFocus()
Public Event ExpandToggle(ByVal Expanded As Boolean)

'==============================================================================
' SECTION 7: PRIVATE MEMBER VARIABLES
'==============================================================================
' State
Private m_Initialized   As Boolean
Private m_IsHovering     As Boolean
Private m_IsPressed      As Boolean
Private m_HotButton      As Long       ' TitleBar: 0=none, 1=close, 2=max, 3=min, 4=icon
Private m_HasFocus       As Boolean    ' TextBox focus state
Private m_IsPlaceholder  As Boolean    ' TextBox placeholder active
Private m_GdipToken      As Long       ' GDI+ startup token

' Fullscreen state
Private m_IsFullScreen   As Boolean
Private m_SavedLeft      As Long       ' Saved form position (twips)
Private m_SavedTop       As Long
Private m_SavedWidth     As Long
Private m_SavedHeight    As Long       ' Saved form height (twips)
Private m_SavedTBHeight  As Long       ' Saved titlebar height (twips)

' Toggle animation state
Private m_AnimPos        As Single     ' Current thumb position (0.0=OFF, 1.0=ON)
Private m_AnimTarget     As Single     ' Target position
Private m_AnimFrom       As Single     ' Starting position
Private m_AnimStart      As Long       ' GetTickCount at animation start
Private m_Animating      As Boolean    ' Animation in progress
Private Const ANIM_DURATION As Long = 300  ' Animation duration in ms

' Expand panel animation state
Private m_ExpandAnimPos    As Single   ' 0.0=collapsed, 1.0=expanded
Private m_ExpandAnimTarget As Single
Private m_ExpandAnimFrom   As Single
Private m_ExpandAnimStart  As Long
Private m_ExpandAnimating  As Boolean
Private m_NotchHover       As Boolean  ' Notch button hover state
Private m_NeedDrawIcon     As Boolean  ' Flag to draw icon after GDI+ cleanup
Private m_pSubclassThunk   As Long     ' ASM thunk for WM_SYSCOMMAND subclass
Private m_OldWndProc       As Long     ' Original form window procedure
Private m_SubclassHwnd     As Long     ' Subclassed form hWnd

' Notch button constants
Private Const NOTCH_WIDTH  As Long = 20   ' Width of notch button (px)
Private Const NOTCH_HEIGHT As Long = 44   ' Height of notch button (px)
Private Const NOTCH_RADIUS As Long = 10   ' Corner radius of notch

' Properties - Universal
Private m_ControlType    As czControlType
Private m_Caption        As String
Private m_BackColor      As Long       ' OLE_COLOR
Private m_ForeColor      As Long       ' OLE_COLOR
Private m_CornerRadius   As Long
Private m_FontName       As String
Private m_FontSize       As Single
Private m_FontBold       As Boolean
Private m_Enabled        As Boolean
Private m_Alignment      As czAlignment

' Properties - Button
Private m_ButtonStyle    As czButtonStyle
Private m_ButtonColor    As Long

' Properties - TitleBar
Private m_ShowMinButton  As Boolean
Private m_ShowMaxButton  As Boolean
Private m_ShowCloseButton As Boolean
Private m_ShowFullScreenButton As Boolean
Private m_AutoHandleButtons As Boolean

' Properties - Panel
Private m_BorderColor    As Long
Private m_BorderWidth    As Long
Private m_ShowHeader     As Boolean
Private m_HeaderText     As String
Private m_Expandable     As Boolean
Private m_PanelExpanded  As Boolean
Private m_ExpandWidth    As Long

' Properties - Toggle
Private m_Checked        As Boolean
Private m_OnColor        As Long
Private m_OffColor       As Long

' Properties - TextBox
Private m_Text           As String
Private m_PlaceholderText As String
Private m_FocusBorderColor As Long
Private m_PasswordChar   As String

' Properties - ProgressBar
Private m_Progress       As Long
Private m_BarColor       As Long
Private m_ShowPercent    As Boolean

'==============================================================================
' SECTION 8: USERCONTROL LIFECYCLE
'==============================================================================
Private Sub UserControl_Initialize()
    InitGDIPlus
    m_Initialized = True
End Sub

Private Sub UserControl_Terminate()
    m_Initialized = False
    UnsubclassForm
    On Error Resume Next
    tmrTrack.Enabled = False
    On Error GoTo 0
    ShutdownGDIPlus
End Sub

Private Sub UserControl_InitProperties()
    ' Set defaults for new instance
    m_ControlType = czButton
    m_Caption = ""
    m_BackColor = CLR_FORM_BG
    m_ForeColor = CLR_WHITE
    m_CornerRadius = 8
    m_FontName = "Segoe UI"
    m_FontSize = 10
    m_FontBold = False
    m_Enabled = True
    m_Alignment = alLeft
    
    m_ButtonStyle = bsPrimary
    m_ButtonColor = CLR_AMBER
    
    m_ShowMinButton = True
    m_ShowMaxButton = True
    m_ShowCloseButton = True
    m_ShowFullScreenButton = True
    m_AutoHandleButtons = True
    
    m_BorderColor = CLR_BORDER
    m_BorderWidth = 1
    m_ShowHeader = False
    m_HeaderText = ""
    m_Expandable = False
    m_PanelExpanded = False
    m_ExpandWidth = 200
    
    m_Checked = False
    m_OnColor = CLR_AMBER
    m_OffColor = CLR_GRAY
    
    m_Text = ""
    m_PlaceholderText = ""
    m_FocusBorderColor = CLR_ACCENT
    m_PasswordChar = ""
    
    m_Progress = 0
    m_BarColor = CLR_ACCENT
    m_ShowPercent = True
    
    ConfigureForType
    RedrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_ControlType = .ReadProperty("ControlType", czButton)
        m_Caption = .ReadProperty("Caption", "")
        m_BackColor = ReadColorProp(PropBag, "BackColor", CLR_FORM_BG)
        m_ForeColor = ReadColorProp(PropBag, "ForeColor", CLR_WHITE)
        m_CornerRadius = .ReadProperty("CornerRadius", 8)
        m_FontName = .ReadProperty("FontName", "Segoe UI")
        m_FontSize = .ReadProperty("FontSize", 10!)
        m_FontBold = .ReadProperty("FontBold", False)
        m_Enabled = .ReadProperty("Enabled", True)
        m_Alignment = .ReadProperty("Alignment", alLeft)
        
        m_ButtonStyle = .ReadProperty("ButtonStyle", bsPrimary)
        m_ButtonColor = ReadColorProp(PropBag, "ButtonColor", CLR_AMBER)
        
        m_ShowMinButton = .ReadProperty("ShowMinButton", True)
        m_ShowMaxButton = .ReadProperty("ShowMaxButton", True)
        m_ShowCloseButton = .ReadProperty("ShowCloseButton", True)
        m_ShowFullScreenButton = .ReadProperty("ShowFullScreenButton", True)
        m_AutoHandleButtons = .ReadProperty("AutoHandleButtons", True)
        
        m_BorderColor = ReadColorProp(PropBag, "BorderColor", CLR_BORDER)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_ShowHeader = .ReadProperty("ShowHeader", False)
        m_HeaderText = .ReadProperty("HeaderText", "")
        m_Expandable = .ReadProperty("Expandable", False)
        m_PanelExpanded = .ReadProperty("PanelExpanded", False)
        m_ExpandWidth = .ReadProperty("ExpandWidth", 200)
        
        m_Checked = .ReadProperty("Checked", False)
        m_OnColor = ReadColorProp(PropBag, "OnColor", CLR_AMBER)
        m_OffColor = ReadColorProp(PropBag, "OffColor", CLR_GRAY)
        
        m_Text = .ReadProperty("Text", "")
        m_PlaceholderText = .ReadProperty("PlaceholderText", "")
        m_FocusBorderColor = ReadColorProp(PropBag, "FocusBorderColor", CLR_ACCENT)
        m_PasswordChar = .ReadProperty("PasswordChar", "")
        
        m_Progress = .ReadProperty("Progress", 0)
        m_BarColor = ReadColorProp(PropBag, "BarColor", CLR_ACCENT)
        m_ShowPercent = .ReadProperty("ShowPercent", True)
    End With
    
    ' Sync animation position with initial checked state
    If m_Checked Then m_AnimPos = 1! Else m_AnimPos = 0!
    
    ConfigureForType
    RedrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ControlType", m_ControlType, czButton
        .WriteProperty "Caption", m_Caption, ""
        .WriteProperty "BackColor", ColorToHex(m_BackColor), ColorToHex(CLR_FORM_BG)
        .WriteProperty "ForeColor", ColorToHex(m_ForeColor), ColorToHex(CLR_WHITE)
        .WriteProperty "CornerRadius", m_CornerRadius, 8
        .WriteProperty "FontName", m_FontName, "Segoe UI"
        .WriteProperty "FontSize", m_FontSize, 10!
        .WriteProperty "FontBold", m_FontBold, False
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "Alignment", m_Alignment, alLeft
        
        .WriteProperty "ButtonStyle", m_ButtonStyle, bsPrimary
        .WriteProperty "ButtonColor", ColorToHex(m_ButtonColor), ColorToHex(CLR_AMBER)
        
        .WriteProperty "ShowMinButton", m_ShowMinButton, True
        .WriteProperty "ShowMaxButton", m_ShowMaxButton, True
        .WriteProperty "ShowCloseButton", m_ShowCloseButton, True
        .WriteProperty "ShowFullScreenButton", m_ShowFullScreenButton, True
        .WriteProperty "AutoHandleButtons", m_AutoHandleButtons, True
        
        .WriteProperty "BorderColor", ColorToHex(m_BorderColor), ColorToHex(CLR_BORDER)
        .WriteProperty "BorderWidth", m_BorderWidth, 1
        .WriteProperty "ShowHeader", m_ShowHeader, False
        .WriteProperty "HeaderText", m_HeaderText, ""
        .WriteProperty "Expandable", m_Expandable, False
        .WriteProperty "PanelExpanded", m_PanelExpanded, False
        .WriteProperty "ExpandWidth", m_ExpandWidth, 200
        
        .WriteProperty "Checked", m_Checked, False
        .WriteProperty "OnColor", ColorToHex(m_OnColor), ColorToHex(CLR_AMBER)
        .WriteProperty "OffColor", ColorToHex(m_OffColor), ColorToHex(CLR_GRAY)
        
        .WriteProperty "Text", m_Text, ""
        .WriteProperty "PlaceholderText", m_PlaceholderText, ""
        .WriteProperty "FocusBorderColor", ColorToHex(m_FocusBorderColor), ColorToHex(CLR_ACCENT)
        .WriteProperty "PasswordChar", m_PasswordChar, ""
        
        .WriteProperty "Progress", m_Progress, 0
        .WriteProperty "BarColor", ColorToHex(m_BarColor), ColorToHex(CLR_ACCENT)
        .WriteProperty "ShowPercent", m_ShowPercent, True
    End With
End Sub

Private Sub UserControl_Paint()
    RedrawControl
End Sub

Private Sub UserControl_Resize()
    If m_ControlType = czTextBox Then PositionTextBox
    If m_ControlType = czTitleBar Then ApplyFormRoundCorners
    RedrawControl
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    ' Redraw when parent's BackColor changes (for corner transparency)
    If PropertyName = "BackColor" Then RedrawControl
End Sub

Private Sub ApplyFormRoundCorners()
    ' Only at runtime, not design-time
    On Error Resume Next
    If Not Ambient.UserMode Then Exit Sub
    
    Dim hWndForm As Long
    hWndForm = UserControl.Parent.hWnd
    If hWndForm = 0 Then Exit Sub
    
    ' Subclass form for taskbar close / Alt+F4 support
    SubclassFormForClose hWndForm
    
    ' When maximized or fullscreen: no rounding
    If IsZoomed(hWndForm) <> 0 Or m_IsFullScreen Then
        ' Remove region so form fills entire screen
        SetWindowRgn hWndForm, 0&, 1&
        ' Tell DWM no rounding (Windows 11)
        Dim noRound As Long: noRound = 1  ' DWMWCP_DONOTROUND
        DwmSetWindowAttribute hWndForm, DWMWA_WINDOW_CORNER_PREFERENCE, noRound, 4&
        Exit Sub
    End If
    
    ' Try Windows 11 DWM rounded corners first (smooth, native)
    Dim cornerPref As Long: cornerPref = DWMWCP_ROUND
    Dim hr As Long
    hr = DwmSetWindowAttribute(hWndForm, DWMWA_WINDOW_CORNER_PREFERENCE, cornerPref, 4&)
    
    If hr = 0 Then
        ' DWM succeeded (Windows 11) - remove any existing region
        SetWindowRgn hWndForm, 0&, 1&
        Exit Sub
    End If
    
    ' Fallback: SetWindowRgn (Windows 10 and earlier)
    Dim rc As WINRECT
    GetClientRect hWndForm, rc
    Dim w As Long: w = rc.Right - rc.Left
    Dim h As Long: h = rc.Bottom - rc.Top
    
    If w > 0 And h > 0 Then
        Dim hRgn As Long
        hRgn = CreateRoundRectRgn(0, 0, w + 1, h + 1, _
                                   FORM_CORNER_RADIUS, FORM_CORNER_RADIUS)
        SetWindowRgn hWndForm, hRgn, 1&
        ' Windows takes ownership of hRgn - do NOT delete
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' SECTION 9: MOUSE & TIMER HANDLING
'==============================================================================
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If Not m_Enabled Then Exit Sub
    
    Select Case m_ControlType
        Case czButton
            m_IsPressed = True
            RedrawControl
            
        Case czTitleBar
            If m_HotButton = 0 Then
                ' Drag the parent form
                If Ambient.UserMode Then
                    Dim hParent As Long
                    hParent = GetParentHwnd()
                    If hParent <> 0 Then
                        ReleaseCapture
                        SendMessage hParent, WM_NCLBUTTONDOWN, HTCAPTION, 0&
                    End If
                End If
            End If
            
        Case czToggle
            m_Checked = Not m_Checked
            ' Start smooth slide animation
            m_AnimFrom = m_AnimPos
            If m_Checked Then m_AnimTarget = 1! Else m_AnimTarget = 0!
            m_AnimStart = GetTickCount()
            m_Animating = True
            tmrTrack.Interval = 16  ' ~60fps for smooth animation
            tmrTrack.Enabled = True
            RaiseEvent ValueChanged(m_Checked)
            RaiseEvent Click
            
        Case czPanel
            ' Check notch button click
            If m_Expandable Then
                Dim notchCy As Single: notchCy = CSng(UserControl.ScaleHeight) / 2!
                If X <= NOTCH_WIDTH And Y >= notchCy - NOTCH_HEIGHT / 2 And Y <= notchCy + NOTCH_HEIGHT / 2 Then
                    m_PanelExpanded = Not m_PanelExpanded
                    RaiseEvent ExpandToggle(m_PanelExpanded)
                    RedrawControl
                End If
            End If
            RaiseEvent Click
            
        Case czTextBox
            If Not txtEmbed.Visible Then
                txtEmbed.Visible = True
                txtEmbed.Text = ""
                m_IsPlaceholder = False
                txtEmbed.SetFocus
            End If
    End Select
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If Not m_Enabled Then Exit Sub
    
    Select Case m_ControlType
        Case czButton
            If m_IsPressed Then
                m_IsPressed = False
                RedrawControl
                If X >= 0 And Y >= 0 And X < UserControl.ScaleWidth And Y < UserControl.ScaleHeight Then
                    RaiseEvent Click
                End If
            End If
            
        Case czTitleBar
            If m_HotButton > 0 Then
                HandleTitleBarClick m_HotButton
            End If
            
        Case czPanel
            RaiseEvent Click
            
        Case czLabel
            RaiseEvent Click
    End Select
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_IsHovering Then
        m_IsHovering = True
        tmrTrack.Enabled = True
        RaiseEvent MouseEnter
        If m_ControlType = czButton Or m_ControlType = czToggle Then RedrawControl
    End If
    
    ' TitleBar: track button hover
    If m_ControlType = czTitleBar Then
        Dim newHot As Long
        newHot = HitTestTitleBar(X, Y)
        If newHot <> m_HotButton Then
            m_HotButton = newHot
            RedrawControl
        End If
    End If
    
    ' Panel: track notch button hover
    If m_ControlType = czPanel And m_Expandable Then
        Dim nCy As Single: nCy = CSng(UserControl.ScaleHeight) / 2!
        Dim wasHover As Boolean: wasHover = m_NotchHover
        m_NotchHover = (X <= NOTCH_WIDTH And Y >= nCy - NOTCH_HEIGHT / 2 And Y <= nCy + NOTCH_HEIGHT / 2)
        If m_NotchHover <> wasHover Then RedrawControl
    End If
End Sub

Private Sub UserControl_DblClick()
    If m_ControlType = czTitleBar Then
        If m_ShowMaxButton And Not m_IsFullScreen Then
            If Ambient.UserMode And m_AutoHandleButtons Then
                ToggleMaximize
            End If
            RaiseEvent MaximizeClick
        End If
    Else
        RaiseEvent DblClick
    End If
End Sub

Private Sub tmrTrack_Timer()
    ' --- Toggle slide animation (60fps, ease-out cubic) ---
    If m_Animating Then
        Dim elapsed As Long
        elapsed = GetTickCount() - m_AnimStart
        
        Dim progress As Single
        If elapsed >= ANIM_DURATION Then
            progress = 1!
            m_Animating = False
            tmrTrack.Interval = 50  ' Restore normal interval
        Else
            progress = CSng(elapsed) / CSng(ANIM_DURATION)
        End If
        
        ' Ease-in-out cubic: smooth acceleration and deceleration
        ' Like CSS cubic-bezier(0.4, 0, 0.2, 1) — Material Design standard
        Dim eased As Single
        If progress < 0.5! Then
            eased = 4! * progress * progress * progress
        Else
            Dim p As Single: p = -2! * progress + 2!
            eased = 1! - (p * p * p) / 2!
        End If
        
        ' Interpolate position
        m_AnimPos = m_AnimFrom + (m_AnimTarget - m_AnimFrom) * eased
        RedrawControl
    End If
    
    ' --- Fullscreen: auto-show/hide titlebar ---
    If m_IsFullScreen Then
        On Error Resume Next
        Dim screenPt As POINTAPI
        GetCursorPos screenPt
        
        ' Titlebar height in pixels
        Dim tbPx As Long
        tbPx = m_SavedTBHeight \ Screen.TwipsPerPixelY
        
        If screenPt.Y <= 5 Then
            ' Mouse at top edge - show titlebar
            If Not UserControl.Extender.Visible Then
                UserControl.Extender.Visible = True
                RedrawControl
            End If
        ElseIf screenPt.Y > tbPx + 50 Then
            ' Mouse well below titlebar - hide titlebar
            If UserControl.Extender.Visible Then
                UserControl.Extender.Visible = False
            End If
        End If
        On Error GoTo 0
    End If
    
    ' --- Mouse hover tracking ---
    Dim pt As POINTAPI
    GetCursorPos pt
    ScreenToClient UserControl.hWnd, pt
    
    If pt.X < 0 Or pt.Y < 0 Or pt.X >= UserControl.ScaleWidth Or pt.Y >= UserControl.ScaleHeight Then
        m_IsHovering = False
        m_HotButton = 0
        m_IsPressed = False
        If Not m_Animating And Not m_IsFullScreen Then tmrTrack.Enabled = False
        RaiseEvent MouseLeave
        RedrawControl
    End If
End Sub

'==============================================================================
' SECTION 10: TEXTBOX EVENT FORWARDING
'==============================================================================
Private Sub txtEmbed_Change()
    If Not m_IsPlaceholder Then
        m_Text = txtEmbed.Text
        RaiseEvent TextChanged(m_Text)
    End If
End Sub

Private Sub txtEmbed_GotFocus()
    m_HasFocus = True
    If m_IsPlaceholder Then
        txtEmbed.Text = ""
        txtEmbed.ForeColor = TranslateColor(m_ForeColor)
        m_IsPlaceholder = False
    End If
    RedrawControl
    RaiseEvent EditGotFocus
End Sub

Private Sub txtEmbed_LostFocus()
    m_HasFocus = False
    m_Text = txtEmbed.Text
    
    ' Show placeholder if empty
    If Len(m_Text) = 0 And Len(m_PlaceholderText) > 0 Then
        txtEmbed.ForeColor = TranslateColor(CLR_PLACEHOLDER)
        txtEmbed.Text = m_PlaceholderText
        m_IsPlaceholder = True
    End If
    
    RedrawControl
    RaiseEvent EditLostFocus
End Sub

'==============================================================================
' SECTION 11: GDI+ LIFECYCLE
'==============================================================================

Private Sub InitGDIPlus()
    Dim gsi As GdiplusStartupInput
    gsi.GdiplusVersion = 1
    GdiplusStartup m_GdipToken, gsi, 0&
End Sub

Private Sub ShutdownGDIPlus()
    ' Intentionally empty - never call GdiplusShutdown.
    ' GDI+ cannot be re-initialized after shutdown in VB6 IDE.
    ' The OS cleans up when the process exits.
End Sub

'==============================================================================
' SECTION 12: COLOR HELPER FUNCTIONS
'==============================================================================
Private Function ARGB(ByVal A As Long, ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
    ' Build GDI+ ARGB color from components (handles signed Long overflow)
    Dim lColor As Long
    lColor = B Or (G * &H100&) Or (R * &H10000)
    If A >= &H80 Then
        ARGB = lColor Or ((A And &H7F&) * &H1000000) Or &H80000000
    Else
        ARGB = lColor Or (A * &H1000000)
    End If
End Function

Private Function ColorToARGB(ByVal oleClr As Long, Optional ByVal Alpha As Long = 255) As Long
    ' Convert OLE_COLOR (COLORREF &H00BBGGRR) to GDI+ ARGB (&HAARRGGBB)
    Dim rgbClr As Long
    OleTranslateColor oleClr, 0, rgbClr
    Dim R As Long, G As Long, B As Long
    R = rgbClr And &HFF&
    G = (rgbClr And &HFF00&) \ &H100&
    B = (rgbClr And &HFF0000) \ &H10000
    ColorToARGB = ARGB(Alpha, R, G, B)
End Function

Private Function TranslateColor(ByVal oleClr As Long) As Long
    ' Convert OLE_COLOR to COLORREF (for VB6 controls)
    Dim rgbClr As Long
    OleTranslateColor oleClr, 0, rgbClr
    TranslateColor = rgbClr
End Function

' Convert VB6 Long (BGR) to web hex format "#RRGGBB"
Private Function ColorToHex(ByVal clr As Long) As String
    Dim R As Long: R = clr And &HFF&
    Dim G As Long: G = (clr \ &H100&) And &HFF&
    Dim B As Long: B = (clr \ &H10000) And &HFF&
    ColorToHex = "#" & Right$("0" & Hex$(R), 2) & Right$("0" & Hex$(G), 2) & Right$("0" & Hex$(B), 2)
End Function

' Convert web hex "#RRGGBB" or "#AARRGGBB" or VB6 "&HBBGGRR" to Long (BGR)
Private Function HexToColor(ByVal sColor As String) As Long
    sColor = Trim$(sColor)
    If Left$(sColor, 1) = "#" Then
        sColor = Mid$(sColor, 2)
        If Len(sColor) = 8 Then
            ' #AARRGGBB - ignore alpha, take RGB
            sColor = Mid$(sColor, 3)
        End If
        If Len(sColor) <> 6 Then Exit Function
        Dim R As Long: R = Val("&H" & Mid$(sColor, 1, 2))
        Dim G As Long: G = Val("&H" & Mid$(sColor, 3, 2))
        Dim B As Long: B = Val("&H" & Mid$(sColor, 5, 2))
        HexToColor = RGB(R, G, B)
    ElseIf Left$(sColor, 2) = "&H" Or Left$(sColor, 2) = "&h" Then
        HexToColor = Val(sColor)
    Else
        HexToColor = Val(sColor)
    End If
End Function

' Read color property that may be stored as Long (old) or String hex (new)
Private Function ReadColorProp(PropBag As PropertyBag, ByVal Name As String, ByVal Default As Long) As Long
    Dim v As Variant
    v = PropBag.ReadProperty(Name, Default)
    If VarType(v) = vbString Then
        ReadColorProp = HexToColor(CStr(v))
    Else
        ReadColorProp = CLng(v)
    End If
End Function

Private Function LightenColor(ByVal clrARGB As Long, ByVal Percent As Long) As Long
    ' Lighten an ARGB color by percentage
    Dim A As Long, R As Long, G As Long, B As Long
    SplitARGB clrARGB, A, R, G, B
    R = R + (255 - R) * Percent \ 100: If R > 255 Then R = 255
    G = G + (255 - G) * Percent \ 100: If G > 255 Then G = 255
    B = B + (255 - B) * Percent \ 100: If B > 255 Then B = 255
    LightenColor = ARGB(A, R, G, B)
End Function

Private Function DarkenColor(ByVal clrARGB As Long, ByVal Percent As Long) As Long
    ' Darken an ARGB color by percentage
    Dim A As Long, R As Long, G As Long, B As Long
    SplitARGB clrARGB, A, R, G, B
    R = R - R * Percent \ 100: If R < 0 Then R = 0
    G = G - G * Percent \ 100: If G < 0 Then G = 0
    B = B - B * Percent \ 100: If B < 0 Then B = 0
    DarkenColor = ARGB(A, R, G, B)
End Function

Private Sub SplitARGB(ByVal clr As Long, ByRef A As Long, ByRef R As Long, ByRef G As Long, ByRef B As Long)
    ' Extract ARGB components from GDI+ color
    B = clr And &HFF&
    G = (clr And &HFF00&) \ &H100&
    R = (clr And &HFF0000) \ &H10000
    If clr And &H80000000 Then
        A = ((clr And &H7F000000) \ &H1000000) Or &H80
    Else
        A = (clr And &H7F000000) \ &H1000000
    End If
End Sub

'==============================================================================
' SECTION 13: DRAWING HELPER FUNCTIONS
'==============================================================================
Private Function CreateRoundRectPath(ByVal X As Single, ByVal Y As Single, _
                                     ByVal w As Single, ByVal h As Single, _
                                     ByVal R As Single) As Long
    ' Create a GDI+ GraphicsPath for a rounded rectangle
    Dim hPath As Long
    GdipCreatePath GP_FILL_ALTERNATE, hPath
    
    If R <= 0 Then
        ' Plain rectangle
        GdipAddPathLine hPath, X, Y, X + w, Y
        GdipAddPathLine hPath, X + w, Y, X + w, Y + h
        GdipAddPathLine hPath, X + w, Y + h, X, Y + h
        GdipClosePathFigure hPath
    Else
        ' Clamp radius
        If R > w / 2! Then R = w / 2!
        If R > h / 2! Then R = h / 2!
        Dim d As Single: d = R * 2!
        
        ' Four corner arcs (GDI+ auto-connects with lines)
        GdipAddPathArc hPath, X, Y, d, d, 180!, 90!             ' Top-left
        GdipAddPathArc hPath, X + w - d, Y, d, d, 270!, 90!     ' Top-right
        GdipAddPathArc hPath, X + w - d, Y + h - d, d, d, 0!, 90! ' Bottom-right
        GdipAddPathArc hPath, X, Y + h - d, d, d, 90!, 90!      ' Bottom-left
        GdipClosePathFigure hPath
    End If
    
    CreateRoundRectPath = hPath
End Function

Private Sub FillRoundRect(ByVal hG As Long, ByVal X As Single, ByVal Y As Single, _
                          ByVal w As Single, ByVal h As Single, _
                          ByVal R As Single, ByVal clr As Long)
    ' Fill a rounded rectangle with solid color
    Dim hPath As Long, hBrush As Long
    hPath = CreateRoundRectPath(X, Y, w, h, R)
    GdipCreateSolidFill clr, hBrush
    GdipFillPath hG, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
End Sub

Private Sub DrawRoundRectBorder(ByVal hG As Long, ByVal X As Single, ByVal Y As Single, _
                                ByVal w As Single, ByVal h As Single, _
                                ByVal R As Single, ByVal clr As Long, _
                                Optional ByVal penWidth As Single = 1!)
    ' Draw rounded rectangle outline
    Dim hPath As Long, hPen As Long
    Dim inset As Single: inset = penWidth / 2!
    hPath = CreateRoundRectPath(X + inset, Y + inset, w - penWidth, h - penWidth, R)
    GdipCreatePen1 clr, penWidth, GP_UNIT_PIXEL, hPen
    GdipDrawPath hG, hPen, hPath
    GdipDeletePen hPen
    GdipDeletePath hPath
End Sub

Private Sub FillGradientRoundRect(ByVal hG As Long, ByVal X As Single, ByVal Y As Single, _
                                  ByVal w As Single, ByVal h As Single, _
                                  ByVal R As Single, ByVal clr1 As Long, ByVal clr2 As Long)
    ' Fill a rounded rectangle with vertical gradient
    Dim hPath As Long, hBrush As Long
    Dim rc As RECTF
    hPath = CreateRoundRectPath(X, Y, w, h, R)
    rc.X = X: rc.Y = Y: rc.Width = w: rc.Height = h + 1!  ' +1 avoids bottom edge artifact
    GdipCreateLineBrushFromRect rc, clr1, clr2, GP_GRADIENT_VERT, GP_WRAP_CLAMP, hBrush
    GdipFillPath hG, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
End Sub

Private Sub DrawText(ByVal hG As Long, ByVal sText As String, _
                     ByVal X As Single, ByVal Y As Single, _
                     ByVal w As Single, ByVal h As Single, _
                     ByVal sFontName As String, ByVal fSize As Single, _
                     ByVal fStyle As Long, ByVal clr As Long, _
                     ByVal hAlign As Long, ByVal vAlign As Long, _
                     Optional ByVal trimming As Long = -1)
    ' Draw anti-aliased text using GDI+
    If Len(sText) = 0 Then Exit Sub
    
    Dim hFamily As Long, hFont As Long, hFormat As Long, hBrush As Long
    Dim rc As RECTF
    
    ' Create font family (with fallback)
    GdipCreateFontFamilyFromName StrPtr(sFontName), 0, hFamily
    If hFamily = 0 Then GdipCreateFontFamilyFromName StrPtr("Arial"), 0, hFamily
    If hFamily = 0 Then Exit Sub
    
    ' Create font, format, brush
    GdipCreateFont hFamily, fSize, fStyle, GP_UNIT_POINT, hFont
    GdipCreateStringFormat 0, 0, hFormat
    GdipSetStringFormatAlign hFormat, hAlign
    GdipSetStringFormatLineAlign hFormat, vAlign
    If trimming >= 0 Then GdipSetStringFormatTrimming hFormat, trimming
    GdipCreateSolidFill clr, hBrush
    
    ' Draw
    rc.X = X: rc.Y = Y: rc.Width = w: rc.Height = h
    GdipDrawString hG, StrPtr(sText), -1, hFont, rc, hFormat, hBrush
    
    ' Cleanup
    If hBrush <> 0 Then GdipDeleteBrush hBrush
    If hFormat <> 0 Then GdipDeleteStringFormat hFormat
    If hFont <> 0 Then GdipDeleteFont hFont
    If hFamily <> 0 Then GdipDeleteFontFamily hFamily
End Sub

'==============================================================================
' SECTION 14: TITLEBAR HELPERS
'==============================================================================
Private Function HitTestTitleBar(ByVal X As Single, ByVal Y As Single) As Long
    ' Returns: 0=none, 1=close, 5=fullscreen, 2=max, 3=min, 4=icon
    Dim w As Long: w = UserControl.ScaleWidth
    HitTestTitleBar = 0
    
    Dim btnRight As Long: btnRight = w
    
    ' Close button (rightmost)
    If m_ShowCloseButton Then
        If X >= btnRight - TB_BTN_WIDTH Then
            HitTestTitleBar = 1: Exit Function
        End If
        btnRight = btnRight - TB_BTN_WIDTH
    End If
    
    ' Fullscreen button
    If m_ShowFullScreenButton Then
        If X >= btnRight - TB_BTN_WIDTH And X < btnRight Then
            HitTestTitleBar = 5: Exit Function
        End If
        btnRight = btnRight - TB_BTN_WIDTH
    End If
    
    ' Maximize button (hidden in fullscreen)
    If m_ShowMaxButton And Not m_IsFullScreen Then
        If X >= btnRight - TB_BTN_WIDTH And X < btnRight Then
            HitTestTitleBar = 2: Exit Function
        End If
        btnRight = btnRight - TB_BTN_WIDTH
    End If
    
    ' Minimize button
    If m_ShowMinButton Then
        If X >= btnRight - TB_BTN_WIDTH And X < btnRight Then
            HitTestTitleBar = 3: Exit Function
        End If
    End If
    
    ' Icon area
    If X <= TB_ICON_WIDTH Then
        HitTestTitleBar = 4: Exit Function
    End If
End Function

Private Sub HandleTitleBarClick(ByVal btn As Long)
    Select Case btn
        Case 1 ' Close
            RaiseEvent CloseClick
            If m_AutoHandleButtons And Ambient.UserMode Then
                Dim hP As Long: hP = GetParentHwnd()
                If hP <> 0 Then SendMessage hP, WM_CLOSE, 0&, 0&
            End If
        Case 2 ' Maximize/Restore
            RaiseEvent MaximizeClick
            If m_AutoHandleButtons And Ambient.UserMode Then
                ToggleMaximize
            End If
        Case 3 ' Minimize
            RaiseEvent MinimizeClick
            If m_AutoHandleButtons And Ambient.UserMode Then
                Dim hP2 As Long: hP2 = GetParentHwnd()
                If hP2 <> 0 Then SendMessage hP2, WM_SYSCOMMAND, SC_MINIMIZE, 0&
            End If
        Case 4 ' Icon
            RaiseEvent IconClick
        Case 5 ' FullScreen
            RaiseEvent FullScreenClick
            If m_AutoHandleButtons And Ambient.UserMode Then
                ToggleFullScreen
            End If
    End Select
End Sub

Private Sub ToggleMaximize()
    Dim hP As Long: hP = GetParentHwnd()
    If hP = 0 Then Exit Sub
    
    If IsZoomed(hP) <> 0 Then
        ' Restore from maximized
        SendMessage hP, WM_SYSCOMMAND, SC_RESTORE, 0&
        ' Re-apply rounded corners
        ApplyFormRoundCorners
    Else
        ' Maximize - remove rounded corners first
        SetWindowRgn hP, 0&, 1&
        SendMessage hP, WM_SYSCOMMAND, SC_MAXIMIZE, 0&
    End If
    
    ' Redraw to update maximize/restore glyph
    RedrawControl
End Sub

Private Sub ToggleFullScreen()
    On Error Resume Next
    Dim hP As Long: hP = GetParentHwnd()
    If hP = 0 Then Exit Sub
    
    If m_IsFullScreen Then
        ' Exit fullscreen
        m_IsFullScreen = False
        
        ' Show titlebar
        UserControl.Extender.Visible = True
        
        ' Restore form position
        UserControl.Parent.Move m_SavedLeft, m_SavedTop, m_SavedWidth, m_SavedHeight
        ApplyFormRoundCorners
        tmrTrack.Interval = 50
    Else
        ' Enter fullscreen
        m_IsFullScreen = True
        
        ' Save titlebar height
        m_SavedTBHeight = UserControl.Extender.Height
        
        ' Save form position
        With UserControl.Parent
            m_SavedLeft = .Left
            m_SavedTop = .Top
            m_SavedWidth = .Width
            m_SavedHeight = .Height
        End With
        
        ' Hide titlebar
        UserControl.Extender.Visible = False
        
        ' Cover entire screen including taskbar
        UserControl.Parent.Move 0, 0, Screen.Width, Screen.Height
        
        ' Remove rounded corners
        SetWindowRgn hP, 0&, 1&
        Dim noRound As Long: noRound = 1
        DwmSetWindowAttribute hP, DWMWA_WINDOW_CORNER_PREFERENCE, noRound, 4&
        
        ' Enable timer for auto-show titlebar on mouse top
        tmrTrack.Interval = 16
        tmrTrack.Enabled = True
    End If
    
    RedrawControl
    On Error GoTo 0
End Sub

Private Function GetParentHwnd() As Long
    On Error Resume Next
    If Ambient.UserMode Then
        GetParentHwnd = UserControl.Parent.hWnd
    End If
    On Error GoTo 0
End Function

'==============================================================================
' SECTION 15: CONTROL CONFIGURATION
'==============================================================================
'==============================================================================
' SECTION: Subclass form for WM_SYSCOMMAND SC_CLOSE (ASM thunk)
' Enables taskbar close / thumbnail X / Alt+F4 on borderless forms
'==============================================================================
Private Sub SubclassFormForClose(ByVal hWnd As Long)
    If m_pSubclassThunk <> 0 Then Exit Sub  ' Already subclassed
    
    ' Add WS_SYSMENU | WS_MINIMIZEBOX to window style
    Dim style As Long
    style = GetWindowLong(hWnd, GWL_STYLE) Or WS_SYSMENU Or WS_MINIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, style
    SetWindowPos hWnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    
    ' Get original wndproc
    m_OldWndProc = GetWindowLong(hWnd, GWL_WNDPROC)
    m_SubclassHwnd = hWnd
    
    ' Get PostMessageA function address
    Dim hU32 As Long: hU32 = GetModuleHandle("user32")
    Dim pPost As Long: pPost = GetProcAddress(hU32, "PostMessageA")
    If pPost = 0 Then Exit Sub
    
    ' Allocate executable memory for thunk
    m_pSubclassThunk = VirtualAlloc(0&, 64&, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If m_pSubclassThunk = 0 Then Exit Sub
    
    ' Build x86 machine code:
    ' if (msg == WM_SYSCOMMAND && (wParam & 0xFFF0) == SC_CLOSE)
    '   { PostMessage(hwnd, WM_CLOSE, 0, 0); return 0; }
    ' else { jmp oldProc; }
    Dim code(63) As Byte
    
    ' 00: cmp dword [esp+8], 0x112  (msg == WM_SYSCOMMAND?)
    code(0) = &H81: code(1) = &H7C: code(2) = &H24: code(3) = &H8
    code(4) = &H12: code(5) = &H1: code(6) = &H0: code(7) = &H0
    ' 08: jne callOriginal (+0x25 -> offset 0x2F)
    code(8) = &H75: code(9) = &H25
    ' 0A: mov eax, [esp+12]  (wParam)
    code(&HA) = &H8B: code(&HB) = &H44: code(&HC) = &H24: code(&HD) = &HC
    ' 0E: and eax, 0xFFF0
    code(&HE) = &H25: code(&HF) = &HF0: code(&H10) = &HFF: code(&H11) = &H0: code(&H12) = &H0
    ' 13: cmp eax, 0xF060  (SC_CLOSE?)
    code(&H13) = &H3D: code(&H14) = &H60: code(&H15) = &HF0: code(&H16) = &H0: code(&H17) = &H0
    ' 18: jne callOriginal (+0x15 -> offset 0x2F)
    code(&H18) = &H75: code(&H19) = &H15
    ' 1A: push 0 (lParam)
    code(&H1A) = &H6A: code(&H1B) = &H0
    ' 1C: push 0 (wParam)
    code(&H1C) = &H6A: code(&H1D) = &H0
    ' 1E: push 0x10 (WM_CLOSE)
    code(&H1E) = &H6A: code(&H1F) = &H10
    ' 20: push [esp+16] (hwnd, +12 for 3 pushes + 4 original)
    code(&H20) = &HFF: code(&H21) = &H74: code(&H22) = &H24: code(&H23) = &H10
    ' 24: call [pThunk+0x35] (PostMessage)
    code(&H24) = &HFF: code(&H25) = &H15
    Dim pAddr As Long: pAddr = m_pSubclassThunk + &H35
    CopyMemory code(&H26), pAddr, 4
    ' 2A: xor eax, eax (return 0)
    code(&H2A) = &H33: code(&H2B) = &HC0
    ' 2C: ret 16 (stdcall cleanup)
    code(&H2C) = &HC2: code(&H2D) = &H10: code(&H2E) = &H0
    ' 2F: callOriginal - jmp [pThunk+0x39] (oldProc)
    code(&H2F) = &HFF: code(&H30) = &H25
    pAddr = m_pSubclassThunk + &H39
    CopyMemory code(&H31), pAddr, 4
    ' 35: data - PostMessage address
    CopyMemory code(&H35), pPost, 4
    ' 39: data - old wndproc address
    CopyMemory code(&H39), m_OldWndProc, 4
    
    ' Copy to executable memory
    CopyMemory ByVal m_pSubclassThunk, code(0), 64
    
    ' Install subclass
    SetWindowLong hWnd, GWL_WNDPROC, m_pSubclassThunk
End Sub

Private Sub UnsubclassForm()
    If m_pSubclassThunk <> 0 And m_SubclassHwnd <> 0 Then
        SetWindowLong m_SubclassHwnd, GWL_WNDPROC, m_OldWndProc
        VirtualFree m_pSubclassThunk, 0&, MEM_RELEASE
        m_pSubclassThunk = 0
        m_OldWndProc = 0
        m_SubclassHwnd = 0
    End If
End Sub

Private Sub ConfigureForType()
    Select Case m_ControlType
        Case czTextBox
            txtEmbed.BackColor = TranslateColor(m_BackColor)
            txtEmbed.ForeColor = TranslateColor(m_ForeColor)
            If Len(m_PasswordChar) > 0 Then
                txtEmbed.PasswordChar = Left$(m_PasswordChar, 1)
            Else
                txtEmbed.PasswordChar = ""
            End If
            
            ' Show placeholder or text
            If Len(m_Text) > 0 Then
                txtEmbed.Text = m_Text
                m_IsPlaceholder = False
                txtEmbed.Visible = True
            ElseIf Len(m_PlaceholderText) > 0 Then
                txtEmbed.ForeColor = TranslateColor(CLR_PLACEHOLDER)
                txtEmbed.Text = m_PlaceholderText
                m_IsPlaceholder = True
                txtEmbed.Visible = True
            Else
                txtEmbed.Text = ""
                m_IsPlaceholder = False
                txtEmbed.Visible = True
            End If
            
            UpdateTextBoxFont
            PositionTextBox
            
        Case Else
            txtEmbed.Visible = False
    End Select
End Sub

Private Sub PositionTextBox()
    If m_ControlType <> czTextBox Then Exit Sub
    Dim pad As Long
    pad = m_CornerRadius
    If pad < 6 Then pad = 6
    
    Dim w As Long: w = UserControl.ScaleWidth
    Dim h As Long: h = UserControl.ScaleHeight
    Dim tbH As Long: tbH = h - pad * 2
    If tbH < 14 Then tbH = 14
    Dim tbTop As Long: tbTop = (h - tbH) \ 2
    
    On Error Resume Next
    txtEmbed.Move pad, tbTop, w - pad * 2, tbH
    On Error GoTo 0
End Sub

Private Sub UpdateTextBoxFont()
    On Error Resume Next
    txtEmbed.Font.Name = m_FontName
    txtEmbed.Font.Size = m_FontSize
    txtEmbed.Font.Bold = m_FontBold
    On Error GoTo 0
End Sub

'==============================================================================
' SECTION 16: MAIN RENDERING DISPATCHER
'==============================================================================
Private Sub RedrawControl()
    If Not m_Initialized Then Exit Sub
    
    Dim w As Long, h As Long
    w = UserControl.ScaleWidth
    h = UserControl.ScaleHeight
    If w <= 0 Or h <= 0 Then Exit Sub
    
    ' Clear using VB6 native (always works)
    UserControl.Cls
    
    ' Create GDI+ graphics from DC
    Dim hG As Long
    If GdipCreateFromHDC(UserControl.hDC, hG) <> 0 Then Exit Sub
    
    ' Enable high-quality rendering
    GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
    GdipSetTextRenderingHint hG, GP_TEXT_CLEARTYPE
    GdipSetPixelOffsetMode hG, GP_PIXELOFFSET_HQ
    
    ' Clear background
    ' For rounded controls: use parent's BackColor so corners look transparent
    ' For flat controls (Label, TitleBar): use own BackColor
    Dim clearColor As Long
    Select Case m_ControlType
        Case czLabel, czTitleBar
            clearColor = ColorToARGB(m_BackColor)
        Case Else
            On Error Resume Next
            clearColor = ColorToARGB(Ambient.BackColor)
            If Err.Number <> 0 Then clearColor = ColorToARGB(m_BackColor)
            On Error GoTo 0
    End Select
    
    Dim hBrush As Long
    GdipCreateSolidFill clearColor, hBrush
    GdipFillRectangle hG, hBrush, 0!, 0!, CSng(w), CSng(h)
    GdipDeleteBrush hBrush
    
    ' Dispatch to type-specific renderer
    Select Case m_ControlType
        Case czTitleBar:    RenderTitleBar hG, w, h
        Case czButton:      RenderButton hG, w, h
        Case czLabel:       RenderLabel hG, w, h
        Case czPanel:       RenderPanel hG, w, h
        Case czTextBox:     RenderTextBox hG, w, h
        Case czToggle:      RenderToggle hG, w, h
        Case czProgressBar: RenderProgressBar hG, w, h
    End Select
    
    GdipDeleteGraphics hG
    
    ' Draw form icon AFTER GDI+ is released
    If m_NeedDrawIcon Then
        m_NeedDrawIcon = False
        On Error Resume Next
        Dim icSize As Long: icSize = 16
        Dim icX As Long: icX = TB_ICON_WIDTH \ 2 - icSize \ 2
        Dim icY As Long: icY = h \ 2 - icSize \ 2
        Dim hIcon As Long
        Dim hSmallIcon As Long
        Dim needDestroy As Boolean: needDestroy = False
        
        ' Get icon handle from form
        hIcon = UserControl.Parent.Icon.Handle
        
        If Err.Number = 0 And hIcon <> 0 Then
            ' Extract 16x16 version from icon resource (not scaled from 32x32)
            hSmallIcon = CopyImage(hIcon, IMAGE_ICON, icSize, icSize, LR_COPYFROMRESOURCE)
            If hSmallIcon <> 0 Then
                DrawIconEx UserControl.hDC, icX, icY, hSmallIcon, icSize, icSize, 0, 0, DI_NORMAL
                DestroyIcon hSmallIcon
            Else
                ' Fallback: draw original (will be scaled)
                DrawIconEx UserControl.hDC, icX, icY, hIcon, icSize, icSize, 0, 0, DI_NORMAL
            End If
        End If
        On Error GoTo 0
    End If
    
    ' Flush persistent bitmap to screen
    If UserControl.AutoRedraw Then UserControl.Refresh
End Sub

'==============================================================================
' SECTION 17: RENDER - czLabel
'==============================================================================
Private Sub RenderLabel(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim clr As Long
    If m_Enabled Then
        clr = ColorToARGB(m_ForeColor)
    Else
        clr = ColorToARGB(m_ForeColor, 120)
    End If
    
    Dim fStyle As Long
    If m_FontBold Then fStyle = GP_FONT_BOLD Else fStyle = GP_FONT_REGULAR
    
    DrawText hG, m_Caption, 2!, 0!, CSng(w) - 4!, CSng(h), _
             m_FontName, m_FontSize, fStyle, clr, CLng(m_Alignment), GP_ALIGN_CENTER, GP_TRIM_ELLIPSIS
End Sub

'==============================================================================
' SECTION 18: RENDER - czButton
'==============================================================================
Private Sub RenderButton(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim fillColor As Long
    Dim textColor As Long
    Dim R As Single: R = CSng(m_CornerRadius)
    
    ' Determine button face color based on style
    Select Case m_ButtonStyle
        Case bsPrimary:   fillColor = ColorToARGB(m_ButtonColor)
        Case bsSecondary: fillColor = ColorToARGB(m_BackColor)
        Case bsDanger:    fillColor = ColorToARGB(CLR_DANGER)
        Case bsText:      fillColor = ColorToARGB(m_BackColor)
    End Select
    
    ' Apply state modifiers
    If Not m_Enabled Then
        fillColor = ColorToARGB(m_ButtonColor, 80)
    ElseIf m_IsPressed Then
        fillColor = DarkenColor(fillColor, 15)
    ElseIf m_IsHovering Then
        fillColor = LightenColor(fillColor, 15)
    End If
    
    ' Draw button face
    If m_ButtonStyle = bsText Then
        ' Text-only button: no background, just draw text with hover highlight
        If m_IsHovering Then
            FillRoundRect hG, 0!, 0!, CSng(w), CSng(h), R, ColorToARGB(CLR_BORDER, 60)
        End If
    Else
        ' Filled button with subtle gradient
        Dim topColor As Long: topColor = LightenColor(fillColor, 8)
        FillGradientRoundRect hG, 1!, 1!, CSng(w) - 2!, CSng(h) - 2!, R, topColor, fillColor
    End If
    
    ' Draw border for secondary style
    If m_ButtonStyle = bsSecondary Then
        DrawRoundRectBorder hG, 0!, 0!, CSng(w), CSng(h), R, ColorToARGB(CLR_BORDER), 1!
    End If
    
    ' Draw caption
    If m_Enabled Then
        textColor = ColorToARGB(m_ForeColor)
    Else
        textColor = ColorToARGB(m_ForeColor, 120)
    End If
    
    Dim fStyle As Long
    If m_FontBold Then fStyle = GP_FONT_BOLD Else fStyle = GP_FONT_REGULAR
    
    Dim yOff As Single: If m_IsPressed Then yOff = 1! Else yOff = 0!
    DrawText hG, m_Caption, 4!, yOff, CSng(w) - 8!, CSng(h), _
             m_FontName, m_FontSize, fStyle, textColor, GP_ALIGN_CENTER, GP_ALIGN_CENTER
End Sub

'==============================================================================
' SECTION 19: RENDER - czPanel
'==============================================================================
Private Sub RenderPanel(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim R As Single: R = CSng(m_CornerRadius)
    Dim panelColor As Long: panelColor = ColorToARGB(m_BackColor)
    
    ' When collapsed (notch only), skip panel body rendering
    If m_Expandable And w <= NOTCH_WIDTH Then
        ' Only draw notch button
        Dim cy As Single: cy = CSng(h) / 2!
        Dim nw As Single: nw = CSng(w)
        Dim nh As Single: nh = CSng(NOTCH_HEIGHT)
        Dim nr As Single: nr = CSng(NOTCH_RADIUS)
        
        ' Notch background
        Dim nbColor As Long
        If m_NotchHover Then
            nbColor = LightenColor(panelColor, 30)
        Else
            nbColor = LightenColor(panelColor, 15)
        End If
        
        ' Simple pill shape for collapsed notch
        FillRoundRect hG, 0!, cy - nh / 2!, nw, nh, nr, nbColor
        DrawRoundRectBorder hG, 0!, cy - nh / 2!, nw, nh, nr, _
                           ColorToARGB(CLR_BORDER, 100), 1!
        
        ' Chevron < (expand)
        Dim chevColor As Long: chevColor = ColorToARGB(CLR_WHITE, 220)
        Dim hChPen As Long
        GdipCreatePen1 chevColor, 1.5!, GP_UNIT_PIXEL, hChPen
        Dim cx As Single: cx = nw / 2!
        GdipDrawLine hG, hChPen, cx + 2!, cy - 5!, cx - 2!, cy
        GdipDrawLine hG, hChPen, cx - 2!, cy, cx + 2!, cy + 5!
        GdipDeletePen hChPen
        Exit Sub
    End If
    
    ' === Normal panel rendering (full or expanded) ===
    
    ' Fill panel with rounded rect
    FillRoundRect hG, 0!, 0!, CSng(w), CSng(h), R, panelColor
    
    ' Draw header strip if enabled
    If m_ShowHeader And Len(m_HeaderText) > 0 Then
        Dim headerH As Single: headerH = m_FontSize * 2.4!
        Dim headerColor As Long: headerColor = LightenColor(panelColor, 5)
        
        Dim hPath As Long, hBr As Long
        GdipCreatePath GP_FILL_ALTERNATE, hPath
        Dim d As Single: d = R * 2!
        GdipAddPathArc hPath, 0!, 0!, d, d, 180!, 90!
        GdipAddPathArc hPath, CSng(w) - d, 0!, d, d, 270!, 90!
        GdipAddPathLine hPath, CSng(w), R, CSng(w), headerH
        GdipAddPathLine hPath, CSng(w), headerH, 0!, headerH
        GdipAddPathLine hPath, 0!, headerH, 0!, R
        GdipClosePathFigure hPath
        
        GdipCreateSolidFill headerColor, hBr
        GdipFillPath hG, hBr, hPath
        GdipDeleteBrush hBr
        GdipDeletePath hPath
        
        Dim hPenSep As Long
        GdipCreatePen1 ColorToARGB(CLR_BORDER), 1!, GP_UNIT_PIXEL, hPenSep
        GdipDrawLine hG, hPenSep, 0!, headerH, CSng(w), headerH
        GdipDeletePen hPenSep
        
        DrawText hG, m_HeaderText, 12!, 0!, CSng(w) - 24!, headerH, _
                 m_FontName, m_FontSize, GP_FONT_BOLD, ColorToARGB(m_ForeColor), _
                 GP_ALIGN_NEAR, GP_ALIGN_CENTER
    End If
    
    ' Draw border
    If m_BorderWidth > 0 Then
        DrawRoundRectBorder hG, 0!, 0!, CSng(w), CSng(h), R, _
                           ColorToARGB(m_BorderColor), CSng(m_BorderWidth)
    End If
    
    ' Draw expand notch button at left edge (when expandable)
    If m_Expandable Then
        Dim ncw As Single: ncw = CSng(NOTCH_WIDTH)
        Dim nch As Single: nch = CSng(NOTCH_HEIGHT)
        Dim ncr As Single: ncr = CSng(NOTCH_RADIUS)
        Dim ncy As Single: ncy = CSng(h) / 2!
        
        Dim ncColor As Long
        If m_NotchHover Then
            ncColor = LightenColor(panelColor, 30)
        Else
            ncColor = LightenColor(panelColor, 15)
        End If
        
        ' Pill shape notch at x=0
        FillRoundRect hG, 0!, ncy - nch / 2!, ncw, nch, ncr, ncColor
        
        ' Chevron > (collapse) since panel is expanded
        Dim chvColor As Long: chvColor = ColorToARGB(CLR_WHITE, 220)
        Dim hCPen As Long
        GdipCreatePen1 chvColor, 1.5!, GP_UNIT_PIXEL, hCPen
        Dim chX As Single: chX = ncw / 2!
        GdipDrawLine hG, hCPen, chX - 2!, ncy - 5!, chX + 2!, ncy
        GdipDrawLine hG, hCPen, chX + 2!, ncy, chX - 2!, ncy + 5!
        GdipDeletePen hCPen
    End If
End Sub

'==============================================================================
' SECTION 20: RENDER - czTitleBar
'==============================================================================
Private Sub RenderTitleBar(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim btnW As Single: btnW = CSng(TB_BTN_WIDTH)
    Dim cx As Single, cy As Single  ' Center of current button
    cy = CSng(h) / 2!
    
    Dim glyphColor As Long: glyphColor = ColorToARGB(CLR_WHITE, 200)
    Dim hPen As Long, hBrush As Long
    
    ' --- Draw window control buttons (right-to-left) ---
    Dim btnX As Single: btnX = CSng(w)
    Dim g As Single: g = 4.5!  ' Half glyph size (~9px, matches Win11)
    
    ' Switch to pixel-perfect mode for ALL glyphs
    GdipSetPixelOffsetMode hG, 0  ' Default (no half-pixel offset)
    GdipSetSmoothingMode hG, 0    ' No anti-alias for crisp lines
    
    ' Close button
    If m_ShowCloseButton Then
        btnX = btnX - btnW
        cx = btnX + btnW / 2!
        
        ' Hover highlight
        If m_HotButton = 1 Then
            GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
            FillRoundRect hG, btnX, 0!, btnW, CSng(h), 0!, ColorToARGB(CLR_DANGER)
            GdipSetSmoothingMode hG, 0
            glyphColor = ColorToARGB(CLR_WHITE)
        Else
            glyphColor = ColorToARGB(CLR_WHITE, 200)
        End If
        
        ' Draw X glyph - pixel-aligned coordinates for sharp lines
        Dim xc As Long: xc = CLng(cx)
        Dim yc As Long: yc = CLng(cy)
        GdipCreatePen1 glyphColor, 1!, GP_UNIT_PIXEL, hPen
        GdipDrawLine hG, hPen, CSng(xc - 4), CSng(yc - 4), CSng(xc + 4), CSng(yc + 4)
        GdipDrawLine hG, hPen, CSng(xc + 4), CSng(yc - 4), CSng(xc - 4), CSng(yc + 4)
        GdipDeletePen hPen
    End If
    
    ' Fullscreen button
    If m_ShowFullScreenButton Then
        btnX = btnX - btnW
        cx = btnX + btnW / 2!
        
        If m_HotButton = 5 Then
            GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
            FillRoundRect hG, btnX, 0!, btnW, CSng(h), 0!, ColorToARGB(CLR_WHITE, 30)
            GdipSetSmoothingMode hG, 0
        End If
        
        glyphColor = ColorToARGB(CLR_WHITE, 200)
        GdipCreatePen1 glyphColor, 1!, GP_UNIT_PIXEL, hPen
        Dim fc As Long: fc = CLng(cx)
        Dim fy As Long: fy = CLng(cy)
        
        If m_IsFullScreen Then
            ' Exit fullscreen: 4 inward arrows (shrink)
            GdipDrawLine hG, hPen, CSng(fc - 5), CSng(fy - 1), CSng(fc - 1), CSng(fy - 1)
            GdipDrawLine hG, hPen, CSng(fc - 1), CSng(fy - 5), CSng(fc - 1), CSng(fy - 1)
            GdipDrawLine hG, hPen, CSng(fc + 1), CSng(fy - 1), CSng(fc + 5), CSng(fy - 1)
            GdipDrawLine hG, hPen, CSng(fc + 1), CSng(fy - 5), CSng(fc + 1), CSng(fy - 1)
            GdipDrawLine hG, hPen, CSng(fc - 5), CSng(fy + 1), CSng(fc - 1), CSng(fy + 1)
            GdipDrawLine hG, hPen, CSng(fc - 1), CSng(fy + 1), CSng(fc - 1), CSng(fy + 5)
            GdipDrawLine hG, hPen, CSng(fc + 1), CSng(fy + 1), CSng(fc + 5), CSng(fy + 1)
            GdipDrawLine hG, hPen, CSng(fc + 1), CSng(fy + 1), CSng(fc + 1), CSng(fy + 5)
        Else
            ' Enter fullscreen: 4 outward corner brackets (expand)
            GdipDrawLine hG, hPen, CSng(fc - 5), CSng(fy - 5), CSng(fc - 1), CSng(fy - 5)
            GdipDrawLine hG, hPen, CSng(fc - 5), CSng(fy - 5), CSng(fc - 5), CSng(fy - 1)
            GdipDrawLine hG, hPen, CSng(fc + 1), CSng(fy - 5), CSng(fc + 5), CSng(fy - 5)
            GdipDrawLine hG, hPen, CSng(fc + 5), CSng(fy - 5), CSng(fc + 5), CSng(fy - 1)
            GdipDrawLine hG, hPen, CSng(fc - 5), CSng(fy + 1), CSng(fc - 5), CSng(fy + 5)
            GdipDrawLine hG, hPen, CSng(fc - 5), CSng(fy + 5), CSng(fc - 1), CSng(fy + 5)
            GdipDrawLine hG, hPen, CSng(fc + 5), CSng(fy + 1), CSng(fc + 5), CSng(fy + 5)
            GdipDrawLine hG, hPen, CSng(fc + 1), CSng(fy + 5), CSng(fc + 5), CSng(fy + 5)
        End If
        GdipDeletePen hPen
    End If
    
    ' Maximize/Restore button (hidden in fullscreen)
    If m_ShowMaxButton And Not m_IsFullScreen Then
        btnX = btnX - btnW
        cx = btnX + btnW / 2!
        
        If m_HotButton = 2 Then
            GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
            FillRoundRect hG, btnX, 0!, btnW, CSng(h), 0!, ColorToARGB(CLR_WHITE, 30)
            GdipSetSmoothingMode hG, 0
        End If
        
        glyphColor = ColorToARGB(CLR_WHITE, 200)
        GdipCreatePen1 glyphColor, 1!, GP_UNIT_PIXEL, hPen
        
        ' Check if parent is maximized
        Dim isMax As Boolean
        If Ambient.UserMode Then
            Dim hP As Long: hP = GetParentHwnd()
            If hP <> 0 Then isMax = (IsZoomed(hP) <> 0)
        End If
        
        ' Pixel-aligned coordinates
        Dim gx As Long: gx = CLng(cx)
        Dim gy As Long: gy = CLng(cy)
        
        If isMax Then
            ' Restore icon: two overlapping rectangles (Win11 style)
            GdipDrawRectangle hG, hPen, CSng(gx - 2), CSng(gy - 4), 6!, 6!
            GdipDrawRectangle hG, hPen, CSng(gx - 4), CSng(gy - 2), 6!, 6!
            ' Fill overlap area with background to separate
            GdipCreateSolidFill ColorToARGB(m_BackColor), hBrush
            GdipFillRectangle hG, hBrush, CSng(gx - 2), CSng(gy - 2), 5!, 5!
            GdipDeleteBrush hBrush
            GdipDrawRectangle hG, hPen, CSng(gx - 2), CSng(gy - 2), 4!, 4!
        Else
            ' Maximize icon: single rectangle (9x9 to match Win11)
            GdipDrawRectangle hG, hPen, CSng(gx - 4), CSng(gy - 4), 9!, 9!
        End If
        GdipDeletePen hPen
    End If
    
    ' Minimize button
    If m_ShowMinButton Then
        btnX = btnX - btnW
        cx = btnX + btnW / 2!
        
        If m_HotButton = 3 Then
            GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
            FillRoundRect hG, btnX, 0!, btnW, CSng(h), 0!, ColorToARGB(CLR_WHITE, 30)
            GdipSetSmoothingMode hG, 0
        End If
        
        glyphColor = ColorToARGB(CLR_WHITE, 200)
        GdipCreatePen1 glyphColor, 1!, GP_UNIT_PIXEL, hPen
        Dim mx As Long: mx = CLng(cx)
        Dim my As Long: my = CLng(cy)
        GdipDrawLine hG, hPen, CSng(mx - 5), CSng(my), CSng(mx + 4), CSng(my)
        GdipDeletePen hPen
    End If
    
    ' --- Draw Icon area (left) ---
    If m_HotButton = 4 Then
        FillRoundRect hG, 0!, 0!, CSng(TB_ICON_WIDTH), CSng(h), 0!, ColorToARGB(CLR_WHITE, 20)
    End If
    
    ' Draw form icon area
    Dim iconCx As Long: iconCx = TB_ICON_WIDTH \ 2
    Dim iconCy As Long: iconCy = CLng(cy)
    
    If Ambient.UserMode Then
        ' Icon will be drawn AFTER GDI+ cleanup (in RedrawControl)
        ' to avoid GDI/GDI+ conflicts
        m_NeedDrawIcon = True
    Else
        ' Design-time: draw gear circle fallback
        GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
        GdipSetPixelOffsetMode hG, GP_PIXELOFFSET_HQ
        GdipCreatePen1 ColorToARGB(CLR_WHITE, 180), 1!, GP_UNIT_PIXEL, hPen
        GdipDrawEllipse hG, hPen, CSng(iconCx) - 4.5!, CSng(iconCy) - 4.5!, 9!, 9!
        GdipDeletePen hPen
        GdipCreateSolidFill ColorToARGB(m_BackColor), hBrush
        GdipFillEllipse hG, hBrush, CSng(iconCx) - 2!, CSng(iconCy) - 2!, 4!, 4!
        GdipDeleteBrush hBrush
    End If
    
    ' Restore rendering modes
    GdipSetSmoothingMode hG, GP_SMOOTH_ANTIALIAS
    GdipSetPixelOffsetMode hG, GP_PIXELOFFSET_HQ
    
    ' --- Draw Caption ---
    Dim textLeft As Single: textLeft = CSng(TB_ICON_WIDTH) + 4!
    Dim textWidth As Single: textWidth = btnX - textLeft
    If textWidth > 0 Then
        Dim fStyle As Long
        If m_FontBold Then fStyle = GP_FONT_BOLD Else fStyle = GP_FONT_BOLD  ' Title is always bold
        
        DrawText hG, m_Caption, textLeft, 0!, textWidth, CSng(h), _
                 m_FontName, m_FontSize, fStyle, ColorToARGB(m_ForeColor), _
                 GP_ALIGN_CENTER, GP_ALIGN_CENTER
    End If
End Sub

'==============================================================================
' SECTION 21: RENDER - czToggle
'==============================================================================
Private Sub RenderToggle(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim trackH As Single: trackH = CSng(h) * 0.65!
    If trackH < 18! Then trackH = 18!
    If trackH > CSng(h) - 4! Then trackH = CSng(h) - 4!
    
    Dim trackW As Single: trackW = trackH * 2!
    If trackW > CSng(w) - 4! Then trackW = CSng(w) - 4!
    
    Dim trackX As Single: trackX = (CSng(w) - trackW) / 2!
    Dim trackY As Single: trackY = (CSng(h) - trackH) / 2!
    Dim trackR As Single: trackR = trackH / 2!
    
    ' Interpolate track color between off and on based on animation position
    Dim t As Single: t = m_AnimPos  ' 0.0 = OFF, 1.0 = ON
    Dim onR As Long, onG As Long, onB As Long
    Dim offR As Long, offG As Long, offB As Long
    Dim mixR As Long, mixG As Long, mixB As Long
    
    onR = (m_OnColor And &HFF&): onG = ((m_OnColor \ &H100&) And &HFF&): onB = ((m_OnColor \ &H10000) And &HFF&)
    offR = (m_OffColor And &HFF&): offG = ((m_OffColor \ &H100&) And &HFF&): offB = ((m_OffColor \ &H10000) And &HFF&)
    
    mixR = CLng(offR + (onR - offR) * t)
    mixG = CLng(offG + (onG - offG) * t)
    mixB = CLng(offB + (onB - offB) * t)
    
    Dim trackColor As Long
    trackColor = ARGB(255, CByte(mixB), CByte(mixG), CByte(mixR))
    
    If Not m_Enabled Then trackColor = ColorToARGB(CLR_GRAY, 100)
    If m_IsHovering And m_Enabled Then trackColor = LightenColor(trackColor, 10)
    
    ' Draw track
    FillRoundRect hG, trackX, trackY, trackW, trackH, trackR, trackColor
    
    ' Draw thumb - interpolate position
    Dim thumbPad As Single: thumbPad = 3!
    Dim thumbD As Single: thumbD = trackH - thumbPad * 2!
    
    ' Hover zoom: grow thumb by 2px (1px each side)
    Dim grow As Single: grow = 0!
    If m_IsHovering And m_Enabled Then grow = 2!
    
    Dim thumbOff As Single: thumbOff = trackX + thumbPad - (grow / 2!)
    Dim thumbOn As Single: thumbOn = trackX + trackW - (thumbD + grow) - thumbPad + (grow / 2!)
    Dim thumbX As Single: thumbX = thumbOff + (thumbOn - thumbOff) * t
    Dim thumbY As Single: thumbY = trackY + thumbPad - (grow / 2!)
    
    ' Apply grow
    thumbD = thumbD + grow
    
    ' Thumb shadow
    Dim hBrShadow As Long
    GdipCreateSolidFill ARGB(50, 0, 0, 0), hBrShadow
    GdipFillEllipse hG, hBrShadow, thumbX + 1!, thumbY + 1!, thumbD, thumbD
    GdipDeleteBrush hBrShadow
    
    ' Thumb circle (white)
    Dim hBrThumb As Long
    GdipCreateSolidFill ARGB(255, 255, 255, 255), hBrThumb
    GdipFillEllipse hG, hBrThumb, thumbX, thumbY, thumbD, thumbD
    GdipDeleteBrush hBrThumb
End Sub

'==============================================================================
' SECTION 22: RENDER - czProgressBar
'==============================================================================
Private Sub RenderProgressBar(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim R As Single: R = CSng(m_CornerRadius)
    If R > CSng(h) / 2! Then R = CSng(h) / 2!
    
    Dim trackColor As Long: trackColor = ColorToARGB(CLR_BORDER)
    Dim barColor As Long: barColor = ColorToARGB(m_BarColor)
    
    ' Draw track
    FillRoundRect hG, 0!, 0!, CSng(w), CSng(h), R, trackColor
    
    ' Draw bar
    Dim pct As Long
    pct = m_Progress
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
    
    If pct > 0 Then
        Dim barW As Single
        barW = CSng(w) * CSng(pct) / 100!
        If barW < R * 2! Then barW = R * 2!
        
        Dim topBar As Long: topBar = LightenColor(barColor, 10)
        FillGradientRoundRect hG, 0!, 0!, barW, CSng(h), R, topBar, barColor
    End If
    
    ' Draw percentage text
    If m_ShowPercent Then
        Dim fStyle As Long: fStyle = GP_FONT_BOLD
        Dim txtSize As Single: txtSize = m_FontSize
        If txtSize > CSng(h) * 0.6! Then txtSize = CSng(h) * 0.6!
        
        DrawText hG, CStr(pct) & "%", 0!, 0!, CSng(w), CSng(h), _
                 m_FontName, txtSize, fStyle, ColorToARGB(CLR_WHITE), _
                 GP_ALIGN_CENTER, GP_ALIGN_CENTER
    End If
End Sub

'==============================================================================
' SECTION 23: RENDER - czTextBox
'==============================================================================
Private Sub RenderTextBox(ByVal hG As Long, ByVal w As Long, ByVal h As Long)
    Dim R As Single: R = CSng(m_CornerRadius)
    
    ' Determine border width
    Dim bw As Single
    If m_HasFocus Then bw = 2! Else bw = 1!
    
    ' Inset by half the border width so fill + border align perfectly
    Dim inset As Single: inset = bw / 2!
    
    ' Create a single path for both fill and border (inset from edges)
    Dim hPath As Long
    hPath = CreateRoundRectPath(inset, inset, CSng(w) - bw, CSng(h) - bw, R)
    
    ' Fill inside with BackColor
    Dim hBr As Long
    GdipCreateSolidFill ColorToARGB(m_BackColor), hBr
    GdipFillPath hG, hBr, hPath
    GdipDeleteBrush hBr
    
    ' Draw border on same path
    Dim borderClr As Long
    If m_HasFocus Then
        borderClr = ColorToARGB(m_FocusBorderColor)
    Else
        borderClr = ColorToARGB(m_BorderColor)
    End If
    
    Dim hPen As Long
    GdipCreatePen1 borderClr, bw, GP_UNIT_PIXEL, hPen
    GdipDrawPath hG, hPen, hPath
    GdipDeletePen hPen
    
    GdipDeletePath hPath
End Sub

'==============================================================================
' SECTION 24: PUBLIC PROPERTIES
'==============================================================================

' --- Universal Properties ---

Public Property Get ControlType() As czControlType
    ControlType = m_ControlType
End Property

Public Property Let ControlType(ByVal vNewValue As czControlType)
    m_ControlType = vNewValue
    ConfigureForType
    RedrawControl
    PropertyChanged "ControlType"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    m_Caption = vNewValue
    RedrawControl
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As String
    BackColor = ColorToHex(m_BackColor)
End Property

Public Property Let BackColor(ByVal vNewValue As String)
    m_BackColor = HexToColor(vNewValue)
    If m_ControlType = czTextBox Then
        txtEmbed.BackColor = TranslateColor(m_BackColor)
    End If
    RedrawControl
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As String
    ForeColor = ColorToHex(m_ForeColor)
End Property

Public Property Let ForeColor(ByVal vNewValue As String)
    m_ForeColor = HexToColor(vNewValue)
    If m_ControlType = czTextBox And Not m_IsPlaceholder Then
        txtEmbed.ForeColor = TranslateColor(m_ForeColor)
    End If
    RedrawControl
    PropertyChanged "ForeColor"
End Property

Public Property Get CornerRadius() As Long
    CornerRadius = m_CornerRadius
End Property

Public Property Let CornerRadius(ByVal vNewValue As Long)
    If vNewValue < 0 Then vNewValue = 0
    m_CornerRadius = vNewValue
    If m_ControlType = czTextBox Then PositionTextBox
    RedrawControl
    PropertyChanged "CornerRadius"
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal vNewValue As String)
    m_FontName = vNewValue
    If m_ControlType = czTextBox Then UpdateTextBoxFont
    RedrawControl
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal vNewValue As Single)
    If vNewValue < 1 Then vNewValue = 1
    m_FontSize = vNewValue
    If m_ControlType = czTextBox Then UpdateTextBoxFont
    RedrawControl
    PropertyChanged "FontSize"
End Property

Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
    m_FontBold = vNewValue
    If m_ControlType = czTextBox Then UpdateTextBoxFont
    RedrawControl
    PropertyChanged "FontBold"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_Enabled = vNewValue
    If m_ControlType = czTextBox Then txtEmbed.Enabled = m_Enabled
    UserControl.Enabled = m_Enabled
    RedrawControl
    PropertyChanged "Enabled"
End Property

Public Property Get Alignment() As czAlignment
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal vNewValue As czAlignment)
    m_Alignment = vNewValue
    RedrawControl
    PropertyChanged "Alignment"
End Property

' --- Button Properties ---

Public Property Get ButtonStyle() As czButtonStyle
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal vNewValue As czButtonStyle)
    m_ButtonStyle = vNewValue
    RedrawControl
    PropertyChanged "ButtonStyle"
End Property

Public Property Get ButtonColor() As String
    ButtonColor = ColorToHex(m_ButtonColor)
End Property

Public Property Let ButtonColor(ByVal vNewValue As String)
    m_ButtonColor = HexToColor(vNewValue)
    RedrawControl
    PropertyChanged "ButtonColor"
End Property

' --- TitleBar Properties ---

Public Property Get ShowMinButton() As Boolean
    ShowMinButton = m_ShowMinButton
End Property

Public Property Let ShowMinButton(ByVal vNewValue As Boolean)
    m_ShowMinButton = vNewValue
    RedrawControl
    PropertyChanged "ShowMinButton"
End Property

Public Property Get ShowMaxButton() As Boolean
    ShowMaxButton = m_ShowMaxButton
End Property

Public Property Let ShowMaxButton(ByVal vNewValue As Boolean)
    m_ShowMaxButton = vNewValue
    RedrawControl
    PropertyChanged "ShowMaxButton"
End Property

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = m_ShowCloseButton
End Property

Public Property Let ShowCloseButton(ByVal vNewValue As Boolean)
    m_ShowCloseButton = vNewValue
    RedrawControl
    PropertyChanged "ShowCloseButton"
End Property

Public Property Get ShowFullScreenButton() As Boolean
    ShowFullScreenButton = m_ShowFullScreenButton
End Property

Public Property Let ShowFullScreenButton(ByVal vNewValue As Boolean)
    m_ShowFullScreenButton = vNewValue
    RedrawControl
    PropertyChanged "ShowFullScreenButton"
End Property

Public Property Get AutoHandleButtons() As Boolean
    AutoHandleButtons = m_AutoHandleButtons
End Property

Public Property Let AutoHandleButtons(ByVal vNewValue As Boolean)
    m_AutoHandleButtons = vNewValue
    PropertyChanged "AutoHandleButtons"
End Property

' --- Panel Properties ---

Public Property Get BorderColor() As String
    BorderColor = ColorToHex(m_BorderColor)
End Property

Public Property Let BorderColor(ByVal vNewValue As String)
    m_BorderColor = HexToColor(vNewValue)
    RedrawControl
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal vNewValue As Long)
    If vNewValue < 0 Then vNewValue = 0
    m_BorderWidth = vNewValue
    RedrawControl
    PropertyChanged "BorderWidth"
End Property

Public Property Get ShowHeader() As Boolean
    ShowHeader = m_ShowHeader
End Property

Public Property Let ShowHeader(ByVal vNewValue As Boolean)
    m_ShowHeader = vNewValue
    RedrawControl
    PropertyChanged "ShowHeader"
End Property

Public Property Get HeaderText() As String
    HeaderText = m_HeaderText
End Property

Public Property Let HeaderText(ByVal vNewValue As String)
    m_HeaderText = vNewValue
    RedrawControl
    PropertyChanged "HeaderText"
End Property

' --- Toggle Properties ---

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Checked(ByVal vNewValue As Boolean)
    m_Checked = vNewValue
    RedrawControl
    PropertyChanged "Checked"
End Property

Public Property Get OnColor() As String
    OnColor = ColorToHex(m_OnColor)
End Property

Public Property Let OnColor(ByVal vNewValue As String)
    m_OnColor = HexToColor(vNewValue)
    RedrawControl
    PropertyChanged "OnColor"
End Property

Public Property Get OffColor() As String
    OffColor = ColorToHex(m_OffColor)
End Property

Public Property Let OffColor(ByVal vNewValue As String)
    m_OffColor = HexToColor(vNewValue)
    RedrawControl
    PropertyChanged "OffColor"
End Property

' --- TextBox Properties ---

Public Property Get Text() As String
    If m_ControlType = czTextBox And Not m_IsPlaceholder Then
        Text = txtEmbed.Text
    Else
        Text = m_Text
    End If
End Property

Public Property Let Text(ByVal vNewValue As String)
    m_Text = vNewValue
    If m_ControlType = czTextBox Then
        m_IsPlaceholder = False
        txtEmbed.ForeColor = TranslateColor(m_ForeColor)
        txtEmbed.Text = m_Text
        txtEmbed.Visible = True
    End If
    PropertyChanged "Text"
End Property

Public Property Get PlaceholderText() As String
    PlaceholderText = m_PlaceholderText
End Property

Public Property Let PlaceholderText(ByVal vNewValue As String)
    m_PlaceholderText = vNewValue
    If m_ControlType = czTextBox Then ConfigureForType
    PropertyChanged "PlaceholderText"
End Property

Public Property Get FocusBorderColor() As String
    FocusBorderColor = ColorToHex(m_FocusBorderColor)
End Property

Public Property Let FocusBorderColor(ByVal vNewValue As String)
    m_FocusBorderColor = HexToColor(vNewValue)
    RedrawControl
    PropertyChanged "FocusBorderColor"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = m_PasswordChar
End Property

Public Property Let PasswordChar(ByVal vNewValue As String)
    m_PasswordChar = vNewValue
    If m_ControlType = czTextBox Then
        If Len(vNewValue) > 0 Then
            txtEmbed.PasswordChar = Left$(vNewValue, 1)
        Else
            txtEmbed.PasswordChar = ""
        End If
    End If
    PropertyChanged "PasswordChar"
End Property

' --- ProgressBar Properties ---

Public Property Get Progress() As Long
    Progress = m_Progress
End Property

Public Property Let Progress(ByVal vNewValue As Long)
    If vNewValue < 0 Then vNewValue = 0
    If vNewValue > 100 Then vNewValue = 100
    m_Progress = vNewValue
    RedrawControl
    PropertyChanged "Progress"
End Property

Public Property Get BarColor() As String
    BarColor = ColorToHex(m_BarColor)
End Property

Public Property Let BarColor(ByVal vNewValue As String)
    m_BarColor = HexToColor(vNewValue)
    RedrawControl
    PropertyChanged "BarColor"
End Property

Public Property Get ShowPercent() As Boolean
    ShowPercent = m_ShowPercent
End Property

Public Property Let ShowPercent(ByVal vNewValue As Boolean)
    m_ShowPercent = vNewValue
    RedrawControl
    PropertyChanged "ShowPercent"
End Property

Public Property Get Expandable() As Boolean
    Expandable = m_Expandable
End Property

Public Property Let Expandable(ByVal vNewValue As Boolean)
    m_Expandable = vNewValue
    RedrawControl
    PropertyChanged "Expandable"
End Property

Public Property Get PanelExpanded() As Boolean
    PanelExpanded = m_PanelExpanded
End Property

Public Property Let PanelExpanded(ByVal vNewValue As Boolean)
    m_PanelExpanded = vNewValue
    RedrawControl
    PropertyChanged "PanelExpanded"
End Property

Public Property Get ExpandWidth() As Long
    ExpandWidth = m_ExpandWidth
End Property

Public Property Let ExpandWidth(ByVal vNewValue As Long)
    m_ExpandWidth = vNewValue
    PropertyChanged "ExpandWidth"
End Property

'==============================================================================
' SECTION 25: PUBLIC METHODS
'==============================================================================
Public Sub Refresh()
    ' Force a complete redraw
    RedrawControl
End Sub

Public Function WebColor(ByVal hex As String) As Long
    ' Convert web color format (#RRGGBB) to VB6 Long (BGR)
    ' Usage: czButton1.BackColor = czButton1.WebColor("#F5A623")
    hex = Replace(hex, "#", "")
    If Len(hex) <> 6 Then Exit Function
    Dim R As Long: R = Val("&H" & Mid$(hex, 1, 2))
    Dim G As Long: G = Val("&H" & Mid$(hex, 3, 2))
    Dim B As Long: B = Val("&H" & Mid$(hex, 5, 2))
    WebColor = RGB(R, G, B)
End Function
