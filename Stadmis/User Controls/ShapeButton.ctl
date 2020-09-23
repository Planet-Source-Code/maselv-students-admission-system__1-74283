VERSION 5.00
Begin VB.UserControl ShapeButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ShapeButton.ctx":0000
End
Attribute VB_Name = "ShapeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     maselv_e@yahoo.co.uk / elvasmasika@lexeme-kenya.com / masika_elvas@live.com *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************

'Masika Elvas (masika_elvas@programmer.net)
'Sun 10th Oct 2010
'- Formatted the whole codes to appear neat

Option Explicit

'>> Needed For The API Call >> GradientFill >> Background
Private Type GRADIENT_RECT
    
    UPPERLEFT As Long
    LOWERRIGHT As Long
    
End Type

Private Type TRIVERTEX
    
    x As Long
    y As Long
    Red As Integer
    Blue As Integer
    Green As Integer
    Alpha As Integer
    
End Type

'............................................................
'>>Needed For The API Call >> CreatePolygonRgn >> Polygon
'>>Postion For Picture And Caption ...etc.
Private Type PointAPI
    
    x As Long
    y As Long
    
End Type

'............................................................
'>> Constant Types Used With CreatePolygonRgn
Private Const WINDING    As Long = &H2

'............................................................
'>> Constant Types Used With GradientFill
Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V As Long = &H1

'............................................................
'>> Constant Types For LinsStyle
Private Const BDR_VISUAL As Long = vb3DDKShadow
Private Const BDR_VISUAL1 As Long = vbButtonShadow
Private Const BDR_VISUAL2 As Long = vb3DHighlight

'............................................................
'>> Constant Types For LinsStyle
Private Const BDR_FLAT1  As Long = vb3DDKShadow
Private Const BDR_FLAT2  As Long = vb3DHighlight

'............................................................
'>> Constant Types For LinsStyle
Private Const BDR_JAVA1  As Long = vbButtonShadow
Private Const BDR_JAVA2  As Long = vb3DHighlight

'............................................................
'>> Constant Values For White Border >> IF Mouse Down Button
Private Const BDR_PRESSED As Long = vb3DHighlight

'............................................................
'>> Constant Values For Gold Border >> IF Mouse Over Button
Private Const BDR_GOLDXP_DARK As Long = &H109ADC
Private Const BDR_GOLDXP_NORMAL1 As Long = &H31B2F0
Private Const BDR_GOLDXP_NORMAL2 As Long = &H90D6F7
Private Const BDR_GOLDXP_LIGHT1 As Long = &HCEF3FF
Private Const BDR_GOLDXP_LIGHT2 As Long = &H8CDBFF

'............................................................
'>> Constant Types For LinsStyle
Private Const BDR_VISTA1 As Long = vbWhite
Private Const BDR_VISTA2 As Long = &HCFB073

'............................................................
'>> Constant Types For FocusRect
'>> If Display Properties Colors IS Hight Color(16Bit)Use
'Private Const BDR_FOCUSRECT As Long = &HC0C0C0            '>> Color Gray

'>> If Display Properties Colors IS TrueColor(32Bit)Use
Private Const BDR_FOCUSRECT As Long = &HD1D1D1             '>> Color LightGray
'>> Constant Types For FocusRect Java
Private Const BDR_FOCUSRECT_JAVA As Long = &HCC9999        '>> ColorMauve
'>> Constant Values For FocusRect Vista
Private Const BDR_FOCUSRECT_VISTA As Long = 16698372
'>> Constant Values For FocusRect Xp
Private Const BDR_BLUEXP_DARK As Long = &HD98D59
Private Const BDR_BLUEXP_NORMAL1 As Long = &HE2A981
Private Const BDR_BLUEXP_NORMAL2 As Long = &HF0D1B5
Private Const BDR_BLUEXP_LIGHT1 As Long = &HF7D7BD
Private Const BDR_BLUEXP_LIGHT2 As Long = &HFFE7CE

'............................................................
'>> Constant Types For HandPointer
Private Const CURSOR_HAND = 32649&

'............................................................
'>> API Declare Function's
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Private Declare Function GetCapture& Lib "user32" ()
Private Declare Function SetCapture& Lib "user32" (ByVal hWnd&)
Private Declare Function ReleaseCapture& Lib "user32" ()
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As PointAPI) As Long
Private Declare Function TextOutW Lib "gdi32" Alias "TextOutA" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function TextOutA Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, lpPoint As PointAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hDc As Long, ByRef pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDc As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'............................................................
'>> Button Declarations
'............................................................
Public Enum EnumButtonShape
    
    Rectangle
    RoundedRectangle
    'Round
    Diamond
    Top_Triangle
    Left_Triangle
    Right_Triangle
    Down_Triangle
    top_Arrow
    Left_Arrow
    Right_Arrow
    Down_Arrow
    
End Enum

Public Enum EnumButtonStyle
    
    Custom
    Visual
    Flat
    OverFlat
    Java
    XPOffice
    WinXp
    Vista
    Glass
    
End Enum

Public Enum EnumButtonStyleColors
    
    Transparent
    SingleColor
    Gradient_H
    Gradient_V
    TubeCenter_H
    TubeTopBottom_H
    TubeCenter_V
    TubeTopBottom_V
    
End Enum

Public Enum EnumButtonTheme
    
    NoTheme
    XpBlue
    XpOlive
    XPSilver
    Visual2005
    Norton2005
    RedColor
    GreenColor
    BlueColor
    
End Enum

Public Enum EnumButtonType
    
    Button
    CheckBox
    
End Enum

Public Enum EnumCaptionAlignment
    
    TopCaption
    LeftCaption
    CenterCaption
    RightCaption
    BottomCaption
    
End Enum

Public Enum EnumCaptionEffect
    
    Default
    Raised
    Sunken
    Outline
    
End Enum

Public Enum EnumCaptionStyle
    
    Normal
    HorizontalFill
    VerticalFill
    
End Enum

Public Enum EnumDropDown
    
    None
    LeftDropDown
    RightDropDown
    
End Enum

Public Enum EnumPictureAlignment
    
    TopPicture
    LeftPicture
    CenterPicture
    RightPicture
    BottomPicture
    
End Enum

'............................................................
'>> Property Member Variables
'............................................................
Private mButtonShape       As EnumButtonShape
Private mButtonStyle       As EnumButtonStyle
Private mButtonStyleColors As EnumButtonStyleColors
Private mButtonTheme       As EnumButtonTheme
Private mButtonType        As EnumButtonType
Private mCaptionAlignment  As EnumCaptionAlignment
Private mCaptionEffect     As EnumCaptionEffect
Private mCaptionStyle      As EnumCaptionStyle
Private mDropDown          As EnumDropDown
Private mPictureAlignment  As EnumPictureAlignment

'............................................................
'>> Property Color Variables
Dim mBackColor           As OLE_COLOR
Dim mBackColorPressed    As OLE_COLOR
Dim mBackColorHover      As OLE_COLOR
'............................................................
Dim mBorderColor         As OLE_COLOR
Dim mBorderColorPressed  As OLE_COLOR
Dim mBorderColorHover    As OLE_COLOR
'............................................................
Dim mForeColor           As OLE_COLOR
Dim mForeColorPressed    As OLE_COLOR
Dim mForeColorHover      As OLE_COLOR
'............................................................
Dim mEffectColor         As OLE_COLOR

Dim mCaption             As String
Dim mFocusRect           As Boolean
Dim mFocused             As Boolean
Dim mValue               As Boolean
Dim mHandPointer         As Boolean
Dim mPicture             As Picture
Dim mPictureGray         As Boolean
Dim mTagExtra            As String
'............................................................
Dim CaptionPos(&H1)        As PointAPI                          '>> Postion Text
Dim PicturePos(&H1)        As PointAPI                          '>> Postion Picture
Dim Fo                   As PointAPI               '>> Region Of FocusRect
'............................................................

'>> Mouse Button >> Hovered Or Pressed
Private MouseMove, MouseDown As Boolean

Dim P(&H0 To &H7)            As PointAPI
Dim PL(&H0 To &H7)           As PointAPI
Dim Lines                As PointAPI
Dim hRgn                 As Long

'............................................................
'>> Button Event Declaration's
'............................................................
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(data As DataObject, AllowedEffects As Long)
'............................................................
'>> Start Properties Button
'............................................................
Public Property Get hDc() As Long
    hDc = UserControl.hDc
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Property Get ButtonShape() As EnumButtonShape
    ButtonShape = mButtonShape
End Property

Public Property Let ButtonShape(ByVal EBS As EnumButtonShape)
    mButtonShape = EBS
    PropertyChanged "ButtonShape"
    UserControl_Paint
End Property

Public Property Get ButtonStyle() As EnumButtonStyle
    ButtonStyle = mButtonStyle
End Property

Public Property Let ButtonStyle(ByVal EBS As EnumButtonStyle)
    
    mButtonStyle = EBS
    PropertyChanged "ButtonStyle"
    
    '.............................................................
    '>> ButtonStyle,Visual, Flat,OverFlat,Java,XpOffice,WinXp
    '>> Vista,Glass
    '>> Default BackColors To ButtonStyle At Hover,Press,Off
    '>> 'Default ForeColors To ButtonStyle At Hover,Press,Off
    '.............................................................
    
    Select Case mButtonStyle
        
        Case Is = Visual
            
            mBackColor = vbButtonFace
            mBackColorHover = vbButtonFace
            mBackColorPressed = vbButtonFace
            'mForeColor = vbBlack
            'mForeColorHover = vbBlack
            'mForeColorPressed = vbBlack
            
        Case Is = Flat
            
            mBackColor = vbButtonFace
            mBackColorHover = vbButtonFace
            mBackColorPressed = vbButtonFace
            'mForeColor = vbBlack
            ' mForeColorHover = vbBlack
            'mForeColorPressed = vbBlack
            
        Case Is = OverFlat
            
            mBackColor = vbButtonFace
            mBackColorHover = vbButtonFace
            mBackColorPressed = vbButtonFace
            'mForeColor = vbBlack
            'mForeColorHover = vbBlack
            'mForeColorPressed = vbBlack
            
        Case Is = Java
            
            If Not mButtonStyleColors = Transparent Then mButtonStyleColors = SingleColor
            mBackColor = vbButtonFace
            mBackColorHover = vbButtonFace
            mBackColorPressed = &H999999
            'mForeColor = vbBlack
            'mForeColorHover = vbBlack
            'mForeColorPressed = vbBlack
            
        Case Is = XPOffice
            
            mBackColor = vbButtonFace
            mBackColorHover = &HB6A59F
            mBackColorPressed = &HAF8C80
            'mForeColor = vbBlack
            'mForeColorHover = vbBlack
            'mForeColorPressed = vbBlack
            
        Case Is = WinXp
            
            If Not mButtonStyleColors = Transparent Then mButtonStyleColors = Gradient_V
            mBackColor = &HD2DEDD
            mBackColorHover = vbWhite
            mBackColorPressed = vbWhite
            'mForeColor = vbBlack
            'mForeColorHover = vbBlack
            'mForeColorPressed = vbBlack
            
        Case Is = Vista
            
            mBackColor = &HD8D8D8
            mBackColorHover = &HF7DBA5
            mBackColorPressed = &HEFCE92
            'mForeColor = &H8000000D    'vbhightlight
            'mForeColorHover = &H8000000D
            'mForeColorPressed = &H8000000D
            
        Case Is = Glass
            
            mBackColor = vbDesktop
            mBackColorHover = vbDesktop
            mBackColorPressed = vbDesktop
            'mForeColor = vbWhite
            'mForeColorHover = vbWhite
            'mForeColorPressed = vbWhite
            
    End Select
    
    UserControl_Paint
    
End Property

Public Property Get ButtonStyleColors() As EnumButtonStyleColors
    ButtonStyleColors = mButtonStyleColors
End Property

Public Property Let ButtonStyleColors(ByVal EBSC As EnumButtonStyleColors)
    mButtonStyleColors = EBSC
    PropertyChanged "ButtonStyleColors"
    UserControl_Paint
End Property

Public Property Get ButtonTheme() As EnumButtonTheme
    ButtonTheme = mButtonTheme
End Property

Public Property Let ButtonTheme(ByVal EBT As EnumButtonTheme)
    
    mButtonTheme = EBT
    PropertyChanged "ButtonTheme"
    
    '.............................................................
    '>> ButtonTheme,XpBlue,XpOlive,XPSilver,Visual2005,Norton2005
    '>> RedColor,GreenColor,BlueColor.
    '.............................................................
    
    Select Case mButtonTheme
        
        Case Is = XpBlue
            
            mBackColor = &HF1B39A
            mBackColorHover = &HFDF7F4
            mBackColorPressed = &HFAE6DD
            mBorderColor = &HF1B39A
            mBorderColorHover = &HF1B39A
            mBorderColorPressed = &HF1B39A
            
        Case Is = XpOlive
            
            mBackColor = &H3DB4A2
            mBackColorHover = &HDFF4F2
            mBackColorPressed = &HB8E7E0
            mBorderColor = &H3DB4A2
            mBorderColorHover = &H3DB4A2
            mBorderColorPressed = &H3DB4A2
            
        Case Is = XPSilver
            
            mBackColor = &HB5A09D
            mBackColorHover = &HF7F5F4
            mBackColorPressed = &HECE7E6
            mBorderColor = &HB5A09D
            mBorderColorHover = mBorderColor
            mBorderColorPressed = mBorderColor
            
        Case Is = Visual2005
            
            mBackColor = &HABC2C2
            mBackColorHover = &HF2F8F8
            mBackColorPressed = &HE6EDEE
            mBorderColor = &HABC2C2
            mBorderColorHover = mBorderColor
            mBorderColorPressed = mBorderColor
            
        Case Is = Norton2005
            
            mBackColor = &H2C5F5
            mBackColorHover = &HE2F9FF
            mBackColorPressed = &HAFEFFE
            mBorderColor = &H2C5F5
            mBorderColorHover = mBorderColor
            mBorderColorPressed = mBorderColor
            
        Case Is = RedColor
            
            mBackColor = &H26368B
            mBackColorHover = &H4763FF
            mBackColorPressed = &H425CEE
            mBorderColor = vbRed
            mBorderColorHover = mBorderColor
            mBorderColorPressed = mBorderColor
            
        Case Is = GreenColor
            
            mBackColor = &H578B2E
            mBackColorHover = &H9FFF54
            mBackColorPressed = &H80CD43
            mBorderColor = vbGreen
            mBorderColorHover = mBorderColor
            mBorderColorPressed = mBorderColor
            
        Case Is = BlueColor
            
            mBackColor = &H8B4027
            mBackColorHover = &HFF7648
            mBackColorPressed = &HCD5F3A
            mBorderColor = vbBlue
            mBorderColorHover = mBorderColor
            mBorderColorPressed = mBorderColor
            
    End Select
    
    UserControl_Paint
    
End Property

Public Property Get ButtonType() As EnumButtonType
    ButtonType = mButtonType
End Property

Public Property Let ButtonType(ByVal EBT As EnumButtonType)
    mButtonType = EBT
    PropertyChanged "ButtonType"
    UserControl_Paint
End Property

Public Property Get CaptionAlignment() As EnumCaptionAlignment
    CaptionAlignment = mCaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal ECA As EnumCaptionAlignment)
    mCaptionAlignment = ECA
    PropertyChanged "CaptionAlignment"
    UserControl_Paint
End Property

Public Property Get CaptionEffect() As EnumCaptionEffect
    CaptionEffect = mCaptionEffect
End Property

Public Property Let CaptionEffect(ByVal ECE As EnumCaptionEffect)
    mCaptionEffect = ECE
    PropertyChanged "CaptionEffect"
    UserControl_Paint
End Property

Public Property Get CaptionStyle() As EnumCaptionStyle
    CaptionStyle = mCaptionStyle
End Property

Public Property Let CaptionStyle(ByVal ECS As EnumCaptionStyle)
    mCaptionStyle = ECS
    PropertyChanged "CaptionStyle"
    UserControl_Paint
End Property

Public Property Get DropDown() As EnumDropDown
    DropDown = mDropDown
End Property

Public Property Let DropDown(ByVal EDD As EnumDropDown)
    mDropDown = EDD
    PropertyChanged "DropDown"
    UserControl_Paint
End Property

Public Property Get PictureAlignment() As EnumPictureAlignment
    PictureAlignment = mPictureAlignment
End Property

Public Property Let PictureAlignment(ByVal EPA As EnumPictureAlignment)
    mPictureAlignment = EPA
    PropertyChanged "PictureAlignment"
    UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    mBackColor = New_Color
    PropertyChanged "BackColor"
    UserControl_Paint
End Property

Public Property Get BackColorPressed() As OLE_COLOR
    BackColorPressed = mBackColorPressed
End Property

Public Property Let BackColorPressed(ByVal New_Color As OLE_COLOR)
    mBackColorPressed = New_Color
    PropertyChanged "BackColorPressed"
    UserControl_Paint
End Property

Public Property Get BackColorHover() As OLE_COLOR
    BackColorHover = mBackColorHover
End Property

Public Property Let BackColorHover(ByVal New_Color As OLE_COLOR)
    mBackColorHover = New_Color
    PropertyChanged "BackColorHover"
    UserControl_Paint
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(ByVal New_Color As OLE_COLOR)
    mBorderColor = New_Color
    PropertyChanged "BorderColor"
    UserControl_Paint
End Property

Public Property Get BorderColorPressed() As OLE_COLOR
    BorderColorPressed = mBorderColorPressed
End Property

Public Property Let BorderColorPressed(ByVal New_Color As OLE_COLOR)
    mBorderColorPressed = New_Color
    PropertyChanged "BorderColorPressed"
    UserControl_Paint
End Property

Public Property Get BorderColorHover() As OLE_COLOR
    BorderColorHover = mBorderColorHover
End Property

Public Property Let BorderColorHover(ByVal New_Color As OLE_COLOR)
    mBorderColorHover = New_Color
    PropertyChanged "BorderColorHover"
    UserControl_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal New_Color As OLE_COLOR)
    mForeColor = New_Color
    PropertyChanged "ForeColor"
    UserControl_Paint
End Property

Public Property Get ForeColorPressed() As OLE_COLOR
    ForeColorPressed = mForeColorPressed
End Property

Public Property Let ForeColorPressed(ByVal New_Color As OLE_COLOR)
    mForeColorPressed = New_Color
    PropertyChanged "ForeColorPressed"
    UserControl_Paint
End Property

Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = mForeColorHover
End Property

Public Property Let ForeColorHover(ByVal New_Color As OLE_COLOR)
    mForeColorHover = New_Color
    PropertyChanged "ForeColorHover"
    UserControl_Paint
End Property

Public Property Get EffectColor() As OLE_COLOR
    EffectColor = mEffectColor
End Property

Public Property Let EffectColor(ByVal New_Color As OLE_COLOR)
    mEffectColor = New_Color
    PropertyChanged "EffectColor"
    UserControl_Paint
End Property

'--------------------------------------------------------------------------------------------------------------------
'Masika Elvas (masika_elvas@programmer.net)
'Mon 11th Oct 2010 17:40
'- Added Property to allow use of Access Keys on the control
Public Property Get AccessKey() As String
Attribute AccessKey.VB_ProcData.VB_Invoke_Property = ";Data"
    AccessKey = UserControl.AccessKeys
End Property

Public Property Let AccessKey(ByVal NewAccessKey As String)
On Local Error GoTo Handle_LetAccessKey_Error
    
    UserControl.AccessKeys = NewAccessKey
    PropertyChanged "AccessKey"
    
Handle_LetAccessKey_Error:
    
End Property

'--------------------------------------------------------------------------------------------------------------------

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Caption.VB_UserMemId = 0
Attribute Caption.VB_MemberFlags = "1004"
    Caption = mCaption
End Property

Public Property Let Caption(ByVal NewCaption As String)
On Local Error GoTo Handle_LetCaption_Error
    
    'Masika Elvas (masika_elvas@programmer.net)
    'Mon 11th Oct 2010 17:40
    
    'If an Access Key has been specified then...
    If VBA.InStr(NewCaption, "&") <> &H0 Then
        
        'Extract the Access Key and assign to the AccessKey Property
        AccessKey = VBA.Mid(NewCaption, VBA.InStr(NewCaption, "&") + &H1, &H1)
        
        'Remove the Access Key setting from the Caption
        NewCaption = VBA.Replace(NewCaption, "&", VBA.vbNullString)
        
        'I have chosen this way because the '&' symbol remains visible
        'in the Caption when an Access Key is specified
        
    End If 'Close respective IF..THEN block statement
    
    mCaption = NewCaption
    PropertyChanged "Caption"
    UserControl_Paint
    
Handle_LetCaption_Error:
    
End Property

Public Property Get TagExtra() As String
    TagExtra = mTagExtra
End Property

Public Property Let TagExtra(ByVal NewTagExtra As String)
    mTagExtra = NewTagExtra
    PropertyChanged "TagExtra"
    UserControl_Paint
End Property

Public Property Get FocusRect() As Boolean
    FocusRect = mFocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As Boolean)
    mFocusRect = New_FocusRect
    PropertyChanged "FocusRect"
    UserControl_Paint
End Property

'Masika Elvas (masika_elvas@programmer.net)
'Sun 10th Oct 2010
'- Property invisible in the Properties section
'- Execute click command when set to True
Public Property Get Value() As Boolean
Attribute Value.VB_MemberFlags = "400"
    Value = mValue
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    mValue = New_Value
    PropertyChanged "Value"
    UserControl_Paint
    If mValue Then Call UserControl_Click
End Property

Public Property Get HandPointer() As Boolean
    HandPointer = mHandPointer
End Property

Public Property Let HandPointer(ByVal New_HandPointer As Boolean)
    mHandPointer = New_HandPointer
    PropertyChanged "HandPointer"
    UserControl_Paint
End Property

Public Property Get Picture() As Picture
    Set Picture = mPicture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set mPicture = New_Picture
    PropertyChanged "Picture"
    UserControl_Paint
End Property

Public Property Get PictureGray() As Boolean
    PictureGray = mPictureGray
End Property

Public Property Let PictureGray(ByVal New_PictureGray As Boolean)
    mPictureGray = New_PictureGray
    PropertyChanged "PictureGray"
    UserControl_Paint
End Property

'............................................................
'............................................................
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    UserControl_Paint
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    UserControl_Paint
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon() = New_MouseIcon
    PropertyChanged "MouseIcon"
    UserControl_Paint
End Property
'............................................................
'>> End Properties
'............................................................

'............................................................
'>> Convert Color's >> Red,Green,Blue To VB Color's
'............................................................
Private Sub ConvertRGB(ByVal Color As Long, R, G, b As Long)
    TranslateColor Color, &H0, Color
    R = Color And vbRed
    G = (Color And vbGreen) / 256
    b = (Color And vbBlue) / 65536
End Sub

'Masika Elvas (masika_elvas@programmer.net)
'Sun 10th Oct 2010
'- Called UserControl's click event so as to respond to DefaultCancel settings
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call UserControl_Click
End Sub

'............................................................
'>> Start UserControl
'............................................................
Private Sub UserControl_Click()
    '>> Start The Button Pressed
    If Me.ButtonType = CheckBox Then mValue = MouseDown = False
    UserControl_Paint
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    '>> Start The Button Pressed Too
    MouseDown = True
    UserControl_Paint
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    '>> Show FocusRect If Click Button
    mFocused = True
    UserControl_Paint
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    '>> Space Or '>> Enter
    If KeyCode = 32 Or KeyCode = 13 Then MouseDown = True: UserControl_Paint
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If MouseDown = True Then
        MouseDown = False: UserControl_Paint
        '>> Return Button To Original Shape If ButtonType >> CheckBox
        If KeyCode = vbKeyReturn Then UserControl_MouseDown &H0, &H0, &H0, &H0
    End If
    
    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub

Private Sub UserControl_LostFocus()
    '>> Visible FocusRect If Clicked Another Button
    mFocused = False
    UserControl_Paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '>> Start The Button Pressed
    MouseDown = True
    UserControl_Paint
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Dim PT      As PointAPI
    Dim Hovered As Boolean                         '>> False Or Ture
    
    With UserControl
        
        GetCursorPos PT
        ScreenToClient .hWnd, PT
        
        '>> We Don't Need To Check Postion Mouse
        If Not UserControl.Ambient.UserMode Then Exit Sub
        
        ' See If The Mouse Is Over The Control
        If (PT.x < &H0) Or (PT.y < &H0) Or (PT.x > .ScaleWidth) Or (PT.y > .ScaleHeight) Then ReleaseCapture Else SetCapture .hWnd
        
        '>> Mouse Is Leave Of Button or Over
        Hovered = Not ((PT.x < &H0) Or (PT.y < &H0) Or (PT.x > .ScaleWidth) Or (PT.y > .ScaleHeight))
        
        ' Redraw The Control If Necessary
        If MouseMove <> Hovered Then MouseMove = Hovered: MouseDown = False: UserControl_Paint
        
    End With
    
    RaiseEvent MouseMove(Button, Shift, x, y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Pressed As Boolean
    
    If GetCapture() <> UserControl.hWnd Then SetCapture UserControl.hWnd
    
    ' Raise The Click Event If The Button Is Currently Pressed
    If Pressed Then RaiseEvent Click
    
    Pressed = MouseDown
    
    ' Stop The Button Pressed
    MouseDown = False: UserControl_Paint
    RaiseEvent MouseUp(Button, Shift, x, y)
    
End Sub

Private Sub UserControl_InitProperties()
    'On Local Error Resume Next
    
    UserControl.Width = 1200
    UserControl.Height = 500

    mButtonShape = Rectangle
    mButtonStyle = Visual
    mButtonStyleColors = SingleColor
    mButtonTheme = NoTheme
    mButtonType = Button
    mCaptionAlignment = CenterCaption
    mCaptionStyle = Normal
    mCaptionEffect = Default
    mDropDown = None
    mPictureAlignment = CenterPicture

    mBackColor = vbButtonFace
    mBackColorPressed = vbButtonFace
    mBackColorHover = vbButtonFace

    'mBorderColor = vbBlack
    'mBorderColorHover = vbBlack
    'mBorderColorPressed = vbBlack

    mForeColor = vbBlack
    mForeColorPressed = vbRed
    mForeColorHover = vbBlue
    mEffectColor = vbWhite

    mCaption = Ambient.DisplayName
    mFocusRect = True
    mValue = False
    mHandPointer = False
    Set mPicture = Nothing
    mPictureGray = False
    UserControl.Enabled = True
    UserControl.Font = "Tahoma"
    UserControl_Paint
    
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    'On Local Error Resume Next
    With UserControl
        
        .AutoRedraw = True
        .ScaleMode = &H3
        .Cls
        
        '............................................................
        '>> Width = Height To Nice Show Border's.
        '............................................................
        If .Width < 375 Then .Width = 375
        If .Height < 375 Then .Height = 375
        If Not mButtonShape = Rectangle And Not mButtonShape = RoundedRectangle Then .Height = .Width
        
        '............................................................
        '>> Change MousePointer To HandCursor
        '............................................................
        If MouseDown Then
            If mFocused And mHandPointer Then SetCursor LoadCursor(&H0, CURSOR_HAND)
        ElseIf MouseMove Then
            If mHandPointer Then SetCursor LoadCursor(&H0, CURSOR_HAND)
        Else
            If mFocused And mHandPointer Then SetCursor LoadCursor(&H0, CURSOR_HAND)
        End If
        
        '............................................................
        '>> Use Value For Viwe CheckBox(ButtonPressed) At TimeShow
        '............................................................
        MouseDown = (mButtonType = CheckBox And mValue)
        
        '............................................................
        '>> Change BackColor When Mouse Hovered Or Pressed Or Off
        '............................................................
        Dim BkRed(&H0 To &H1), BkGreen(&H0 To &H1), BkBlue(&H0 To &H1) As Long
        Dim vert(&H1) As TRIVERTEX, GRect As GRADIENT_RECT
        
        If MouseDown Then
            ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
            ConvertRGB mBackColorPressed, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
        ElseIf MouseMove Then
            ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
            ConvertRGB mBackColorHover, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
        Else
            ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
            ConvertRGB mBackColorPressed, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
        End If
        
        vert(&H0).Red = Val("&H" & VBA.Hex(BkRed(&H0)) & "00"): vert(&H0).Green = Val("&H" & VBA.Hex(BkGreen(&H0)) & "00"): vert(&H0).Blue = Val("&H" & VBA.Hex(BkBlue(&H0)) & "00"): vert(&H0).Alpha = &H0&
        vert(&H1).Red = Val("&H" & VBA.Hex(BkRed(&H1)) & "00"): vert(&H1).Green = Val("&H" & VBA.Hex(BkGreen(&H1)) & "00"): vert(&H1).Blue = Val("&H" & VBA.Hex(BkBlue(&H1)) & "00"): vert(&H1).Alpha = &H0&
        GRect.UPPERLEFT = &H1: GRect.LOWERRIGHT = &H0
        
        '............................................................
        '>> ButtonStyleColors,Transparent,SingleColour
        '>> GradientHorizontalFill,VerticalGradientFill
        '>> SwapHorizontalFill,SwapVerticalFill,GradientTubeCenter_H
        '>> ,GradientTubeTopBottom_H,GradientTubeCenter_V
        '>> GradientTubeTopBottom_V
        '............................................................
        Select Case mButtonStyleColors
            
            Case Is = Transparent
                .BackStyle = &H0
                
            Case Is = SingleColor
                .BackStyle = &H1
                .BackColor = VBA.IIf(MouseDown, mBackColorPressed, VBA.IIf(MouseMove, mBackColorHover, mBackColor))
                
            Case Is = Gradient_H
                .BackStyle = &H1
                vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight
                vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                
            Case Is = Gradient_V
                .BackStyle = &H1
                vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight
                vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                
            Case Is = TubeCenter_H
                .BackStyle = &H1
                vert(&H0).x = .ScaleWidth: vert(&H0).y = &H0
                vert(&H1).x = &H0: vert(&H1).y = .ScaleHeight / &H2
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight
                vert(&H1).x = .ScaleWidth: vert(&H1).y = .ScaleHeight / &H2
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                
            Case Is = TubeTopBottom_H
                .BackStyle = &H1
                vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight / &H2
                vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight / &H2
                vert(&H1).x = .ScaleWidth: vert(&H1).y = .ScaleHeight
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                
            Case Is = TubeCenter_V
                .BackStyle = &H1
                vert(&H0).x = &H0: vert(&H0).y = &H0
                vert(&H1).x = .ScaleWidth / &H2: vert(&H1).y = .ScaleHeight
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                vert(&H0).x = .ScaleWidth: vert(&H0).y = .ScaleHeight
                vert(&H1).x = .ScaleWidth / &H2: vert(&H1).y = &H0
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                
            Case Is = TubeTopBottom_V
                .BackStyle = &H1
                vert(&H0).x = .ScaleWidth / &H2: vert(&H0).y = &H0
                vert(&H1).x = &H0: vert(&H1).y = .ScaleHeight
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                vert(&H0).x = .ScaleWidth / &H2: vert(&H0).y = .ScaleHeight
                vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                
        End Select
        
        '.............................................................
        '>> ButtonStyle,Visual, Flat,OverFlat,Java,XpOffice,WinXp
        '>> Vista,Glass
        '>> Set BorderColors To ButtonStyle At Hover,Press,Off
        '.............................................................
        
        Select Case mButtonStyle
            
            Case Is = Visual
                mBorderColor = vb3DDKShadow
                mBorderColorHover = vb3DDKShadow
                mBorderColorPressed = vb3DHighlight
                
            Case Is = Flat
                mBorderColor = vb3DDKShadow
                mBorderColorHover = vb3DDKShadow
                mBorderColorPressed = vb3DHighlight
                
            Case Is = OverFlat
                mBorderColor = mBackColor    '>> Visible BorderColor At RunTime
                mBorderColorHover = vb3DDKShadow
                mBorderColorPressed = vb3DHighlight
                
            Case Is = Java
                mBorderColor = vbButtonShadow
                mBorderColorHover = vbButtonShadow
                mBorderColorPressed = vbButtonShadow
                
            Case Is = XPOffice
                mBorderColor = vbButtonFace
                mBorderColorHover = vbBlack
                mBorderColorPressed = vbBlack
                
                '............................................................
                '>> Change BackColor XpOffice When Mouse Hovered,Pressed,Off
                '............................................................
                
                If Not mButtonStyleColors = Transparent Then
                    
                    If MouseDown Then
                        ConvertRGB &HA57F71, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        ConvertRGB mBackColorPressed, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                    ElseIf MouseMove Then
                        ConvertRGB &HAC9891, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        ConvertRGB mBackColorHover, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                    Else
                        ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        ConvertRGB mBackColor, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                    End If
                    
                    vert(&H0).Red = Val("&H" & VBA.Hex(BkRed(&H0)) & "00"): vert(&H0).Green = Val("&H" & VBA.Hex(BkGreen(&H0)) & "00"): vert(&H0).Blue = Val("&H" & VBA.Hex(BkBlue(&H0)) & "00"): vert(&H0).Alpha = &H0&
                    vert(&H1).Red = Val("&H" & VBA.Hex(BkRed(&H1)) & "00"): vert(&H1).Green = Val("&H" & VBA.Hex(BkGreen(&H1)) & "00"): vert(&H1).Blue = Val("&H" & VBA.Hex(BkBlue(&H1)) & "00"): vert(&H1).Alpha = &H0&
                    GRect.UPPERLEFT = &H1: GRect.LOWERRIGHT = &H0
                    vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight
                    vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                    GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                    
                End If
                
            Case Is = WinXp
                
                mBorderColor = vbBlack
                mBorderColorHover = vbBlack
                mBorderColorPressed = vbBlack
                
                '............................................................
                '>> Change BackColor WinXp When Mouse Hovered,Pressed,Off
                '............................................................
                
                If Not mButtonStyleColors = Transparent Then
                    
                    If MouseDown Then
                        ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        ConvertRGB &HE0E6E6, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                    ElseIf .Enabled Then
                        ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        ConvertRGB Me.BackColorPressed, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                    Else
                        ConvertRGB &HDBE7E7, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        ConvertRGB &HDBE7E7, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        mBorderColor = &H8DABAB
                    End If
                    
                    vert(&H0).Red = Val("&H" & VBA.Hex(BkRed(&H0)) & "00"): vert(&H0).Green = Val("&H" & VBA.Hex(BkGreen(&H0)) & "00"): vert(&H0).Blue = Val("&H" & VBA.Hex(BkBlue(&H0)) & "00"): vert(&H0).Alpha = &H0&
                    vert(&H1).Red = Val("&H" & VBA.Hex(BkRed(&H1)) & "00"): vert(&H1).Green = Val("&H" & VBA.Hex(BkGreen(&H1)) & "00"): vert(&H1).Blue = Val("&H" & VBA.Hex(BkBlue(&H1)) & "00"): vert(&H1).Alpha = &H0&
                    GRect.UPPERLEFT = &H1: GRect.LOWERRIGHT = &H0
                    vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight
                    vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                    GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                    
                End If
                
            Case Is = Vista
                
                mBorderColor = &H8F8F8E
                mBorderColorHover = &HB17F3C
                mBorderColorPressed = &H5C411D
                
                '............................................................
                '>> Change BackColor Vista When Mouse Hovered,Pressed,Off
                '............................................................
                
                If Me.ButtonStyle = Vista And Not Me.ButtonStyleColors = Transparent Then
                    
                    If MouseDown Then
                        ConvertRGB &HF8ECD5, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB mBackColorPressed, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                    ElseIf MouseMove Then
                        ConvertRGB &HFEFCF7, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB mBackColorHover, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                    ElseIf .Enabled Then
                        ConvertRGB &HEFEFEF, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                    Else
                        ConvertRGB &HF4F4F4, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB &HF0F0F0, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        mBorderColor = &HB5B2AD
                    End If
                    
                    vert(&H0).Red = Val("&H" & VBA.Hex(BkRed(&H0)) & "00"): vert(&H0).Green = Val("&H" & VBA.Hex(BkGreen(&H0)) & "00"): vert(&H0).Blue = Val("&H" & VBA.Hex(BkBlue(&H0)) & "00"): vert(&H0).Alpha = &H0&
                    vert(&H1).Red = Val("&H" & VBA.Hex(BkRed(&H1)) & "00"): vert(&H1).Green = Val("&H" & VBA.Hex(BkGreen(&H1)) & "00"): vert(&H1).Blue = Val("&H" & VBA.Hex(BkBlue(&H1)) & "00"): vert(&H1).Alpha = &H0&
                    GRect.UPPERLEFT = &H1: GRect.LOWERRIGHT = &H0
                    
                    If mButtonShape = Top_Triangle Or mButtonShape = Down_Triangle Or mButtonShape = top_Arrow Or mButtonShape = Down_Arrow Then
                        vert(&H0).x = .ScaleWidth / 1.5: vert(&H0).y = &H0
                        vert(&H1).x = &H0: vert(&H1).y = .ScaleHeight
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                        vert(&H0).x = .ScaleWidth / &H2: vert(&H0).y = .ScaleHeight
                        vert(&H1).x = .ScaleWidth + 25: vert(&H1).y = &H0
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                    Else
                        vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight / 1.5
                        vert(&H1).x = .ScaleWidth: vert(&H1).y = &H0
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                        vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight / &H2
                        vert(&H1).x = .ScaleWidth: vert(&H1).y = .ScaleHeight + 25
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                    End If
                    
                End If
                
            Case Is = Glass
                
                mBorderColor = mBackColor
                mBorderColorHover = mBackColorHover
                mBorderColorPressed = mBackColorPressed
                
                '............................................................
                '>> Change BackColor Glass When Mouse Hovered,Pressed,Off
                '............................................................
                
                If Not mButtonStyleColors = Transparent Then
                    
                    If MouseDown Then
                        ConvertRGB vbWhite, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB mBackColorPressed, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                    ElseIf MouseMove Then
                        ConvertRGB vbWhite, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB mBackColorHover, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                    ElseIf .Enabled Then
                        ConvertRGB vbWhite, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB mBackColor, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                    Else
                        ConvertRGB vbWhite, BkRed(&H1), BkGreen(&H1), BkBlue(&H1)
                        ConvertRGB &HA4B6B7, BkRed(&H0), BkGreen(&H0), BkBlue(&H0)
                        mBorderColor = &HA4B6B7
                    End If
                    
                    vert(&H0).Red = Val("&H" & VBA.Hex(BkRed(&H0)) & "00"): vert(&H0).Green = Val("&H" & VBA.Hex(BkGreen(&H0)) & "00"): vert(&H0).Blue = Val("&H" & VBA.Hex(BkBlue(&H0)) & "00"): vert(&H0).Alpha = &H0&
                    vert(&H1).Red = Val("&H" & VBA.Hex(BkRed(&H1)) & "00"): vert(&H1).Green = Val("&H" & VBA.Hex(BkGreen(&H1)) & "00"): vert(&H1).Blue = Val("&H" & VBA.Hex(BkBlue(&H1)) & "00"): vert(&H1).Alpha = &H0&
                    GRect.UPPERLEFT = &H1: GRect.LOWERRIGHT = &H0
                    
                    If mButtonShape = Top_Triangle Or mButtonShape = Down_Triangle Or mButtonShape = top_Arrow Or mButtonShape = Down_Arrow Then
                        vert(&H0).x = .ScaleWidth: vert(&H0).y = &H0
                        vert(&H1).x = -.ScaleWidth / &H6: vert(&H1).y = .ScaleHeight
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                        vert(&H0).x = .ScaleWidth / &H2: vert(&H0).y = .ScaleHeight
                        vert(&H1).x = .ScaleWidth + .ScaleWidth / &H3: vert(&H1).y = &H0
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_H
                    Else
                        vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight
                        vert(&H1).x = .ScaleWidth: vert(&H1).y = -.ScaleHeight / &H6
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                        vert(&H0).x = &H0: vert(&H0).y = .ScaleHeight / &H2
                        vert(&H1).x = .ScaleWidth: vert(&H1).y = .ScaleHeight + .ScaleHeight / &H3
                        GradientFill .hDc, vert(&H0), &H2, GRect, &H1, GRADIENT_FILL_RECT_V
                        
                    End If
                    
                End If
                
        End Select
        
        '............................................................
        '>> Get Postion Picture
        '............................................................
        
        If mPicture Is Nothing Then
            
            PicturePos(&H1).x = &H0
            PicturePos(&H1).y = &H0
            
        Else
            
            Dim mPicH&
            
            mPicH = mPicture.Height
            
            PicturePos(&H1).x = ScaleX(mPicture.Width, vbHimetric, vbPixels)
            PicturePos(&H1).y = ScaleY(mPicture.Height, vbHimetric, vbPixels)
            
            '............................................................
            '>> PictureAlignment,Top PictureLeft Picture,Center Picture
            '>> Right Picture,Bottom Picture
            '............................................................
            
            Select Case mPictureAlignment
                
                Case Is = TopPicture
                    PicturePos(&H0).x = (.ScaleWidth - PicturePos(&H1).x) / &H2
                    PicturePos(&H0).y = (.ScaleHeight - PicturePos(&H1).y) / .ScaleHeight + &H4
                Case Is = LeftPicture
                    PicturePos(&H0).x = (.ScaleWidth - PicturePos(&H1).x) / .ScaleWidth + &H4
                    PicturePos(&H0).y = (.ScaleHeight - PicturePos(&H1).y) / &H2
                Case Is = CenterPicture
                    PicturePos(&H0).x = (.ScaleWidth - PicturePos(&H1).x) / &H2
                    PicturePos(&H0).y = (.ScaleHeight - PicturePos(&H1).y) / &H2
                Case Is = RightPicture
                    PicturePos(&H0).x = (.ScaleWidth - PicturePos(&H1).x) - &H5
                    PicturePos(&H0).y = (.ScaleHeight - PicturePos(&H1).y) / &H2
                Case Is = BottomPicture
                    PicturePos(&H0).x = (.ScaleWidth - PicturePos(&H1).x) / &H2
                    PicturePos(&H0).y = (.ScaleHeight - PicturePos(&H1).y) - &H5
                    
            End Select
            
            '............................................................
            '>> Moving The Picture When Mouse Hovered Or Pressed
            '............................................................
            
            If Enabled Then
                
                Dim GrPic As PointAPI
                Dim Gray As Long
                
                If MouseDown Then
                    .PaintPicture mPicture, PicturePos(&H0).x + &H1, PicturePos(&H0).y + &H1    '>> +1 Pixel To Top And Right
                ElseIf MouseMove Then
                    
                    .PaintPicture mPicture, PicturePos(&H0).x, PicturePos(&H0).y
                    '>> If mButtonStyle = XPOffice Then -&H1 Pixel About Postion Picture When Mouse Hovered
                    If mButtonStyle = XPOffice Then .PaintPicture mPicture, PicturePos(&H0).x - &H1, PicturePos(&H0).y - &H1
                    
                Else
                    
                    .PaintPicture mPicture, PicturePos(&H0).x, PicturePos(&H0).y
                    
                    '>> Add Little Of Code's To Draw GrayScale Picture.
                    If mPictureGray = True Then
                        
                        For GrPic.x = PicturePos(&H0).x To PicturePos(&H0).x + PicturePos(&H1).x
                            
                            For GrPic.y = PicturePos(&H0).y To PicturePos(&H0).y + PicturePos(&H1).y
                                
                                Gray = 255 And (GetPixel(.hDc, GrPic.x, GrPic.y))
                                'Gray = vbRed And (GetPixel(.hDC, GrPic.X, GrPic.Y))
                                If GetPixel(.hDc, GrPic.x, GrPic.y) <> mBackColor Then SetPixel hDc, GrPic.x, GrPic.y, RGB(Gray, Gray, Gray)
                                
                            Next GrPic.y
                            
                        Next GrPic.x
                        
                    End If
                    
                End If
                
            Else                                               '>> If Enabled = False Then
                
                '............................................................
                '>> Converting Picture Colour's To Black Color If Button
                '>> Disabled
                '............................................................
                
                .PaintPicture mPicture, PicturePos(&H0).x, PicturePos(&H0).y
                
                For GrPic.x = PicturePos(&H0).x To PicturePos(&H0).x + PicturePos(&H1).x
                    
                    For GrPic.y = PicturePos(&H0).y To PicturePos(&H0).y + PicturePos(&H1).y
                        
                        Gray = 255 And (GetPixel(.hDc, GrPic.x, GrPic.y)) / 65536
                        If GetPixel(.hDc, GrPic.x, GrPic.y) <> mBackColor Then SetPixel hDc, GrPic.x, GrPic.y, RGB(Gray, Gray, Gray)
                        
                    Next GrPic.y
                    
                Next GrPic.x
                
            End If
            
        End If
        
        '............................................................
        '>> CaptionAlignment,Get Size And Postion Text,Top Text
        '>>Left Text,Center Text,Right Text,Bottom Text
        '............................................................
        
        If Len(mCaption) < &H0 Then Exit Sub
        
        Select Case mCaptionAlignment
            
            Case Is = TopCaption
                GetTextExtentPoint32 .hDc, mCaption, Len(mCaption), CaptionPos(&H1)
                CaptionPos(&H0).x = (.ScaleWidth - CaptionPos(&H1).x) / &H2
                CaptionPos(&H0).y = (.ScaleHeight - CaptionPos(&H1).y) / .ScaleHeight + &H5
            Case Is = LeftCaption
                GetTextExtentPoint32 .hDc, mCaption, Len(mCaption), CaptionPos(&H1)
                CaptionPos(&H0).x = (.ScaleWidth - CaptionPos(&H1).x) / .ScaleWidth + &H5
                CaptionPos(&H0).y = (.ScaleHeight - CaptionPos(&H1).y) / &H2
            Case Is = CenterCaption
                GetTextExtentPoint32 .hDc, mCaption, Len(mCaption), CaptionPos(&H1)
                CaptionPos(&H0).x = (.ScaleWidth - CaptionPos(&H1).x) / &H2
                CaptionPos(&H0).y = (.ScaleHeight - CaptionPos(&H1).y) / &H2
            Case Is = RightCaption
                GetTextExtentPoint32 .hDc, mCaption, Len(mCaption), CaptionPos(&H1)
                CaptionPos(&H0).x = (.ScaleWidth - CaptionPos(&H1).x) - &H5
                CaptionPos(&H0).y = (.ScaleHeight - CaptionPos(&H1).y) / &H2
            Case Is = BottomCaption
                GetTextExtentPoint32 .hDc, mCaption, Len(mCaption), CaptionPos(&H1)
                CaptionPos(&H0).x = (.ScaleWidth - CaptionPos(&H1).x) / &H2
                CaptionPos(&H0).y = (.ScaleHeight - CaptionPos(&H1).y) - &H5
                
        End Select
        
        '............................................................
        '>> CaptionEffect,Default,Raised,Sunken,Outline
        '............................................................
        
        Dim EFF As PointAPI
        
        If Enabled Then
            
            Select Case mCaptionEffect
                
                Case Is = Default
                    
                    If MouseDown Then
                        .ForeColor = mForeColorPressed
                        TextOutA .hDc, CaptionPos(&H0).x + &H1, CaptionPos(&H0).y + &H1, mCaption, Len(mCaption)
                    ElseIf MouseMove Then
                        .ForeColor = mForeColorHover
                        TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                    Else
                        .ForeColor = mForeColor
                        TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                    End If
                    
                Case Is = Raised
                    
                    EFF.x = -&H1                             '>> -&H1 Pixel From Down To Top
                    EFF.y = -&H1                             '>> -&H1 Pixel From Right To Left
                    
                    If MouseDown Then
                        .ForeColor = mEffectColor          '>> Vbwhite
                        TextOutW .hDc, CaptionPos(&H0).x + EFF.x + &H1, CaptionPos(&H0).y + EFF.y + &H1, mCaption, Len(mCaption)
                        .ForeColor = mForeColorPressed       '>> mForeColorPressed If You Like.
                        TextOutA .hDc, CaptionPos(&H0).x + &H1, CaptionPos(&H0).y + &H1, mCaption, Len(mCaption)
                    ElseIf MouseMove Then
                        .ForeColor = mEffectColor          '>> Vbwhite
                        TextOutW .hDc, CaptionPos(&H0).x + EFF.x, CaptionPos(&H0).y + EFF.y, mCaption, Len(mCaption)
                        .ForeColor = mForeColorHover
                        TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                    Else
                        .ForeColor = mEffectColor
                        TextOutW .hDc, CaptionPos(&H0).x + EFF.x, CaptionPos(&H0).y + EFF.y, mCaption, Len(mCaption)
                        .ForeColor = mForeColor
                        TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                    End If
                    
                Case Is = Sunken
                    
                    EFF.x = &H1                              '>> +1 Pixel From Top To Down
                    EFF.y = &H1                              '>> +1 Pixel From Left To Right
                    
                    If MouseDown Then
                        
                        .ForeColor = mEffectColor          '>> Vbwhite
                        TextOutW .hDc, CaptionPos(&H0).x + EFF.x + &H1, CaptionPos(&H0).y + EFF.y + &H1, mCaption, Len(mCaption)
                        .ForeColor = mForeColorPressed       '>> mForeColorPressed
                        TextOutA .hDc, CaptionPos(&H0).x + &H1, CaptionPos(&H0).y + &H1, mCaption, Len(mCaption)
                    ElseIf MouseMove Then
                        .ForeColor = mEffectColor          '>> Vbwhite
                        TextOutW .hDc, CaptionPos(&H0).x + EFF.x, CaptionPos(&H0).y + EFF.y, mCaption, Len(mCaption)
                        .ForeColor = mForeColorHover
                        TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                    Else
                        .ForeColor = mEffectColor
                        TextOutW .hDc, CaptionPos(&H0).x + EFF.x, CaptionPos(&H0).y + EFF.y, mCaption, Len(mCaption)
                        .ForeColor = mForeColor
                        TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                    End If
                    
                Case Is = Outline
                    
                    Dim Out As PointAPI
                    
                    For Out.x = -&H1 To &H1                    '>> -&H1 Pixel And +1 Pixel From Top To Down
                        
                        For Out.y = -&H1 To &H1                '>> -&H1 Pixel And +1 Pixel From Left To Right
                            
                            If MouseDown Then
                                .ForeColor = mEffectColor
                                TextOutW .hDc, CaptionPos(&H0).x + Out.x + &H1, CaptionPos(&H0).y + Out.y + &H1, mCaption, Len(mCaption)
                                .ForeColor = mForeColorPressed
                                TextOutA .hDc, CaptionPos(&H0).x + &H1, CaptionPos(&H0).y + &H1, mCaption, Len(mCaption)
                            ElseIf MouseMove Then
                                .ForeColor = mEffectColor
                                TextOutW .hDc, CaptionPos(&H0).x + Out.x, CaptionPos(&H0).y + Out.y, mCaption, Len(mCaption)
                                .ForeColor = mForeColorHover
                                TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                            Else
                                .ForeColor = mEffectColor
                                TextOutW .hDc, CaptionPos(&H0).x + Out.x, CaptionPos(&H0).y + Out.y, mCaption, Len(mCaption)
                                .ForeColor = mForeColor
                                TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
                            End If
                            
                        Next Out.y
                        
                    Next Out.x
                    
            End Select
            
        Else    '>> If Enabled = False Then >> Draw Disabled Text Sunken Look
            
            If mButtonStyle = Java Or mButtonStyle = WinXp Or mButtonStyle = Vista Or mButtonStyle = Glass Then
                .ForeColor = vb3DShadow
                TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
            Else    '>> If Enabled = False Then >> Draw Disabled Text Sunken Look
                .ForeColor = vb3DHighlight
                TextOutW .hDc, CaptionPos(&H0).x + &H1, CaptionPos(&H0).y + &H1, mCaption, Len(mCaption)
                .ForeColor = vb3DShadow
                TextOutA .hDc, CaptionPos(&H0).x, CaptionPos(&H0).y, mCaption, Len(mCaption)
            End If
            
        End If
        
        '............................................................
        '>> Change ForeColor When Mouse Hovered Or Pressed Or Off
        '............................................................
        
        Dim Fore As PointAPI
        Dim pos(&H1) As PointAPI
        Dim FrRed, FrGreen, FrBlue As Integer
        Dim TotalRed(&H1), TotalGreen(&H1), TotalBlue(&H1), R(&H1), G(&H1), b(&H1) As Long
        
        '............................................................
        '>> CaptionStyle,SingleColour,GradientHorizontalFill
        '>> GradientVerticalFill
        '............................................................
        Select Case mCaptionStyle
            
            Case Is = Normal                      '>> SingleColour >> MouseOff,MouseMove,MouseDown
                'We Don't Need To Any More Event's Here
                '>> Look To Event's '>> mCaptionEffect = Default
                
            Case Is = HorizontalFill                       '>> GradientFill Len(Caption)
                
                If MouseDown Then
                    ConvertRGB mForeColor, R(&H1), G(&H1), b(&H1)
                    ConvertRGB mForeColorPressed, R(&H0), G(&H0), b(&H0)
                ElseIf MouseMove Then
                    ConvertRGB mForeColor, R(&H1), G(&H1), b(&H1)
                    ConvertRGB mForeColorHover, R(&H0), G(&H0), b(&H0)
                Else
                    ConvertRGB mForeColor, R(&H1), G(&H1), b(&H1)
                    ConvertRGB mForeColorPressed, R(&H0), G(&H0), b(&H0)
                End If
                
                For Fore.x = CaptionPos(&H0).x To CaptionPos(&H0).x + CaptionPos(&H1).x
                    
                    For Fore.y = CaptionPos(&H0).y To CaptionPos(&H0).y + CaptionPos(&H1).y
                        
                        TotalRed(&H0) = (R(&H0) - R(&H1)) / .ScaleWidth * Fore.x
                        TotalGreen(&H0) = (G(&H0) - G(&H1)) / .ScaleWidth * Fore.x
                        TotalBlue(&H0) = (b(&H0) - b(&H1)) / .ScaleWidth * Fore.x
                        FrRed = R(&H1) + TotalRed(&H0)
                        FrGreen = G(&H1) + TotalGreen(&H0)
                        FrBlue = b(&H1) + TotalBlue(&H0)
                        
                        If GetPixel(hDc, Fore.x, Fore.y) = .ForeColor Then
                            
                            'If GetPixel(hdc, Fore.X, Fore.Y) <> .ForeColor Then'>> For Show Len(Caption) Fill
                            SetPixel .hDc, Fore.x, Fore.y, RGB(FrRed, FrGreen, FrBlue)
                            R(&H1) = R(&H1) + TotalRed(&H1)
                            G(&H1) = G(&H1) + TotalGreen(&H1)
                            b(&H1) = b(&H1) + TotalBlue(&H1)
                            
                        End If
                        
                    Next Fore.y
                    
                Next Fore.x
                
            Case Is = VerticalFill
                
                If MouseDown Then
                    ConvertRGB mForeColor, R(&H0), G(&H0), b(&H0)
                    ConvertRGB mForeColorPressed, R(&H1), G(&H1), b(&H1)
                ElseIf MouseMove Then
                    ConvertRGB mForeColor, R(&H0), G(&H0), b(&H0)
                    ConvertRGB mForeColorHover, R(&H1), G(&H1), b(&H1)
                Else
                    ConvertRGB mForeColor, R(&H0), G(&H0), b(&H0)
                    ConvertRGB mForeColorPressed, R(&H1), G(&H1), b(&H1)
                End If
                
                For Fore.x = CaptionPos(&H0).x To CaptionPos(&H0).x + CaptionPos(&H1).x
                    
                    For Fore.y = CaptionPos(&H0).y To CaptionPos(&H0).y + CaptionPos(&H1).y
                        
                        TotalRed(&H0) = (R(&H0) - R(&H1)) / .ScaleHeight * Fore.y
                        TotalGreen(&H0) = (G(&H0) - G(&H1)) / .ScaleHeight * Fore.y
                        TotalBlue(&H0) = (b(&H0) - b(&H1)) / .ScaleHeight * Fore.y
                        FrRed = R(&H1) + TotalRed(&H0)
                        FrGreen = G(&H1) + TotalGreen(&H0)
                        FrBlue = b(&H1) + TotalBlue(&H0)
                        
                        If GetPixel(hDc, Fore.x, Fore.y) = .ForeColor Then
                            
                            SetPixel .hDc, Fore.x, Fore.y, RGB(FrRed, FrGreen, FrBlue)
                            R(&H1) = R(&H1) + TotalRed(&H1)
                            G(&H1) = G(&H1) + TotalGreen(&H1)
                            b(&H1) = b(&H1) + TotalBlue(&H1)
                            
                        End If
                        
                    Next Fore.y
                    
                Next Fore.x
                
        End Select
        
        .MaskColor = .BackColor                            '>> Change ButtonStyleColor From Opaque To Transparent
        
        '............................................................
        '>> Change BorderColor When Mouse Hovered Or Pressed Or Off
        '............................................................
        
        .ForeColor = VBA.IIf(MouseDown, mBorderColorPressed, VBA.IIf(MouseMove, mBorderColorHover, mBorderColor))
        
        '............................................................
        '>> Calling ButtonShape's
        '............................................................
        Select Case mButtonShape
            
            Case Is = Rectangle: Redraw_Rectangle
            Case Is = RoundedRectangle: Redraw_RoundedRectangle
            'Case Is = Round: Redraw_Round
            Case Is = Diamond: Redraw_Diamond
            Case Is = Top_Triangle: Redraw_Top_Triangle
            Case Is = Left_Triangle: Redraw_Left_Triangle
            Case Is = Right_Triangle: Redraw_Right_Triangle
            Case Is = Down_Triangle: Redraw_Down_Triangle
            Case Is = top_Arrow: Redraw_Top_Arrow
            Case Is = Left_Arrow: Redraw_Left_Arrow
            Case Is = Right_Arrow: Redraw_Right_Arrow
            Case Is = Down_Arrow: Redraw_Down_Arrow
            
        End Select
        
        '............................................................
        '>> Use GetPixel And SetPixel For Make Point's On FocusRect.
        '............................................................
        
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    For Fo.x = &H3 To .ScaleWidth - &H3 Step &H2
                        
                        For Fo.y = &H3 To .ScaleHeight - &H3 Step &H2
                            If GetPixel(.hDc, Fo.x, Fo.y) = .ForeColor Then SetPixel .hDc, Fo.x, Fo.y, vbBlack
                        Next Fo.y
                        
                    Next Fo.x
                        
                    For Fo.x = &H2 To .ScaleWidth - &H2 Step &H2
                        
                        For Fo.y = &H2 To .ScaleHeight - &H2 Step &H2
                            If GetPixel(.hDc, Fo.x, Fo.y) = .ForeColor Then SetPixel .hDc, Fo.x, Fo.y, vbBlack
                        Next Fo.y
                        
                    Next Fo.x
                    
                End If
                
            End If
            
        End If
        
        '............................................................
        '>> Draw FocusRect Java About The Text.
        '............................................................
        
        If Len(mCaption) > &H0 Then
            
            If mButtonStyle = Java And mFocusRect And mFocused Then
                .ForeColor = BDR_FOCUSRECT_JAVA
                RoundRect hDc, CaptionPos(&H1).x + CaptionPos(&H0).x + &H2, CaptionPos(&H1).y + CaptionPos(&H0).y + &H2, CaptionPos(&H0).x - &H3, CaptionPos(&H0).y - &H2, &H0, &H0
            End If
            
        End If
        
        '............................................................
        '>> DropDown,Left DropDown,Right DropDown
        '............................................................
        
        Dim FLAGS(&H1)     As PointAPI
        
        Select Case mDropDown
            
            Case Is = LeftDropDown
                FLAGS(&H0).x = &H2
                FLAGS(&H0).y = .ScaleHeight / &H2 + FLAGS(&H0).y / .ScaleHeight - &H8
            Case Is = RightDropDown
                FLAGS(&H0).x = .ScaleWidth - 21
                FLAGS(&H0).y = .ScaleHeight / &H2 + FLAGS(&H0).y / .ScaleHeight - &H8
                
        End Select
        
        '............................................................
        '>> DropDown >> Left DropDown,Right DropDown
        '............................................................
        
        If Not mDropDown = None Then                       '>> So If mDropDown=LeftDropDown Or RightDropDown then
            
            .ForeColor = vbBlack                           '>> mForeColor If You Like.
            'Drwa Chevron
            MoveToEx .hDc, FLAGS(&H0).x + &H6, FLAGS(&H0).y + &H5, FLAGS(&H1)
            LineTo .hDc, FLAGS(&H0).x + &H9, FLAGS(&H0).y + &H8
            LineTo .hDc, FLAGS(&H0).x + &HD, FLAGS(&H0).y + &H4
            MoveToEx .hDc, FLAGS(&H0).x + &H7, FLAGS(&H0).y + &H5, FLAGS(&H1)
            LineTo .hDc, FLAGS(&H0).x + &H9, FLAGS(&H0).y + &H7
            LineTo .hDc, FLAGS(&H0).x + &HC, FLAGS(&H0).y + &H4
            MoveToEx .hDc, FLAGS(&H0).x + &H6, FLAGS(&H0).y + &H9, FLAGS(&H1)
            LineTo .hDc, FLAGS(&H0).x + &H9, FLAGS(&H0).y + &HC
            LineTo .hDc, FLAGS(&H0).x + &HD, FLAGS(&H0).y + &H8
            MoveToEx .hDc, FLAGS(&H0).x + &H7, FLAGS(&H0).y + &H9, FLAGS(&H1)
            LineTo .hDc, FLAGS(&H0).x + &H9, FLAGS(&H0).y + &HB
            LineTo .hDc, FLAGS(&H0).x + &HC, FLAGS(&H0).y + &H8
            
        End If
        
        .MaskPicture = .Image                              '>> Change ButtonStyleColor From Transparent To Opaque When Add Picture.
        .Refresh                                           '>> AutoRedraw=True
        
    End With
    
End Sub

Private Sub Redraw_Rectangle()
    
    Dim hRgn As Long
    
    With UserControl
        
        '>> Rectangle >> Make Rectangle Shape.
        LineTo .hDc, &H0, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth - &H1, &H0
        LineTo .hDc, &H0, &H0
        
        SetWindowRgn .hWnd, hRgn, False                    '>>Retune Original Shape To Button Without Cutting
        
        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Rectangle >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth - &H1, &H0, Lines
                    LineTo .hDc, &H0, &H0
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H2, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                    LineTo .hDc, &H0, .ScaleHeight - &H2
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H0, Lines
                    LineTo .hDc, &H0, &H0
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Flat                                 '>> Rectangle >> Draw Flat Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H1, &H0, Lines
                    LineTo .hDc, &H0, &H0
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    
                Else
                    
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H0, Lines
                    LineTo .hDc, &H0, &H0
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    
                End If
                
            Case Is = OverFlat                             '>> Rectangle >> Draw OverFlat Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H1, &H0, Lines
                    LineTo .hDc, &H0, &H0
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H0, Lines
                    LineTo .hDc, &H0, &H0
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Java                                 '>> Rectangle >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth - &H2, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                    LineTo .hDc, &H2, .ScaleHeight - &H2
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                    
                End If
                
            Case Is = WinXp                                '>> Rectangle >> Draw WindowsXp Style
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, &H2, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H3, &H2
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &H2
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, &H1
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H3
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight - &H4, Lines
                    LineTo .hDc, .ScaleWidth - &H3, &H2
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    '>> Rectangle >> Draw Xp FocusRect.
                    
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &H2
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, &H1
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H3
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight - &H4, Lines
                    LineTo .hDc, .ScaleWidth - &H3, &H2
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    
                End If
                
            Case Is = Vista    '>> Rectangle >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, &H1
                    LineTo .hDc, &H1, &H1
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, &H1
                    LineTo .hDc, &H1, &H1
                    
                    '>> Rectangle >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, &H1, &H1, Lines
                        LineTo .hDc, &H1, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth - &H2, &H1
                        LineTo .hDc, &H1, &H1
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Rectangle >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    MoveToEx .hDc, &H4, &H4, Lines
                    LineTo .hDc, &H4, .ScaleHeight - &H5
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight - &H5
                    LineTo .hDc, .ScaleWidth - &H5, &H4
                    LineTo .hDc, &H4, &H4
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_RoundedRectangle()
    
    With UserControl
        
        '>> RoundedRectangle >> Make RoundedRectangle Shape.
        hRgn = CreateRoundRectRgn(&H0, &H0, .ScaleWidth, .ScaleHeight, &HF, &HF)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> RoundedRectangle >> Draw RoundedRectangle Shape.
        RoundRect hDc, &H0, &H0, .ScaleWidth - &H1, .ScaleHeight - &H1, &H10, &H10

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> RoundedRectangle >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, &HA, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H0
                    Arc .hDc, &H11, &H11, &H0, &H0, &HA, &H0, &H0, &HA
                    MoveToEx .hDc, &H0, &HA, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H1, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight
                    Arc .hDc, &H10, .ScaleHeight - &H10, &H0, .ScaleHeight - &H2, -&H10, .ScaleHeight - &H10, &H8, .ScaleHeight
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, &HA, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H1
                    Arc .hDc, &H11, &H11, &H1, &H1, &HA, &H1, &H1, &HA
                    MoveToEx .hDc, &H1, &HA, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H10, &H1, .ScaleHeight - &H2, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    RoundRect hDc, &H0, &H0, .ScaleWidth - &H2, .ScaleHeight - &H2, &HF, &HF
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, &HA, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H0
                    Arc .hDc, &H11, &H11, &H0, &H0, &HA, &H0, &H0, &HA
                    MoveToEx .hDc, &H0, &HA, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H2, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H1, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Flat                                 '>> RoundedRectangle >> Draw Flat Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, &HA, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H0
                    Arc .hDc, &H11, &H11, &H0, &H0, &HA, &H0, &H0, &HA
                    MoveToEx .hDc, &H0, &HA, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H1, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    
                Else
                    
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, &HA, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H0
                    Arc .hDc, &H11, &H11, &H0, &H0, &HA, &H0, &H0, &HA
                    MoveToEx .hDc, &H0, &HA, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H1, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    
                End If
                
            Case Is = OverFlat                             '>> RoundedRectangle >> Draw OverFlat Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, &HA, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H0
                    Arc .hDc, &H11, &H11, &H0, &H0, &HA, &H0, &H0, &HA
                    MoveToEx .hDc, &H0, &HA, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H1, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, &HA, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H0
                    Arc .hDc, &H11, &H11, &H0, &H0, &HA, &H0, &H0, &HA
                    MoveToEx .hDc, &H0, &HA, Lines
                    LineTo .hDc, &H0, .ScaleHeight - &HA
                    Arc .hDc, &H11, .ScaleHeight - &H11, &H0, .ScaleHeight - &H1, -&H10, .ScaleHeight - &H10, &H9, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Java                                 '>> RoundedRectangle >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    RoundRect hDc, &H0, &H0, .ScaleWidth - &H2, .ScaleHeight - &H2, &H10, &H10
                    .ForeColor = BDR_JAVA2
                    RoundRect hDc, &H1, &H1, .ScaleWidth - &H1, .ScaleHeight - &H1, &H10, &H10
                    
                End If
                
            Case Is = WinXp                                '>> RoundedRectangle >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    Arc .hDc, .ScaleWidth - &H11, .ScaleHeight - &H11, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth - &HA, .ScaleHeight - &H3, .ScaleWidth - &H2, .ScaleHeight - &HA
                    MoveToEx .hDc, .ScaleWidth - &H9, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H8, .ScaleHeight - &H4
                    MoveToEx .hDc, .ScaleWidth - &H4, &H4, Lines
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight - &H9
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    Arc .hDc, &H10, .ScaleHeight - &H11, &H1, .ScaleHeight - &H2, -&H11, .ScaleHeight, &H8, .ScaleHeight - &H1
                    MoveToEx .hDc, .ScaleWidth - &H8, .ScaleHeight - &H3, Lines
                    LineTo .hDc, &H8, .ScaleHeight - &H3
                    Arc .hDc, .ScaleWidth - &H11, .ScaleHeight - &H11, .ScaleWidth - &H2, .ScaleHeight - &H2, .ScaleWidth - &HA, .ScaleHeight - &H2, .ScaleWidth - &H1, .ScaleHeight - &HA
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    Arc .hDc, &H11, .ScaleHeight - &H12, &H1, .ScaleHeight - &H3, -&H12, .ScaleHeight, &HA, .ScaleHeight - &H1
                    MoveToEx .hDc, &H1, &H6, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H8
                    MoveToEx .hDc, .ScaleWidth - &H8, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H8, .ScaleHeight - &H4
                    Arc .hDc, .ScaleWidth - &H11, .ScaleHeight - &H11, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth - &HA, .ScaleHeight - &H2, .ScaleWidth - &H1, .ScaleHeight - &HA
                    MoveToEx .hDc, .ScaleWidth - &H3, &H6, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H8
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, &H2, &H7, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H7
                    MoveToEx .hDc, .ScaleWidth - &H4, &H7, Lines
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight - &H7
                    Arc .hDc, &H11, &H11, &H1, &H1, &H8, &H0, &H0, &H8
                    Arc .hDc, .ScaleWidth - &H11, &H11, .ScaleWidth - &H2, &H1, .ScaleWidth - &H1, &H8, .ScaleWidth - &HC, &H1
                    Arc .hDc, .ScaleWidth - &H11, &H11, .ScaleWidth - &H3, &H2, .ScaleWidth - &H1, &H8, .ScaleWidth - &HC, &H1
                    Arc .hDc, &H11, &H11, &H2, &H2, &H8, &H0, &H0, &H8
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, &HA, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H7, &H1
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, &HA, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H2
                    
                    '>> RoundedRectangle >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    Arc .hDc, &H10, .ScaleHeight - &H11, &H1, .ScaleHeight - &H2, -&H11, .ScaleHeight, &H8, .ScaleHeight - &H1
                    MoveToEx .hDc, .ScaleWidth - &H8, .ScaleHeight - &H3, Lines
                    LineTo .hDc, &H8, .ScaleHeight - &H3
                    Arc .hDc, .ScaleWidth - &H11, .ScaleHeight - &H11, .ScaleWidth - &H2, .ScaleHeight - &H2, .ScaleWidth - &HA, .ScaleHeight - &H2, .ScaleWidth - &H1, .ScaleHeight - &HA
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    Arc .hDc, &H11, .ScaleHeight - &H12, &H1, .ScaleHeight - &H3, -&H12, .ScaleHeight, &HA, .ScaleHeight
                    MoveToEx .hDc, &H1, &H6, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H8
                    MoveToEx .hDc, .ScaleWidth - &H8, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H8, .ScaleHeight - &H4
                    Arc .hDc, .ScaleWidth - &H11, .ScaleHeight - &H11, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth - &HA, .ScaleHeight - &H2, .ScaleWidth - &H1, .ScaleHeight - &HA
                    MoveToEx .hDc, .ScaleWidth - &H3, &H6, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H8
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, &H2, &H7, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H7
                    MoveToEx .hDc, .ScaleWidth - &H4, &H7, Lines
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight - &H7
                    Arc .hDc, &H11, &H11, &H1, &H1, &H8, &H0, &H0, &H8
                    Arc .hDc, .ScaleWidth - &H11, &H11, .ScaleWidth - &H2, &H1, .ScaleWidth - &H1, &H8, .ScaleWidth - &HC, &H1
                    Arc .hDc, .ScaleWidth - &H11, &H11, .ScaleWidth - &H3, &H2, .ScaleWidth - &H1, &H8, .ScaleWidth - &HC, &H1
                    Arc .hDc, &H11, &H11, &H2, &H2, &H8, &H0, &H0, &H8
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, &HA, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H7, &H1
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, &HA, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H8, &H2
                    
                End If
                
            Case Is = Vista    '>> RoundedRectangle >> Draw Vista Style.
                
                If MouseDown Then
                    .ForeColor = BDR_VISTA2
                    RoundRect hDc, &H1, &H1, .ScaleWidth - &H2, .ScaleHeight - &H2, &HE, &HE
                Else
                    
                    .ForeColor = BDR_VISTA1
                    RoundRect hDc, &H1, &H1, .ScaleWidth - &H2, .ScaleHeight - &H2, &HE, &HE
                    
                    '>> RoundedRectangle >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        RoundRect hDc, &H1, &H1, .ScaleWidth - &H2, .ScaleHeight - &H2, &HE, &HE
                    End If
                    
                End If
                
        End Select
        
        '>> RoundedRectangle >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    .ForeColor = BDR_FOCUSRECT
                    RoundRect .hDc, &H4, &H4, .ScaleWidth - &H5, .ScaleHeight - &H5, &HF, &HF
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Round()
    
    With UserControl
        
        '>> Round >> Make Round Shape.
        'hRgn = CreateRoundRectRgn(&H0, &H0, .ScaleWidth, .ScaleHeight, .ScaleWidth , .ScaleHeight )
        hRgn = CreateEllipticRgn(&H0, &H0, .ScaleWidth, .ScaleHeight)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Round >> Draw Round Shape.
        RoundRect hDc, &H1, &H1, .ScaleWidth - &H2, .ScaleHeight - &H2, .ScaleWidth - &H2, .ScaleHeight - &H2

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Round >> Draw Visual Style.
                
                If MouseDown Then
                    .ForeColor = BDR_VISUAL
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth - &H2, &H1, &H1, .ScaleHeight - &H2
                    .ForeColor = BDR_VISUAL1
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3, &H2, &H2, .ScaleWidth - &H3, &H2, &H2, .ScaleHeight - &H3
                Else
                    .ForeColor = BDR_VISUAL1
                    Arc .hDc, &H1, &H1, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth / &H3, .ScaleHeight / &H2, .ScaleWidth - &H3, &H1
                    .ForeColor = BDR_VISUAL2
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth - &H2, &H1, &H1, .ScaleHeight - &H2
                End If
                
            Case Is = Flat                                 '>> Round >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth - &H2, &H1, &H1, .ScaleHeight - &H2
                Else
                    .ForeColor = BDR_FLAT2
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth - &H2, &H1, &H1, .ScaleHeight - &H2
                End If
                
            Case Is = OverFlat                             '>> Round >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth - &H2, &H1, &H1, .ScaleHeight - &H2
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth - &H2, &H1, &H1, .ScaleHeight - &H2
                End If
                
            Case Is = Java                                 '>> Round >> Draw Java Style.
                
                If .Enabled Then
                    .ForeColor = BDR_JAVA1
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3, &H2, &H2, .ScaleWidth / &H2, .ScaleHeight - &H3, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_JAVA2
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3, &H2, &H2, .ScaleWidth / &H2, &H2, .ScaleWidth / &H2, .ScaleHeight - &H2
                    Arc .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2, &H1, &H1, .ScaleWidth / &H2, .ScaleHeight - &H2, .ScaleWidth / &H2, &H2
                End If
                
            Case Is = WinXp                                '>> Round >> Draw WindowsXp Style.
                
                If MouseDown Then
                    .ForeColor = BDR_PRESSED
                    Arc .hDc, .ScaleWidth - &H4, .ScaleHeight - &H4, &H3, &H3, .ScaleWidth / &H2, .ScaleHeight - &H2, .ScaleWidth / &H2, &H3
                ElseIf MouseMove Then
                    .ForeColor = BDR_GOLDXP_DARK
                    Arc .hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H3, &H2, .ScaleHeight - &H3, .ScaleWidth - &H3, .ScaleHeight - &H3
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    Arc .hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H4, &H3, &H3, .ScaleWidth - &H2, &H4
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    Arc .hDc, &H3, &H3, .ScaleWidth - &H4, .ScaleHeight - &H3, &H2, &H3, .ScaleWidth - &H2, &H3
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2, &H2, &H2, .ScaleWidth - &H3, &H3, &H3, &H3
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3, &H3, &H3, .ScaleWidth - &H3, &H4, &H4, &H4
                    '>> Round >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    .ForeColor = BDR_BLUEXP_DARK
                    Arc .hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H3, &H2, .ScaleHeight - &H3, .ScaleWidth - &H3, .ScaleHeight - &H3
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    Arc .hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H4, &H3, &H3, .ScaleWidth - &H2, &H4
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    Arc .hDc, &H3, &H3, .ScaleWidth - &H4, .ScaleHeight - &H3, &H2, &H3, .ScaleWidth - &H2, &H3
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2, &H2, &H2, .ScaleWidth - &H3, &H3, &H3, &H3
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    Arc .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3, &H3, &H3, .ScaleWidth - &H3, &H4, &H4, &H4
                End If
                
            Case Is = Vista    '>> Round >> Draw Vista Style.
                
                If MouseDown Then
                    .ForeColor = BDR_VISTA2
                    RoundRect hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth - &H3, .ScaleHeight - &H3
                Else
                    
                    .ForeColor = BDR_VISTA1
                    RoundRect hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth - &H3, .ScaleHeight - &H3
                    
                    '>> Round >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        RoundRect hDc, &H2, &H2, .ScaleWidth - &H3, .ScaleHeight - &H3, .ScaleWidth - &H3, .ScaleHeight - &H3
                    End If
                    
                End If
                
        End Select
        
        '>> Round >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    .ForeColor = BDR_FOCUSRECT
                    RoundRect hDc, &H4, &H4, .ScaleWidth - &H5, .ScaleHeight - &H5, .ScaleWidth - &H5, .ScaleHeight - &H5
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Diamond()
    
    With UserControl
        
        '>> Diamond >> Make Diamond Shape.
        P(&H0).x = .ScaleWidth: P(&H0).y = .ScaleHeight / &H2
        P(&H1).x = .ScaleWidth / &H2: P(&H1).y = &H0
        P(&H2).x = &H0: P(&H2).y = .ScaleHeight / &H2
        P(&H3).x = .ScaleWidth / &H2: P(&H3).y = .ScaleHeight
        hRgn = CreatePolygonRgn(P(&H0), &H4, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Diamond >> Draw Diamond Shape.
        MoveToEx .hDc, &H0, .ScaleHeight / &H2, Lines
        LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
        LineTo .hDc, .ScaleWidth / &H2, &H0
        LineTo .hDc, &H0, .ScaleHeight / &H2
        
        If Not mButtonStyleColors = Transparent Then
            MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
            LineTo .hDc, .ScaleWidth / &H2, &H0
        End If
        
        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Diamond >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Flat                                 '>> Diamond >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                End If
                
            Case Is = OverFlat                             '>> Diamond >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                End If
                
            Case Is = Java                                 '>> Diamond >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H0, Lines
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    
                End If
                
            Case Is = WinXp                                '>> Diamond >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4, Lines
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    SetPixel .hDc, .ScaleWidth / &H2, &H1, vbBlack
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    MoveToEx .hDc, &H1, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    SetPixel .hDc, .ScaleWidth / &H2, &H1, vbBlack
                    
                    '>> Diamond >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    MoveToEx .hDc, &H1, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    SetPixel .hDc, .ScaleWidth / &H2, &H1, vbBlack
                    
                End If
                
            Case Is = Vista    '>> Diamond >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, &H1, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, &H1, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                    '>> Diamond >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, &H1, .ScaleHeight / &H2, Lines
                        LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                        LineTo .hDc, .ScaleWidth / &H2, &H1
                        LineTo .hDc, &H1, .ScaleHeight / &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Diamond >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = .ScaleWidth - &H6: PL(&H0).y = .ScaleHeight / &H2
                    PL(&H1).x = .ScaleWidth / &H2: PL(&H1).y = &H5
                    PL(&H2).x = &H5: PL(&H2).y = .ScaleHeight / &H2
                    PL(&H3).x = .ScaleWidth / &H2: PL(&H3).y = .ScaleHeight - &H6
                    Polygon .hDc, PL(&H0), &H4
                    
                    '>> Return FocusRect For Good Show.
                    For Fo.x = &H2 To .ScaleWidth - &H2 Step &H1
                        
                        For Fo.y = &H1 To .ScaleHeight - &H1 Step &H2
                            If GetPixel(.hDc, Fo.x, Fo.y) = .ForeColor Then SetPixel .hDc, Fo.x, Fo.y, vbBlack
                        Next Fo.y
                        
                    Next Fo.x
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Top_Triangle()
    
    With UserControl
        
        '>> Top_Triangle >> Make Top_Triangle Shape.
        P(&H0).x = &H0: P(&H0).y = .ScaleHeight
        P(&H1).x = .ScaleWidth / &H2: P(&H1).y = &H0
        P(&H2).x = .ScaleWidth: P(&H2).y = .ScaleHeight
        hRgn = CreatePolygonRgn(P(&H0), &H3, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Top_Triangle >> Draw Top_Triangle Shape.
        MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
        LineTo .hDc, &H1, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth / &H2, &H1

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Top_Triangle >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight - &H3
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, &H1, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Flat                                 '>> Top_Triangle >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                End If
                
            Case Is = OverFlat                             '>> Top_Triangle >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                End If
                
            Case Is = Java                                 '>> Top_Triangle >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, &H1, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H1
                    MoveToEx .hDc, &H2, .ScaleHeight - &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    
                End If
                
            Case Is = WinXp                                '>> Top_Triangle >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, &H4, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H2, &H5
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight - &H2
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, &H3, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H5, .ScaleHeight - &H4, Lines
                    LineTo .hDc, .ScaleWidth / &H2, &H5
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight - &H3
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H4, Lines
                    LineTo .hDc, &H3, .ScaleHeight - &H3
                    
                    '>> Top_Triangle >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight - &H2
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, &H3, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H5, .ScaleHeight - &H4, Lines
                    LineTo .hDc, .ScaleWidth / &H2, &H5
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight - &H3
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H4, Lines
                    LineTo .hDc, &H3, .ScaleHeight - &H3
                    
                End If
                
            Case Is = Vista    '>> Top_Triangle >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    
                    '>> Top_Triangle >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                        LineTo .hDc, &H2, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth / &H2, &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Top_Triangle >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = &H7: PL(&H0).y = .ScaleHeight - &H5
                    PL(&H1).x = .ScaleWidth / &H2: PL(&H1).y = &H9
                    PL(&H2).x = .ScaleWidth - &H8: PL(&H2).y = .ScaleHeight - &H5
                    Polygon .hDc, PL(&H0), &H3
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Left_Triangle()
    
    With UserControl
        
        '>> Left_Triangle >> Make Left_Triangle Shape.
        P(&H0).x = &H0: P(&H0).y = .ScaleHeight / &H2
        P(&H1).x = .ScaleWidth: P(&H1).y = &H0
        P(&H2).x = .ScaleWidth: P(&H2).y = .ScaleHeight
        hRgn = CreatePolygonRgn(P(&H0), &H3, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Left_Triangle >> Draw Left_Triangle Shape.
        MoveToEx .hDc, &H1, .ScaleHeight / &H2, Lines
        LineTo .hDc, .ScaleWidth - &H1, &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight - &H1
        LineTo .hDc, &H1, .ScaleHeight / &H2

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Left_Triangle >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H3, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, &H1
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                End If
                
            Case Is = Flat                                 '>> Left_Triangle >> Draw Flat Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2

                End If
                
            Case Is = OverFlat                             '>> Left_Triangle >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H1, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                End If
                
            Case Is = Java                                 '>> Left_Triangle >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth - &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H3
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H1, &H2
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight - &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                End If
                
            Case Is = WinXp                                '>> Left_Triangle >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, .ScaleWidth - &H3, &H5, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H5
                    LineTo .hDc, &H5, .ScaleHeight / &H2
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, .ScaleWidth - &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H1
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth - &H3, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H5, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, &H3
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, &H4, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H4, &H4
                    
                    '>> Left_Triangle >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, .ScaleWidth - &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H1
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth - &H3, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight - &H3
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H5, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, &H3
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, &H4, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H4, &H4
                    
                End If
                
            Case Is = Vista    '>> Left_Triangle >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    
                    '>> Left_Triangle >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, &H2, .ScaleHeight / &H2, Lines
                        LineTo .hDc, .ScaleWidth - &H2, &H2
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight - &H2
                        LineTo .hDc, &H2, .ScaleHeight / &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Left_Triangle >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = &H9: PL(&H0).y = .ScaleHeight / &H2
                    PL(&H1).x = .ScaleWidth - &H5: PL(&H1).y = &H7
                    PL(&H2).x = .ScaleWidth - &H5: PL(&H2).y = .ScaleHeight - &H8
                    Polygon .hDc, PL(&H0), &H3
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Right_Triangle()
    
    With UserControl
        
        '>> Right_Triangle >> Make Right_Triangle Shape.
        P(&H0).x = .ScaleWidth: P(&H0).y = .ScaleHeight / &H2
        P(&H1).x = &H0: P(&H1).y = &H0
        P(&H2).x = &H0: P(&H2).y = .ScaleHeight
        hRgn = CreatePolygonRgn(P(&H0), &H3, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Right_Triangle >> Draw Right_Triangle Shape.
        MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
        LineTo .hDc, &H0, &H1
        LineTo .hDc, &H0, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Right_Triangle >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H0, &H1
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H1, &H2
                    LineTo .hDc, &H1, .ScaleHeight - &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, &H1, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H0, &H1
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Flat                                 '>> Right_Triangle >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H0, &H1
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H0, &H1
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                End If
                
            Case Is = OverFlat                             '>> Right_Triangle >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H0, &H1
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H0, &H1
                    LineTo .hDc, &H0, .ScaleHeight - &H1
                End If
                
            Case Is = Java                                 '>> Right_Triangle >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, &H1, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H1, &H3
                    LineTo .hDc, &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                    
                End If
                
            Case Is = WinXp                                '>> Right_Triangle >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, &H2, &H4, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H5
                    LineTo .hDc, .ScaleWidth - &H6, .ScaleHeight / &H2
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, &H1, &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, &H3, .ScaleHeight - &H5, Lines
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H4, &H4
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H5, &H5
                    
                    '>> Right_Triangle >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, &H1, &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight - &H1
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, &H3, .ScaleHeight - &H5, Lines
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H4, &H4
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H5, &H5
                    
                End If
                
            Case Is = Vista    '>> Right_Triangle >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H1, &H2
                    LineTo .hDc, &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, &H1, &H2
                    LineTo .hDc, &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    
                    '>> Right_Triangle >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                        LineTo .hDc, &H1, &H2
                        LineTo .hDc, &H1, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Right_Triangle >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = .ScaleWidth - &HA: PL(&H0).y = .ScaleHeight / &H2
                    PL(&H1).x = &H4: PL(&H1).y = &H7
                    PL(&H2).x = &H4: PL(&H2).y = .ScaleHeight - &H8
                    Polygon .hDc, PL(&H0), &H3
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Down_Triangle()
    
    With UserControl
        
        '>> Down_Triangle >> Make Down_Triangle Shape.
        P(&H0).x = &H0: P(&H0).y = &H1
        P(&H1).x = .ScaleWidth / &H2: P(&H1).y = .ScaleHeight
        P(&H2).x = .ScaleWidth: P(&H2).y = &H1
        hRgn = CreatePolygonRgn(P(&H0), &H3, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Down_Triangle >> Draw Down_Triangle Shape.
        MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1, Lines
        LineTo .hDc, .ScaleWidth - &H1, &H1
        LineTo .hDc, &H1, &H1
        LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Down_Triangle >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H2, &H2
                    LineTo .hDc, .ScaleWidth - &H3, &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H2, &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                    
                End If
                
            Case Is = Flat                                 '>> Down_Triangle >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                End If
                
            Case Is = OverFlat                             '>> Down_Triangle >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, &H1, &H1
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                End If
                
            Case Is = Java                                 '>> Down_Triangle >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth - &H3, &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth - &H2, &H2, Lines
                    LineTo .hDc, &H2, &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H1, &H1
                    
                End If
                
            Case Is = WinXp                                '>> Down_Triangle >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, &H4, &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H5, &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H5
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, &H3, &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H5, &H4, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H3, &H3
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H5, Lines
                    LineTo .hDc, &H4, &H4
                    
                    '>> Down_Triangle >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, &H3, &H3, Lines
                    LineTo .hDc, .ScaleWidth - &H3, &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H5, &H4, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4, Lines
                    LineTo .hDc, &H3, &H3
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H5, Lines
                    LineTo .hDc, &H4, &H4
                    
                End If
                
            Case Is = Vista    '>> Down_Triangle >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    LineTo .hDc, &H2, &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, &H2
                    LineTo .hDc, &H2, &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    
                    '>> Down_Triangle >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2, Lines
                        LineTo .hDc, .ScaleWidth - &H2, &H2
                        LineTo .hDc, &H2, &H2
                        LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Down_Triangle >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = &H6: PL(&H0).y = &H4
                    PL(&H1).x = .ScaleWidth / &H2: PL(&H1).y = .ScaleHeight - &HA
                    PL(&H2).x = .ScaleWidth - &H8: PL(&H2).y = &H4
                    Polygon .hDc, PL(&H0), &H3
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Top_Arrow()
    
    With UserControl
        
        '>> Top_Arrow >> Make Top_Arrow Shape.
        P(&H0).x = .ScaleWidth / &H2: P(&H0).y = &H0
        P(&H1).x = &H0: P(&H1).y = .ScaleHeight / 1.5
        P(&H2).x = .ScaleWidth / &H3: P(&H2).y = .ScaleHeight / 1.5
        P(&H3).x = .ScaleWidth / &H3: P(&H3).y = .ScaleHeight
        P(&H4).x = .ScaleWidth / 1.5: P(&H4).y = .ScaleHeight
        P(&H5).x = .ScaleWidth / 1.5: P(&H5).y = .ScaleHeight / 1.5
        P(&H6).x = .ScaleWidth: P(&H6).y = .ScaleHeight / 1.5
        hRgn = CreatePolygonRgn(P(&H0), &H7, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Top_Arrow >> Draw Top_Arrow Shape.
        MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
        LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight - 1.5
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, .ScaleWidth / &H2, &H1

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Top_Arrow >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight / 1.5 - &H1
                    MoveToEx .hDc, &H2, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H1
                    
                Else
                    
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / &H3, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    
                End If
                
            Case Is = Flat                                 '>> Top_Arrow >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                End If
                
            Case Is = OverFlat                             '>> Top_Arrow >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H1, Lines
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                End If
                
            Case Is = Java                                 '>> Top_Arrow >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth / &H3, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / &H2, &H1
                    
                End If
                
            Case Is = WinXp                                '>> Top_Arrow >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H3, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight - 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H2, &H4
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H2
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight - 1.5 - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H2, &H5
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H2
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H5, Lines
                    LineTo .hDc, &H5, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight - &H3
                    
                    '>> Top_Arrow >> Draw Xp FocusRec.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H2
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H3
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight - 1.5 - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H2, &H5
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H3, Lines
                    LineTo .hDc, &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H2
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H5, Lines
                    LineTo .hDc, &H5, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight - &H3
                    
                End If
                
            Case Is = Vista    '>> Top_Arrow >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                    LineTo .hDc, &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H2, &H2
                    
                    '>> Top_Arrow >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, .ScaleWidth / &H2, &H2, Lines
                        LineTo .hDc, &H3, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, .ScaleWidth / &H2, &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Top_Arrow >> Draw FocusRec.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = .ScaleWidth / &H2: PL(&H0).y = &H6
                    PL(&H1).x = &H9: PL(&H1).y = .ScaleHeight / 1.5 - &H5
                    PL(&H2).x = .ScaleWidth / &H3 + &H4: PL(&H2).y = .ScaleHeight / 1.5 - &H5
                    PL(&H3).x = .ScaleWidth / &H3 + &H4: PL(&H3).y = .ScaleHeight - &H5
                    PL(&H4).x = .ScaleWidth / 1.5 - &H5: PL(&H4).y = .ScaleHeight - &H5
                    PL(&H5).x = .ScaleWidth / 1.5 - &H5: PL(&H5).y = .ScaleHeight / 1.5 - &H5
                    PL(&H6).x = .ScaleWidth - &HA: PL(&H6).y = .ScaleHeight / 1.5 - &H5
                    Polygon .hDc, PL(&H0), &H7
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Left_Arrow()
    
    With UserControl
        
        '>> Left_Arrow >> Make Left_Arrow Shape.
        P(&H0).x = &H0: P(&H0).y = .ScaleHeight / &H2
        P(&H1).x = .ScaleWidth / 1.5: P(&H1).y = &H0
        P(&H2).x = .ScaleWidth / 1.5: P(&H2).y = .ScaleHeight / &H3
        P(&H3).x = .ScaleWidth: P(&H3).y = .ScaleHeight / &H3
        P(&H4).x = .ScaleWidth: P(&H4).y = .ScaleHeight / 1.5
        P(&H5).x = .ScaleWidth / 1.5: P(&H5).y = .ScaleHeight / 1.5
        P(&H6).x = .ScaleWidth / 1.5: P(&H6).y = .ScaleHeight
        hRgn = CreatePolygonRgn(P(&H0), &H7, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Left_Arrow >> Draw Left_Arrow Shape.
        MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
        LineTo .hDc, &H1, .ScaleHeight / &H2
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Left_Arrow >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3 + &H2, Lines    'LineTo .hdc, .ScaleWidth / 1.5 - &H1, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                End If
                
            Case Is = Flat                                 '>> Left_Arrow >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                End If
                
            Case Is = OverFlat                             '>> Left_Arrow >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                End If
                
            Case Is = Java                                 '>> Left_Arrow >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight - &H1
                    LineTo .hDc, &H1, .ScaleHeight / &H2
                    
                End If
                
            Case Is = WinXp                                '>> Left_Arrow >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight - &H5
                    LineTo .hDc, &H5, .ScaleHeight / &H2
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H1
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight - &H5
                    LineTo .hDc, &H4, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth - &H4, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, &H5
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    
                    '>> Left_Arrow >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H1
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / 1.5 - &H3, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight - &H5
                    LineTo .hDc, &H4, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth - &H4, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, &H5
                    LineTo .hDc, &H3, .ScaleHeight / &H2
                    
                End If
                
            Case Is = Vista    '>> Left_Arrow >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                    LineTo .hDc, &H2, .ScaleHeight / &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3
                    
                    '>> Left_Arrow >> Draw Vista FocusRect.
                    
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H3
                        LineTo .hDc, &H2, .ScaleHeight / &H2
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight - &H3
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H3
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Left_Arrow >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = &H6: PL(&H0).y = .ScaleHeight / &H2
                    PL(&H1).x = .ScaleWidth / 1.5 - &H5: PL(&H1).y = &H9
                    PL(&H2).x = .ScaleWidth / 1.5 - &H5: PL(&H2).y = .ScaleHeight / &H3 + &H4
                    PL(&H3).x = .ScaleWidth - &H5: PL(&H3).y = .ScaleHeight / &H3 + &H4
                    PL(&H4).x = .ScaleWidth - &H5: PL(&H4).y = .ScaleHeight / 1.5 - &H5
                    PL(&H5).x = .ScaleWidth / 1.5 - &H5: PL(&H5).y = .ScaleHeight / 1.5 - &H5
                    PL(&H6).x = .ScaleWidth / 1.5 - &H5: PL(&H6).y = .ScaleHeight - &HA
                    Polygon .hDc, PL(&H0), &H7
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Right_Arrow()
    
    With UserControl
        
        '>> Right_Arrow >> Make Right_Arrow Shape.
        P(&H0).x = .ScaleWidth: P(&H0).y = .ScaleHeight / &H2
        P(&H1).x = .ScaleWidth / &H3: P(&H1).y = &H0
        P(&H2).x = .ScaleWidth / &H3: P(&H2).y = .ScaleHeight / &H3
        P(&H3).x = &H0: P(&H3).y = .ScaleHeight / &H3
        P(&H4).x = &H0: P(&H4).y = .ScaleHeight / 1.5
        P(&H5).x = .ScaleWidth / &H3: P(&H5).y = .ScaleHeight / 1.5
        P(&H6).x = .ScaleWidth / &H3: P(&H6).y = .ScaleHeight
        hRgn = CreatePolygonRgn(P(&H0), &H7, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Right_Arrow >> Draw Right_Arrow Shape.
        MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
        LineTo .hDc, &H0, .ScaleHeight / &H3
        LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
        LineTo .hDc, .ScaleWidth / &H3, &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
        LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
        LineTo .hDc, &H0, .ScaleHeight / 1.5 - &H1

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Right_Arrow >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H2
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                    
                End If
                
            Case Is = Flat                                 '>> Right_Arrow >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                End If
                
            Case Is = OverFlat                             '>> Right_Arrow >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, &H0, .ScaleHeight / 1.5 - &H1, Lines
                    LineTo .hDc, &H0, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H3, &H1
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2
                End If
                
            Case Is = Java                                 '>> Right_Arrow >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, &H0, .ScaleHeight / 1.5 - &H2
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    MoveToEx .hDc, .ScaleWidth - &H1, .ScaleHeight / &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight - &H1
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / 1.5 - &H1
                    LineTo .hDc, &H0, .ScaleHeight / 1.5 - &H1
                End If
                
            Case Is = WinXp                                '>> Right_Arrow >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, &H2, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, &H2, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight - &H5
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / &H2
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, &H2, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, &H3, .ScaleHeight / 1.5 - &H3, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight - &H5
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, &H4, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, &H5
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2
                    
                    '>> Right_Arrow >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, &H2, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, &H2, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, &H3, .ScaleHeight / 1.5 - &H3, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / 1.5 - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight - &H5
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H2
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, &H4, .ScaleHeight / &H3 + &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, &H5
                    LineTo .hDc, .ScaleWidth - &H4, .ScaleHeight / &H2
                    
                End If
                
            Case Is = Vista    '>> Right_Arrow >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2 - &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H2
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                    LineTo .hDc, &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                    LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2 - &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                    LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H2
                    
                    '>> Right_Arrow >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, &H1, .ScaleHeight / 1.5 - &H2, Lines
                        LineTo .hDc, &H1, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, &H3
                        LineTo .hDc, .ScaleWidth - &H2, .ScaleHeight / &H2 - &H1
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight - &H3
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / 1.5 - &H2
                        LineTo .hDc, &H1, .ScaleHeight / 1.5 - &H2
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Right_Arrow >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = .ScaleWidth - &H7: PL(&H0).y = .ScaleHeight / &H2
                    PL(&H1).x = .ScaleWidth / &H3 + &H4: PL(&H1).y = &H9
                    PL(&H2).x = .ScaleWidth / &H3 + &H4: PL(&H2).y = .ScaleHeight / &H3 + &H4
                    PL(&H3).x = &H4: PL(&H3).y = .ScaleHeight / &H3 + &H4
                    PL(&H4).x = &H4: PL(&H4).y = .ScaleHeight / 1.5 - &H5
                    PL(&H5).x = .ScaleWidth / &H3 + &H4: PL(&H5).y = .ScaleHeight / 1.5 - &H5
                    PL(&H6).x = .ScaleWidth / &H3 + &H4: PL(&H6).y = .ScaleHeight - &HA
                    Polygon .hDc, PL(&H0), &H7
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub Redraw_Down_Arrow()
    
    With UserControl
        
        '>> Down_Arrow >> Make Down_Arrow Shape.
        P(&H0).x = .ScaleWidth / &H2: P(&H0).y = .ScaleHeight
        P(&H1).x = &H0: P(&H1).y = .ScaleHeight / &H3
        P(&H2).x = .ScaleWidth / &H3: P(&H2).y = .ScaleHeight / &H3
        P(&H3).x = .ScaleWidth / &H3: P(&H3).y = &H0
        P(&H4).x = .ScaleWidth / 1.5: P(&H4).y = &H0
        P(&H5).x = .ScaleWidth / 1.5: P(&H5).y = .ScaleHeight / &H3
        P(&H6).x = .ScaleWidth: P(&H6).y = .ScaleHeight / &H3
        hRgn = CreatePolygonRgn(P(&H0), &H7, WINDING)
        SetWindowRgn .hWnd, hRgn, True
        DeleteObject hRgn

        '>> Down_Arrow >> Draw Down_Arrow Shape.
        MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
        LineTo .hDc, .ScaleWidth / &H3, &H0
        LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
        LineTo .hDc, &H1, .ScaleHeight / &H3
        LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
        LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
        LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H0

        Select Case mButtonStyle
            
            Case Is = Visual                               '>> Down_Arrow >> Draw Visual Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISUAL
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
                    LineTo .hDc, .ScaleWidth / &H3, &H0
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    
                Else
                    
                    .ForeColor = BDR_VISUAL1
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H2, &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    MoveToEx .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_VISUAL2
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
                    LineTo .hDc, .ScaleWidth / &H3, &H0
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    
                End If
                
            Case Is = Flat                                 '>> Down_Arrow >> Draw Flat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
                    LineTo .hDc, .ScaleWidth / &H3, &H0
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                Else
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
                    LineTo .hDc, .ScaleWidth / &H3, &H0
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                End If
                
            Case Is = OverFlat                             '>> Down_Arrow >> Draw OverFlat Style.
                
                If MouseDown Then
                    .ForeColor = BDR_FLAT1
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
                    LineTo .hDc, .ScaleWidth / &H3, &H0
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                ElseIf MouseMove Then
                    .ForeColor = BDR_FLAT2
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H0, Lines
                    LineTo .hDc, .ScaleWidth / &H3, &H0
                    LineTo .hDc, .ScaleWidth / &H3, .ScaleHeight / &H3
                    LineTo .hDc, &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                End If
                
            Case Is = Java                                 '>> Down_Arrow >> Draw Java Style.
                
                If .Enabled Then
                    
                    .ForeColor = BDR_JAVA1
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H2, &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_JAVA2
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 + &H1, &H1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H1, &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth - &H1, .ScaleHeight / &H3
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H1
                    
                End If
                
            Case Is = WinXp                                '>> Down_Arrow >> Draw WindowsXp Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_PRESSED
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H3, &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, &H2
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    
                ElseIf MouseMove Then
                    
                    .ForeColor = BDR_GOLDXP_DARK
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    .ForeColor = BDR_GOLDXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    .ForeColor = BDR_GOLDXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H3, &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    .ForeColor = BDR_GOLDXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_GOLDXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H2, &H3, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, &H4, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    
                    '>> Down_Arrow >> Draw Xp FocusRect.
                ElseIf mButtonStyle = WinXp And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_BLUEXP_DARK
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H1, &H1
                    .ForeColor = BDR_BLUEXP_NORMAL1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H1, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H2
                    .ForeColor = BDR_BLUEXP_NORMAL2
                    MoveToEx .hDc, .ScaleWidth / 1.5 - &H3, &H2, Lines
                    LineTo .hDc, .ScaleWidth / 1.5 - &H3, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth - &H5, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    .ForeColor = BDR_BLUEXP_LIGHT1
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H1, &H2, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H3
                    .ForeColor = BDR_BLUEXP_LIGHT2
                    MoveToEx .hDc, .ScaleWidth / &H3 + &H2, &H3, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H2, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, &H4, .ScaleHeight / &H3 + &H2
                    LineTo .hDc, .ScaleWidth / &H2, .ScaleHeight - &H4
                    
                End If
                
            Case Is = Vista    '>> Down_Arrow >> Draw Vista Style.
                
                If MouseDown Then
                    
                    .ForeColor = BDR_VISTA2
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2 - &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H1
                    
                Else
                    
                    .ForeColor = BDR_VISTA1
                    MoveToEx .hDc, .ScaleWidth / 1.5, &H1, Lines
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, &H1
                    LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / &H2 - &H1, .ScaleHeight - &H2
                    LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                    LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H1
                    
                    '>> Down_Arrow >> Draw Vista FocusRect.
                    If mButtonStyle = Vista And mFocusRect And mFocused Then
                        
                        .ForeColor = BDR_FOCUSRECT_VISTA
                        MoveToEx .hDc, .ScaleWidth / 1.5, &H1, Lines
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, &H1
                        LineTo .hDc, .ScaleWidth / &H3 + &H1, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, &H3, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, .ScaleWidth / &H2 - &H1, .ScaleHeight - &H2
                        LineTo .hDc, .ScaleWidth - &H3, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, .ScaleHeight / &H3 + &H1
                        LineTo .hDc, .ScaleWidth / 1.5 - &H2, &H1
                        
                    End If
                    
                End If
                
        End Select
        
        '>> Down_Arrow >> Draw FocusRect.
        If Not mButtonStyle = Java And mFocusRect And mFocused Then
            
            If Not mButtonStyle = WinXp And mFocusRect And mFocused Then
                
                If Not mButtonStyle = Vista And mFocusRect And mFocused Then
                    
                    .ForeColor = BDR_FOCUSRECT
                    PL(&H0).x = .ScaleWidth / &H2: PL(&H0).y = .ScaleHeight - &H7
                    PL(&H1).x = &H9: PL(&H1).y = .ScaleHeight / &H3 + &H4
                    PL(&H2).x = .ScaleWidth / &H3 + &H4: PL(&H2).y = .ScaleHeight / &H3 + &H4
                    PL(&H3).x = .ScaleWidth / &H3 + &H4: PL(&H3).y = &H4
                    PL(&H4).x = .ScaleWidth / 1.5 - &H5: PL(&H4).y = &H4
                    PL(&H5).x = .ScaleWidth / 1.5 - &H5: PL(&H5).y = .ScaleHeight / &H3 + &H4
                    PL(&H6).x = .ScaleWidth - &HA: PL(&H6).y = .ScaleHeight / &H3 + &H4
                    Polygon .hDc, PL(&H0), &H7
                    
                End If
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub UserControl_Resize()
    Call UserControl_Paint
End Sub

Private Sub UserControl_Show()
    Call UserControl_Paint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    '>> Load Saved Property Values
    'On Local Error Resume Next
    
    With PropBag
        
        mButtonShape = .ReadProperty("ButtonShape", Rectangle)
        mButtonStyle = .ReadProperty("ButtonStyle", Visual)
        mButtonStyleColors = .ReadProperty("ButtonStyleColors", SingleColor)
        mButtonTheme = .ReadProperty("ButtonTheme", NoTheme)
        mButtonType = .ReadProperty("ButtonType", Button)
        mCaptionAlignment = .ReadProperty("CaptionAlignment", CenterCaption)
        mCaptionEffect = .ReadProperty("CaptionEffect", Default)
        mCaptionStyle = .ReadProperty("CaptionStyle", Normal)
        mDropDown = .ReadProperty("DropDown", None)
        mPictureAlignment = .ReadProperty("PictureAlignment", CenterPicture)
        
        mBackColor = .ReadProperty("BackColor", vbButtonFace)
        mBackColorPressed = .ReadProperty("BackColorPressed", vbButtonFace)
        mBackColorHover = .ReadProperty("BackColorHover", vbButtonFace)
        
        mBorderColor = .ReadProperty("BorderColor", vbBlack)
        mBorderColorPressed = .ReadProperty("BorderColorPressed", vbBlack)
        mBorderColorHover = .ReadProperty("BorderColorHover", vbBlack)
        
        mForeColor = .ReadProperty("ForeColor", vbBlack)
        mForeColorPressed = .ReadProperty("ForeColorPressed", vbRed)
        mForeColorHover = .ReadProperty("ForeColorHover", vbBlue)
        
        mEffectColor = .ReadProperty("EffectColor", vbWhite)
        
        Caption = .ReadProperty("Caption", Ambient.DisplayName)
        mTagExtra = .ReadProperty("TagExtra", VBA.vbNullString)
        UserControl.AccessKeys = .ReadProperty("AccessKey", VBA.vbNullString)
        
        mFocusRect = .ReadProperty("FocusRect", True)
        mValue = .ReadProperty("Value", False)
        mHandPointer = .ReadProperty("HandPointer", False)
        
        Set mPicture = .ReadProperty("Picture", Nothing)
        mPictureGray = .ReadProperty("PictureGray", False)
        
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set Font = .ReadProperty("Font", Ambient.Font)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", &H0)
        
    End With
    
End Sub

Private Sub UserControl_Terminate()
    DeleteObject hRgn                                      '>> Remove The Region From Memory
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    '>> Save Property Values
    'On Local Error Resume Next
    
    With PropBag
        
        .WriteProperty "ButtonShape", mButtonShape, Rectangle
        .WriteProperty "ButtonStyle", mButtonStyle, Visual
        .WriteProperty "ButtonStyleColors", mButtonStyleColors, SingleColor
        .WriteProperty "ButtonTheme", mButtonTheme, NoTheme
        .WriteProperty "ButtonType", mButtonType, Button
        .WriteProperty "CaptionAlignment", mCaptionAlignment, CenterCaption
        .WriteProperty "CaptionEffect", mCaptionEffect, Default
        .WriteProperty "CaptionStyle", mCaptionStyle, Normal
        .WriteProperty "DropDown", mDropDown, None
        .WriteProperty "PictureAlignment", mPictureAlignment, CenterPicture
        
        .WriteProperty "BackColor", mBackColor, vbButtonFace
        .WriteProperty "BackColorPressed", mBackColorPressed, vbButtonFace
        .WriteProperty "BackColorHover", mBackColorHover, vbButtonFace
        
        .WriteProperty "BorderColor", mBorderColor, vbBlack
        .WriteProperty "BorderColorPressed", mBorderColorPressed, vbBlack
        .WriteProperty "BorderColorHover", mBorderColorHover, vbBlack
        
        .WriteProperty "ForeColor", mForeColor, vbBlack
        .WriteProperty "ForeColorPressed", mForeColorPressed, vbRed
        .WriteProperty "ForeColorHover", mForeColorHover, vbBlue
        
        .WriteProperty "EffectColor", mEffectColor, vbWhite
        
        .WriteProperty "Caption", mCaption, UserControl.Name
        .WriteProperty "TagExtra", mTagExtra, VBA.vbNullString
        .WriteProperty "AccessKey", UserControl.AccessKeys, VBA.vbNullString
        
        .WriteProperty "FocusRect", mFocusRect, True
        .WriteProperty "Value", mValue, False
        .WriteProperty "HandPointer", mHandPointer, False
        
        .WriteProperty "Picture", mPicture, Nothing
        .WriteProperty "PictureGray", mPictureGray, False
        
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, &H0
        
    End With
    
End Sub

