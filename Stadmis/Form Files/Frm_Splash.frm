VERSION 5.00
Begin VB.Form Frm_Splash 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   0  'None
   Caption         =   "App Title :Splash Screen"
   ClientHeight    =   4320
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Frm_Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerExpiry 
      Interval        =   200
      Left            =   -600
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -600
      Top             =   2760
   End
   Begin VB.Frame Fra_Splash 
      BackColor       =   &H00FFFFFF&
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   7560
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   650
         Left            =   -960
         Top             =   2400
      End
      Begin Stadmis.ShapeButton ShpBttnSystemInfo 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5040
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "      System Info"
         AccessKey       =   "S"
         Picture         =   "Frm_Splash.frx":0ECA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Stadmis.ShapeButton ShpBttnOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   6360
         TabIndex        =   17
         Top             =   3840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   7
         ButtonStyleColors=   3
         PictureAlignment=   1
         BackColor       =   14211288
         BackColorPressed=   15715986
         BackColorHover  =   16243621
         BorderColor     =   9408398
         BorderColorPressed=   6045981
         BorderColorHover=   11632444
         Caption         =   "    OK"
         AccessKey       =   "O"
         Picture         =   "Frm_Splash.frx":1464
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image ImgDBPhoto 
         Height          =   855
         Left            =   4800
         Stretch         =   -1  'True
         Tag             =   "HW"
         ToolTipText     =   "Masika .S. Elvas"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Shape ShpOutline 
         Height          =   1095
         Left            =   4680
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image ImgVirtualPhoto 
         Height          =   1350
         Left            =   7680
         Picture         =   "Frm_Splash.frx":19FE
         Top             =   2640
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Shape ShpImgOutline 
         BorderColor     =   &H00404040&
         Height          =   615
         Left            =   6720
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   16
         Tag             =   "AutoSizer:Y"
         Top             =   4080
         Width           =   210
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max Users:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   15
         Tag             =   "AutoSizer:Y"
         Top             =   4080
         Width           =   930
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ABCDE-FGHIJ-KLMNO-PQRST-UVWXY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Tag             =   "AutoSizer:Y"
         Top             =   3600
         Width           =   3060
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licence Code:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Tag             =   "AutoSizer:Y"
         Top             =   3600
         Width           =   1140
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "09/03/1986"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Tag             =   "AutoSizer:Y"
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Tag             =   "AutoSizer:Y"
         Top             =   3840
         Width           =   1020
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   1320
         TabIndex        =   10
         Tag             =   "AutoSizer:Y"
         Top             =   4080
         Width           =   420
      End
      Begin VB.Label LblLicenseTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licence Key:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "AutoSizer:Y"
         Top             =   4080
         Width           =   1035
      End
      Begin VB.Label lblDeveloper 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Developer details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   192
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1248
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   7560
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frm_Splash.frx":25DC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   5100
      End
      Begin VB.Image imgLogo 
         Height          =   1065
         Left            =   120
         Picture         =   "Frm_Splash.frx":26DD
         Stretch         =   -1  'True
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LblTradeMark 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TradeMark"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6480
         TabIndex        =   6
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label LblBuild 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Build:001"
         ForeColor       =   &H00000080&
         Height          =   192
         Left            =   6720
         TabIndex        =   5
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label LblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   6180
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   6720
         Picture         =   "Frm_Splash.frx":30C7
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image ImgHeader 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   10
         Picture         =   "Frm_Splash.frx":3F91
         Stretch         =   -1  'True
         Top             =   120
         Width           =   7510
      End
      Begin VB.Label lblSchoolName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   2355
      End
      Begin VB.Label LblCopyright 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6570
         TabIndex        =   1
         Top             =   3280
         Width           =   810
      End
      Begin VB.Label LblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   5910
         TabIndex        =   2
         Top             =   2040
         Width           =   1470
      End
      Begin VB.Image ImgFooter 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   0
         Picture         =   "Frm_Splash.frx":5126
         Stretch         =   -1  'True
         Top             =   3525
         Width           =   7550
      End
      Begin VB.Image ImgEagleSystems 
         Height          =   615
         Left            =   6720
         Picture         =   "Frm_Splash.frx":62BB
         Stretch         =   -1  'True
         ToolTipText     =   "Eagle Systems"
         Top             =   2520
         Width           =   615
      End
      Begin VB.Shape ShpSplash 
         BackColor       =   &H80000013&
         BorderColor     =   &H00808080&
         FillColor       =   &H00CFE1E2&
         FillStyle       =   0  'Solid
         Height          =   3000
         Left            =   0
         Top             =   600
         Width           =   7545
      End
   End
   Begin VB.Image ImgDeveloper 
      Height          =   4005
      Left            =   0
      Picture         =   "Frm_Splash.frx":6E99
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "Frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     maselv_e@yahoo.co.uk / masika_elvas@programmer.net / masika_elvas@live.com  *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************

Option Explicit

Public HinderTransparency As Boolean

Private FadingIn As Boolean

' Reg Key Security Options...
Const KEY_NOTIFY = &H10
Const KEY_SET_VALUE = &H2
Const KEY_QUERY_VALUE = &H1
Const KEY_CREATE_LINK = &H20
Const READ_CONTROL = &H20000
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
             
' Reg Key ROOT Types...
Const REG_SZ = &H1                         ' Unicode nul terminated string
Const REG_DWORD = &H4                      ' 32-bit number
Const ERROR_SUCCESS = &H0
Const HKEY_LOCAL_MACHINE = &H80000002

Const gREGVALSYSINFO = "PATH"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"

Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Public Sub StartSysInfo()
On Local Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        
        ' Validate Existance Of Known 32 Bit File Version
        If (VBA.Dir(SysInfoPath & "\MSINFO32.EXE") <> VBA.vbNullString) Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
        Else ' Error - File Can Not Be Found...
            GoTo SysInfoErr
        End If
        
    Else ' Error - Registry Entry Can Not Be Found...
        GoTo SysInfoErr
    End If
    
    Call VBA.Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub 'Quit this Procedure
    
SysInfoErr:
    
    'Inform User
    vMsgBox "System Information Is Unavailable At This Time", vbOKOnly, App.Title, Me
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    
    vIndex(&H1) = RegOpenKeyEx(KeyRoot, KeyName, &H0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (vIndex(&H1) <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    vBuffer(&H0) = VBA.String$(1024, &H0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    
    ' Get/Create Key Value
    vIndex(&H1) = RegQueryValueEx(hKey, SubKeyRef, &H0, KeyValType, vBuffer(&H0), KeyValSize)
                        
    If (vIndex(&H1) <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (VBA.Asc(VBA.Mid(vBuffer(&H0), KeyValSize, &H1)) = &H0) Then           ' Win95 Adds Null Terminated String...
        vBuffer(&H0) = VBA.Left(vBuffer(&H0), KeyValSize - &H1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        vBuffer(&H0) = VBA.Left(vBuffer(&H0), KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    
    Select Case KeyValType                                  ' Search Data Types...
        
        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = vBuffer(&H0)
            
            ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
            
            For vIndex(&H0) = VBA.Len(vBuffer(&H0)) To &H1 Step -&H1                      ' Convert Each Bit
                KeyVal = KeyVal + VBA.Hex(VBA.Asc(VBA.Mid(vBuffer(&H0), vIndex(&H0), &H1)))   ' Build Value Char. By Char.
            Next vIndex(&H0)
            
            KeyVal = VBA.Format$("&h" + KeyVal)                     ' Convert Double Word To String
            
    End Select 'Close SELECT..CASE block statement
    
    GetKeyValue = True                                      ' Return Success
    vIndex(&H1) = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError: ' Cleanup After An Error Has Occured...
    
    KeyVal = VBA.vbNullString                               ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    vIndex(&H1) = RegCloseKey(hKey)                         ' Close Registry Key
    
End Function

Private Function ApplyTheme() As Boolean
    
    '**************************************************************
    'Apply Theme Settings
    
    Me.BackColor = tTheme.tBackColor
    Fra_Splash.BackColor = Me.BackColor
    ShpSplash.FillColor = Me.BackColor
    
    ImgHeader.Picture = tTheme.tImagePicture
    ImgFooter.Picture = ImgHeader.Picture
    
    ShpBttnOK.ForeColor = tTheme.tButtonForeColor
    ShpBttnOK.BackColor = tTheme.tButtonBackColor
    
    ShpBttnSystemInfo.ForeColor = tTheme.tButtonForeColor
    ShpBttnSystemInfo.BackColor = tTheme.tButtonBackColor
    
    lblSchoolName.ForeColor = tTheme.tForeColor
    LblProductName.ForeColor = tTheme.tForeColor
    lblDeveloper.ForeColor = tTheme.tForeColor
    lblDisclaimer.ForeColor = tTheme.tForeColor
    
    LblVersion.ForeColor = tTheme.tForeColor
    LblBuild.ForeColor = tTheme.tForeColor
    
    LblCopyright.ForeColor = tTheme.tForeColor
    LblTradeMark.ForeColor = tTheme.tForeColor
    
    For vIndex(&H0) = &H0 To LblLicenseTo.UBound Step &H2
        LblLicenseTo(vIndex(&H0)).ForeColor = tTheme.tForeColor
    Next vIndex(&H0)
    
    For vIndex(&H0) = &H1 To LblLicenseTo.UBound Step &H2
        LblLicenseTo(vIndex(&H0)).ForeColor = tTheme.tEntryColor
    Next vIndex(&H0)
    
    '**************************************************************
    
End Function

Private Sub ShpBttnOK_Click()
    Call Fra_Splash_Click
End Sub

Private Sub ShpBttnRegister_Click()
    
    Me.Enabled = False
    
    vRegistered = True
    Frm_SoftwarePatent.Show , Me
    vRegistered = False
    
    LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < 61 And Not vRegistered, True, False)
    
    LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
    
    LblLicenseTo(&H1).Caption = SoftwareSetting.Licences.License_Code
    LblLicenseTo(&H5).Caption = SoftwareSetting.Licences.Key
    LblLicenseTo(&H7).Caption = SoftwareSetting.Licences.Max_Users '(SoftwareSetting.Licences.Max_Users / &HA)
    
    Call TimerExpiry_Timer
    Me.Enabled = True: ShpBttnOK.Refresh
    
End Sub

Private Sub ShpBttnSystemInfo_Click()
  Call StartSysInfo
End Sub

Private Sub Form_Activate()
    
    'If the image control contains the Developer's Photo then...
    If ImgDeveloper.Picture <> &H0 And HinderTransparency Then
        
        'Display Developer's Photo
        ImgDeveloper.Refresh
        
        'Call Procedure to Fit image to image holder
        Call FitPicTo(ImgDeveloper, ImgDBPhoto, ShpOutline)
        
    End If 'Close respective IF..THEN block statement
    
    ShpBttnOK.Visible = HinderTransparency
    ShpBttnSystemInfo.Visible = HinderTransparency
    lblDeveloper.Visible = HinderTransparency
    
    Me.Refresh
    
End Sub

Private Sub Form_Load()
    
    VBA.Randomize
    
    If Not HinderTransparency Then
        
        Transparency Me, &H0
        Call WindowsMinimizeAll 'Minimize all open Windows
        Me.Caption = App.Title
        
    Else
        Timer1.Enabled = False
    End If 'Close respective IF..THEN block statement
    
    Me.Caption = VBA.IIf(HinderTransparency, "About ", VBA.vbNullString) & App.Title
    imgLogo.Picture = Me.Icon
    
    'Call ApplyTheme
    
    FadingIn = True
    
    lblSchoolName.Caption = App.ProductName
    LblProductName.Caption = App.FileDescription
    LblVersion.Caption = "Version " & App.Major & "." & App.Minor
    LblBuild.Caption = "Build " & VBA.Format$(App.Revision, "0000")
    LblCopyright.Caption = App.LegalCopyright
    LblTradeMark.Caption = App.LegalTrademarks
    
    lblDeveloper.Caption = "Developed by:" & VBA.vbCrLf & _
                            "Masika .S. Elvas" & VBA.vbCrLf & _
                            "P.O Box 137, BUNGOMA 50200, KENYA" & VBA.vbCrLf & _
                            "(254)724 688 172 / (254)751 041 184" & VBA.vbCrLf & _
                            "elvasmasika@lexeme-kenya.com"
    
    LblLicenseTo(&H1).Caption = VBA.IIf(School.Name = VBA.vbNullString, App.CompanyName, School.Name & VBA.vbCrLf & School.Location)
    
    LblLicenseTo(&H1).Caption = SoftwareSetting.Licences.License_Code
    LblLicenseTo(&H5).Caption = SoftwareSetting.Licences.Key
    LblLicenseTo(&H7).Caption = SoftwareSetting.Licences.Max_Users
    
    Call TimerExpiry_Timer
    
    LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < 61 And Not vRegistered, True, False)
    
    LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
    
End Sub

Private Sub Fra_Photo_Click()
    Call ImgDBPhoto_Click
End Sub

Private Sub Fra_Splash_Click()
On Local Error Resume Next
    If Not HinderTransparency Then Frm_Login.Show: Timer1.Enabled = False: Timer2.Enabled = False
    Unload Me
End Sub

Private Sub Fra_Splash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If HinderTransparency Then FormDrag Me
End Sub

Private Sub ImgDBPhoto_Click()
    ImgDeveloper.Picture = ImgDBPhoto.Picture
    ImgDeveloper.ToolTipText = ImgDBPhoto.ToolTipText
    Call PhotoClicked(ImgDBPhoto, ImgDeveloper, ShpImgOutline, True)
End Sub

Private Sub ImgEagleSystems_Click()
    
    If HinderTransparency Then
        ImgVirtualPhoto.Picture = ImgEagleSystems.Picture
        ImgVirtualPhoto.ToolTipText = ImgEagleSystems.ToolTipText
        Call PhotoClicked(ImgEagleSystems, ImgVirtualPhoto, ShpImgOutline, True)
    Else
        Call Fra_Splash_Click
    End If
    
End Sub

Private Sub ImgHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If HinderTransparency Then FormDrag Me
End Sub

Private Sub LblSchoolName_Click()
    Call Fra_Splash_Click
End Sub

Private Sub lblDisclaimer_Click()
    Call Fra_Splash_Click
End Sub

Private Sub LblLicenseTo_DblClick(Index As Integer)
On Local Error Resume Next
    If Index = &H2 Then VBA.DeleteSetting App.Title, "Copyright Protection": LblLicenseTo(&H3).ForeColor = &HFF&
End Sub

Private Sub lblProductName_Click()
    Call Fra_Splash_Click
End Sub

Private Sub LblVersion_Click()
    Call Fra_Splash_Click
End Sub

Private Sub Timer1_Timer()
On Local Error GoTo Err
    
    Static Index%
    
    Index = VBA.IIf(FadingIn = True, Index + &H5, Index - &H5)
    Transparency Me, Index
    Exit Sub
    
Err:
    
    If FadingIn = False Then Call Fra_Splash_Click: Exit Sub
    FadingIn = Not FadingIn: Timer1.Enabled = False: Timer2.Enabled = True
    
End Sub

Private Sub Timer2_Timer()
    
    Static vCounter&
    
    vArrayList = VBA.Split(" ", " ")
    
    If vCounter >= UBound(vArrayList) + &H1 Then Timer2.Enabled = False: Timer1.Enabled = True: Exit Sub
    
    Sleep VBA.Int((2500 * VBA.Rnd) + 400)
    VBA.DoEvents
    vCounter = vCounter + &H1
    
End Sub

Private Sub TimerExpiry_Timer()
    LblLicenseTo(&H3).Caption = VBA.Format$(SoftwareSetting.Licences.Expiry_Date, "ddd dd MMM yyyy hh:nn:ss AMPM") & " - Remaining " & VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) - &H1 & " day" & VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) - &H1 = &H1, "", "s")
    LblLicenseTo(&H3).FontBold = (VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < &H3)
End Sub
