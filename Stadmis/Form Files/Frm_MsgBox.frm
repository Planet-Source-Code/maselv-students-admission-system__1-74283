VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_MsgBox 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "App Title : Message Box"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_MsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerAutoClose 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   -600
      Top             =   1920
   End
   Begin VB.Frame Fra_Details 
      BackColor       =   &H00CFE1E2&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Tag             =   "AutoSizer:WH"
      Top             =   360
      Width           =   3975
      Begin VB.TextBox TxtPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE1E2&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "AutoSizer:WH"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label LblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "!!!-WARNING-!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Image MsgBoxImg 
         Height          =   735
         Left            =   120
         Picture         =   "Frm_MsgBox.frx":0ECA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LblPrompt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1020
      End
   End
   Begin MSComctlLib.ImageList MsgBoxImgLst 
      Left            =   -840
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_MsgBox.frx":130C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_MsgBox.frx":1626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_MsgBox.frx":1940
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_MsgBox.frx":22C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox ChkDontShowMsg 
      BackColor       =   &H00CFE1E2&
      Caption         =   "Do not display this message again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "AutoSizer:Y"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Stadmis.ShapeButton ShpBttnButton 
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   1800
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
      Caption         =   "OK"
      AccessKey       =   "O"
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
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_MsgBox.frx":25DA
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Image ImgHeader 
      Height          =   255
      Left            =   0
      Picture         =   "Frm_MsgBox.frx":2DD0
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "Frm_MsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     maselv_e@yahoo.co.uk / elvasmasika@lexeme-kenya.com / masika_elvas@live.com *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************

Option Explicit
    
    Public iMsg$
    Public DisplayDuration&
    Public ParentFrm As Object
    Public DisplayWarning As Boolean
    Public iWarningType As vMsgBoxButtons
    
    Private TimeDisplayed$
    Private mHasSoundCard As Boolean
    Private vCounter&, vDefaultBttn&, vMousePointer&
    
Private Declare Function waveOutGetNumDevs Lib "winmm" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public Function PlayInfoSound(mWarningType As vMsgBoxButtons) As Boolean
    
    If Not SoftwareSetting.SoftwareSound Then Exit Function
    
    Screen.MousePointer = vbHourglass 'Change Mouse Pointer to show Processing state
    
    Dim mSoundFile$
    Dim wFlags%, mVar%
    
    'Check the type of message being displayed
    Select Case mWarningType
        
        Case vbCritical: mSoundFile = "vMsgBox - Critical.wav"
        Case vbExclamation: mSoundFile = "vMsgBox - Warning.wav"
        Case vbQuestion: mSoundFile = "vMsgBox - Question.wav"
        Case Else: mSoundFile = "vMsgBox - Notify.wav"
        
    End Select 'Close SELECT..CASE block statement
    
    Dim iFso As New FileSystemObject
    
    'Check if Software Media Files' Folder has not been specified then set to default
    If SoftwareSetting.MsgBoxSoundFldr = VBA.vbNullString Then SoftwareSetting.MsgBoxSoundFldr = App.Path & "\Tools\Media"
    
    'Check if Software Media Files' Folder has not been created in the Computer the create it
    If Not iFso.FolderExists(SoftwareSetting.MsgBoxSoundFldr) Then CreatePath SoftwareSetting.MsgBoxSoundFldr
    
    'If an appropriate sound has not been identified then Quit this Function
    If mSoundFile = VBA.vbNullString Then Exit Function
    
    PlayInfoSound = PlaySound(Me, SoftwareSetting.MsgBoxSoundFldr & "\" & mSoundFile)
    
    Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
    
End Function

Private Sub ChkDontShowMsg_KeyPress(KeyAscii As Integer)
    Call TxtPrompt_KeyPress(KeyAscii)
End Sub

Private Sub ShpBttnButton_Click(Index As Integer)
On Local Error Resume Next
    
    vArrayList = VBA.Split(ShpBttnButton(Index).Tag, ":")
    vSelectedButton(&H0) = vArrayList(&H0)
    vSelectedButton(&H1) = vArrayList(&H1)
    vSelectedButton(&H2) = vArrayList(&H2)
    TimerAutoClose.Enabled = False
    Unload Me
    
End Sub

Private Sub ShpBttnButton_KeyPress(Index As Integer, KeyAscii As Integer)
    Call TxtPrompt_KeyPress(KeyAscii)
End Sub

Private Sub Form_Activate()
On Local Error Resume Next
    
    vMousePointer = Screen.MousePointer
    
    If Not vShowMsgBox Then Unload Me: Exit Sub
    
    SleepEx &H64, &H1 'Delay for 100 milliseconds
    
    TimeDisplayed = VBA.Now
    
    LblWarning.Visible = DisplayWarning
    
    Me.Caption = VBA.IIf(Me.Caption = "App Title : Message Box", App.Title, Me.Caption)
    
    PlayInfoSound iWarningType
    
    vDefaultBttn = VBA.Val(Me.Tag) - &H1
    ShpBttnButton(vDefaultBttn).Default = True:
    ShpBttnButton(vDefaultBttn).SetFocus
    
    vArrayList = VBA.Split(ShpBttnButton(vDefaultBttn).Tag, ":")
    vSelectedButton(&H0) = vArrayList(&H0)
    vSelectedButton(&H1) = vArrayList(&H1)
    vSelectedButton(&H2) = vArrayList(&H2)
    
    Me.Icon = MsgBoxImgLst.ListImages((iWarningType / &H10)).ExtractIcon
    
    If ChkDontShowMsg.Visible Then ChkDontShowMsg.SetFocus Else ShpBttnButton(vDefaultBttn).SetFocus
    
    iMsgBoxDisplayed = True
    BringWindowToTop Me.hWnd
    
    If ChkDontShowMsg.Visible Then ChkDontShowMsg.SetFocus
    
    'Change Mouse Pointer to show end of Processing state
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    'Call ApplyTheme
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If ChkDontShowMsg.Visible Then
        
        DontShowMsgAgain(VBA.Val(ChkDontShowMsg.DataField), &H0) = ChkDontShowMsg.Value
        DontShowMsgAgain(VBA.Val(ChkDontShowMsg.DataField), &H1) = vSelectedButton(&H0)
        
    End If
    
    Dim iStrObj$, iStrN&
    Dim nArray() As String
    
    nArray = VBA.Split(Me.Caption, " - ")
    If UBound(nArray) > &H0 Then iStrN = VBA.Val(nArray(&H1)) 'Error Number
    
    If Not Nothing Is ParentFrm Then iStrObj = ParentFrm.Name
    Call EventLog(iStrObj, VBA.CStr(iWarningType), Err.Source, iStrN, TxtPrompt.Text, vSelectedButton(&H2))
    
    vShowMsgBox = False: TimerAutoClose.Enabled = False: vCounter = DisplayDuration + &H1
    
    'Unload created buttons
    For vIndex(&H0) = &H1 To ShpBttnButton.UBound
        Unload ShpBttnButton(vIndex(&H0))
    Next vIndex(&H0)
    
    iMsgBoxDisplayed = False 'Denote that the message box has not ben
    
    'Change Mouse Pointer to show end of Processing state
    Screen.MousePointer = vMousePointer
    
End Sub

Private Sub TimerAutoClose_Timer()
    
    If TimeDisplayed = VBA.vbNullString Then Exit Sub
    
    DisplayDuration = VBA.IIf(DisplayDuration = &H0, &H5, DisplayDuration)
    vCounter = VBA.DateDiff("s", TimeDisplayed, VBA.Now)
    TxtPrompt.Text = VBA.Replace(iMsg, "£", DisplayDuration - vCounter)
    ShpBttnButton(&H0).Caption = "(" & DisplayDuration - vCounter & ")"
    If vCounter > DisplayDuration Then TimerAutoClose.Enabled = False: Unload Me
    
End Sub

Private Sub TxtPrompt_GotFocus()
On Local Error Resume Next
    ShpBttnButton(&H0).SetFocus
End Sub

Private Sub TxtPrompt_KeyPress(KeyAscii As Integer)
    
    Screen.MousePointer = vbHourglass 'Change Mouse Pointer to show Processing state
    
    For vIndex(&H0) = &H0 To ShpBttnButton.UBound Step &H1
        
        If KeyAscii = vbKeyReturn Then Call ShpBttnButton_Click(VBA.CInt(VBA.Val(Me.Tag))): Exit For
        
        KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii)))
        
        vArrayList = VBA.Split(ShpBttnButton(vIndex(&H0)).Tag, ":")
        If UBound(vArrayList) >= &H1 Then If KeyAscii = VBA.Asc(VBA.UCase$(vArrayList(&H1))) Then Call ShpBttnButton_Click(VBA.CInt(vIndex(&H0))): Exit For
        
    Next vIndex(&H0)
    
    Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
    
End Sub
