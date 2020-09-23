VERSION 5.00
Begin VB.Form Frm_PleaseWait 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   0  'None
   Caption         =   "App Title : Please Wait..."
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "Frm_PleaseWait.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
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
      Caption         =   "      Cancel"
      AccessKey       =   "C"
      Picture         =   "Frm_PleaseWait.frx":09EA
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
   Begin Stadmis.ShapeButton ShpBttnPause 
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   5
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
      Caption         =   "      Pause"
      AccessKey       =   "P"
      Picture         =   "Frm_PleaseWait.frx":0D84
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
   Begin VB.Label lblProgressBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2AD96&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Tag             =   "AutoSizer:XY"
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Image ImgProgressBar 
      Height          =   375
      Left            =   240
      Picture         =   "Frm_PleaseWait.frx":111E
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:Y"
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Shape ShpBorder 
      BorderColor     =   &H00808080&
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0% Complete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1890
      Width           =   1155
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2220
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading data..."
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1230
   End
   Begin VB.Image ImgHeader 
      Height          =   1095
      Left            =   0
      Picture         =   "Frm_PleaseWait.frx":1A31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
   Begin VB.Image ImgFooter 
      Height          =   1215
      Left            =   0
      Picture         =   "Frm_PleaseWait.frx":22D1
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   5295
   End
End
Attribute VB_Name = "Frm_PleaseWait"
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
'=  SHOWS THE PROGRESS OF A PARTICULAR PROCESS IS SESSION

Option Explicit

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private iStartTime$

Private Sub Form_Activate()
    iStartTime = VBA.Now
End Sub

Private Sub ShpBttnCancel_Click()
    
    'Confirm if the User really wants to cancel the on-going process. If not then Quit this Procedure
    If vMsgBox("Are you sure you want to this Process?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then Exit Sub
    
    vCancelOperation = &H2 'Denote that the process has been cancelled
    
End Sub

Private Sub ShpBttnPause_Click()
    vCancelOperation = VBA.IIf(ShpBttnPause.Caption = "Pause", &H1, &H0)
    ShpBttnPause.Caption = VBA.IIf(ShpBttnPause.Caption = "Continue", "Pause", "Continue")
End Sub

Private Sub Form_Load()
    
    'Set the Form to be always on top of other open windows
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    lblProgressBar.Width = 4815
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'If the wait duration is less than 5 seconds then don't play any sound
    If VBA.DateDiff("s", iStartTime, VBA.Now) < &H5 Then Exit Sub
    
    'Check if Software Media Files' Folder has not been specified then set to default
    If SoftwareSetting.MsgBoxSoundFldr = VBA.vbNullString Then SoftwareSetting.MsgBoxSoundFldr = App.Path & "\Tools\Media"
    
    'Check if Software Media Files' Folder has not been created in the Computer the create it
    If Not vFso.FolderExists(SoftwareSetting.MsgBoxSoundFldr) Then CreatePath SoftwareSetting.MsgBoxSoundFldr
    
    'If 100% complete then...
    If ImgProgressBar.Width = lblProgressBar.Width Then
        Call PlaySound(Me, SoftwareSetting.MsgBoxSoundFldr & "\Progress - Complete.wav")
    Else 'If not 100% complete then...
        Call PlaySound(Me, SoftwareSetting.MsgBoxSoundFldr & "\Progress - Error.wav")
    End If
    
    VBA.DoEvents 'Yield execution so that the operating system can process other events
    'Sleep 250 'Wait for quarter a second
    
End Sub
