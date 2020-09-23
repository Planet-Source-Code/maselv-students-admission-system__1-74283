VERSION 5.00
Begin VB.Form Frm_BriefNotes 
   BackColor       =   &H00CFE1E2&
   Caption         =   "App Title : Brief Info"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   Icon            =   "Frm_BriefNotes.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
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
      Picture         =   "Frm_BriefNotes.frx":09EA
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
      Left            =   4680
      TabIndex        =   3
      Top             =   2760
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
      Picture         =   "Frm_BriefNotes.frx":0D84
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
   Begin Stadmis.AutoSizer AutoSizer 
      Left            =   120
      Top             =   2760
      _extentx        =   661
      _extenty        =   661
   End
   Begin VB.Frame fraBriefNotes 
      BackColor       =   &H00CFE1E2&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Tag             =   "AutoSizer:WH"
      Top             =   600
      Width           =   5535
      Begin VB.TextBox txtBriefNotes 
         BackColor       =   &H00C0FFFF&
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Tag             =   "AutoSizer:WH"
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Information"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2520
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   -120
      Picture         =   "Frm_BriefNotes.frx":131E
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Image ImgHeader 
      Height          =   840
      Left            =   0
      Picture         =   "Frm_BriefNotes.frx":1B14
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "Frm_BriefNotes"
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

Private Sub Form_Activate()
    ShpBttnCancel.Visible = Not txtBriefNotes.Locked
    txtBriefNotes.ForeColor = VBA.IIf(txtBriefNotes.Locked, &H80&, &HFF0000)
    txtBriefNotes.BackColor = VBA.IIf(txtBriefNotes.Locked, fraBriefNotes.BackColor, &HC0FFFF)
End Sub

Private Sub Form_Initialize()
    Me.Caption = App.Title & " : Brief Info"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     If vBuffer(&H0) = VBA.vbNullString Then Cancel = Not CloseFrm(Me, , False)
End Sub

Private Sub ShpBttnCancel_Click()
    Unload Me
End Sub

Private Sub ShpBttnOK_Click()
    vBuffer(&H0) = VBA.IIf(txtBriefNotes.Locked, vBuffer(&H0), txtBriefNotes.Text): Unload Me
End Sub
