VERSION 5.00
Begin VB.Form Frm_School 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : School Information"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6255
   Icon            =   "Frm_School.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSchoolMotto 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Tag             =   "WY"
      Top             =   4680
      Width           =   5775
   End
   Begin VB.TextBox txtEmailAddress 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2880
      Width           =   3735
   End
   Begin VB.TextBox txtPhoneNo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Frame Fra_Logo 
      BackColor       =   &H00CFE1E2&
      Height          =   1935
      Left            =   240
      TabIndex        =   33
      Tag             =   "XY"
      Top             =   720
      Width           =   1935
      Begin VB.Image ImgDBLogo 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Tag             =   "HW"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image ImgVirtualLogo 
         Height          =   255
         Left            =   240
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.TextBox txtWebsite 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Tag             =   "WY"
      Top             =   4080
      Width           =   5775
   End
   Begin VB.TextBox txtPostalAddress 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2280
      MaxLength       =   255
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "WY"
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Frame Fra_School 
      BackColor       =   &H00CFE1E2&
      Height          =   5895
      Left            =   120
      TabIndex        =   34
      Tag             =   "HW"
      Top             =   600
      Width           =   6015
      Begin VB.TextBox txtAbbreviation 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkModule 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Hide Inactive Modules"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   31
         Tag             =   "Y"
         Top             =   5400
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkModule 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Fees Module"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   30
         Tag             =   "Y"
         Top             =   5400
         Value           =   2  'Grayed
         Width           =   1215
      End
      Begin VB.CheckBox chkModule 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Examination Module"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Tag             =   "Y"
         Top             =   5400
         Value           =   2  'Grayed
         Width           =   1815
      End
      Begin VB.ComboBox cboStudentGender 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "Frm_School.frx":076A
         Left            =   4080
         List            =   "Frm_School.frx":0777
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox txtNSSFNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         MaxLength       =   30
         TabIndex        =   23
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox txtNHIFNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   25
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox cboSchoolType 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "Frm_School.frx":0796
         Left            =   4320
         List            =   "Frm_School.frx":07A3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtSchoolCode 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   17
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtSchoolName 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox chkDiscontinued 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Discontinued"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Tag             =   "Y"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1332
      End
      Begin Stadmis.ShapeButton ShpBttnClearLogo 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2160
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
         Caption         =   "      Clear"
         AccessKey       =   "C"
         Picture         =   "Frm_School.frx":07CF
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
      Begin Stadmis.ShapeButton ShpBttnAttachLogo 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   2160
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
         Caption         =   "      Attach"
         AccessKey       =   "A"
         Picture         =   "Frm_School.frx":0B69
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
      Begin VB.Label lblAbbreviation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abbreviation:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3240
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  'Dot
         Height          =   735
         Left            =   120
         Top             =   5040
         Width           =   5775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Software Modules:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   5160
         Width           =   1560
      End
      Begin VB.Label lblStudentGender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Gender:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4080
         TabIndex        =   26
         Top             =   4440
         Width           =   1200
      End
      Begin VB.Label lblNSSFNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NSSF No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   4440
         Width           =   675
      End
      Begin VB.Label lblNHIFNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NHIF No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label lblSchoolType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblSchoolMotto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Tag             =   "Y"
         Top             =   3840
         Width           =   480
      End
      Begin VB.Label lblSchoolCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Code:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2160
         TabIndex        =   16
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label lblPhoneNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   750
      End
      Begin VB.Label lblEmailAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2160
         TabIndex        =   12
         Tag             =   "Y"
         Top             =   2040
         Width           =   1110
      End
      Begin VB.Label lblPostalAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Tag             =   "Y"
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label lblSchoolName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "School Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2160
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Tag             =   "Y"
         Top             =   3240
         Width           =   645
      End
   End
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   4800
      TabIndex        =   37
      Top             =   6720
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
      Caption         =   "      Brief Notes"
      AccessKey       =   "B"
      Picture         =   "Frm_School.frx":0F03
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
      Picture         =   "Frm_School.frx":129D
      Stretch         =   -1  'True
      Tag             =   "WY"
      Top             =   6600
      Width           =   6375
   End
   Begin VB.Label lblRecords 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Records: #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Tag             =   "Y"
      Top             =   6270
      Width           =   1395
   End
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Information"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   420
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   3165
   End
   Begin VB.Image ImgHeader 
      Height          =   615
      Left            =   0
      Picture         =   "Frm_School.frx":1A93
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
   Begin VB.Menu MnuSave 
      Caption         =   "&Save"
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
   End
End
Attribute VB_Name = "Frm_School"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     elvasmasika@lexeme-kenya.com\masika_elvas@live.com                          *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************

Option Explicit

Public IsNewRecord As Boolean

Private myTable$, myTablePryKey$
Private myRecIndex&, TxtKeyBack&, iSetNo&
Private myTableFixedFldName(&H5) As String
Private myRecDisplayON, IsLoading, mySchoolCodeChange, mySchoolTypeChange As Boolean

Public Function ClearEntries() As Boolean
On Local Error GoTo Handle_ClearEntries_Error
    
    Dim MousePointerState%
    Dim myRecDisplayState As Boolean
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayState = myRecDisplayON
    myRecDisplayON = True
    
    mySchoolCodeChange = False: mySchoolTypeChange = False
    Fra_School.Tag = VBA.vbNullString
    txtSchoolCode.Tag = VBA.vbNullString
    txtSchoolCode.Text = VBA.vbNullString
    cboSchoolType.Tag = VBA.vbNullString
    txtSchoolName.Text = VBA.vbNullString
    txtSchoolMotto.Text = VBA.vbNullString
    txtPostalAddress.Text = VBA.IIf(myRecDisplayON, VBA.vbNullString, "P.O Box ")
    txtEmailAddress.Text = VBA.vbNullString
    txtPhoneNo.Text = VBA.IIf(myRecDisplayON, VBA.vbNullString, "254")
    txtWebsite.Text = VBA.vbNullString
    txtLocation.Text = VBA.vbNullString
    txtNSSFNo.Text = VBA.vbNullString
    txtNHIFNo.Text = VBA.vbNullString
    ShpBttnBriefNotes.TagExtra = VBA.vbNullString
    chkDiscontinued.Value = vbUnchecked
    chkModule(&H0).Value = vbChecked
    chkModule(&H1).Value = vbChecked
    chkModule(&H2).Value = vbChecked
    
    'Clear Picture
    ImgVirtualLogo.Picture = Nothing
    ImgVirtualLogo.ToolTipText = VBA.vbNullString
    ImgDBLogo.Picture = Nothing
    ImgDBLogo.ToolTipText = VBA.vbNullString
    Erase sAdditionalPhoto(&H0).vDataBytes
    
Exit_ClearEntries:
    
    myRecDisplayON = myRecDisplayState
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ClearEntries_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Clearing Entries - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ClearEntries
    
End Function

Private Function LockEntries(State As Boolean) As Boolean
On Local Error GoTo Handle_LockEntries_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_School.Enabled = Not State
    txtSchoolMotto.Locked = State
    txtPhoneNo.Locked = State
    txtPostalAddress.Locked = State
    txtEmailAddress.Locked = State
    txtWebsite.Locked = State
    
    If Not State Then
        
        txtSchoolCode.Enabled = ValidUserAccess(Me, iSetNo, &H7, , False)
        txtSchoolName.Enabled = ValidUserAccess(Me, iSetNo, &H1, , False)
        cboSchoolType.Enabled = ValidUserAccess(Me, iSetNo, &H8, , False)
        cboStudentGender.Enabled = ValidUserAccess(Me, iSetNo, &H9, , False)
        
    Else
        
        txtSchoolCode.Enabled = True
        txtSchoolName.Enabled = True
        cboSchoolType.Enabled = True
        cboStudentGender.Enabled = True
        
    End If
    
    MnuSave.Visible = Not State: MnuEdit.Visible = State
    
Exit_LockEntries:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_LockEntries_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Locking Entries - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_LockEntries
    
End Function

Public Function DisplayRecord(vRecordID&) As Boolean
On Local Error GoTo Handle_DisplayRecord_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayON = True 'Denote that Record display process is in progress
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM " & myTable & " WHERE [School ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'Assign the Record's primary key value
            If Not VBA.IsNull(![School ID]) Then Fra_School.Tag = ![School ID]
            If Not VBA.IsNull(![School Code]) Then txtSchoolCode.Text = ![School Code]: txtSchoolCode.Tag = ![School Code]
            If Not VBA.IsNull(![School Name]) Then txtSchoolName.Text = ![School Name]
'            If Not VBA.IsNull(![School Type]) Then cboSchoolType.ListIndex = ![School Type]: cboSchoolType.Tag = ![School Type]
            If Not VBA.IsNull(![Abbreviation]) Then txtAbbreviation.Text = ![Abbreviation]
'            If Not VBA.IsNull(![Student Gender]) Then cboStudentGender.ListIndex = ![School Type]: cboStudentGender.Tag = ![Student Gender]
            If Not VBA.IsNull(![School Motto]) Then txtSchoolMotto.Text = ![School Motto]
            If Not VBA.IsNull(![Postal Address]) Then txtPostalAddress.Text = ![Postal Address]
            If Not VBA.IsNull(![E-mail Address]) Then txtEmailAddress.Text = ![E-mail Address]
            If Not VBA.IsNull(![Phone No]) Then txtPhoneNo.Text = ![Phone No]
            If Not VBA.IsNull(![Website]) Then txtWebsite.Text = ![Website]
            If Not VBA.IsNull(![Location]) Then txtLocation.Text = ![Location]
            If Not VBA.IsNull(![NSSF No]) Then txtNSSFNo.Text = ![NSSF No]
            If Not VBA.IsNull(![NHIF No]) Then txtNHIFNo.Text = ![NHIF No]
            If Not VBA.IsNull(![Brief Notes]) Then ShpBttnBriefNotes.TagExtra = ![Brief Notes]
            chkDiscontinued.Value = VBA.IIf(![Discontinued], vbChecked, vbUnchecked)
            
'            If Not VBA.IsNull(![Software Modules]) Then
'
'                chkModule(&H0).Value = VBA.IIf(VBA.InStr(";" & ![Software Modules] & ";", ";1;") <> &H0, vbChecked, vbUnchecked)
'                chkModule(&H1).Value = VBA.IIf(VBA.InStr(";" & ![Software Modules] & ";", ";2;") <> &H0, vbChecked, vbUnchecked)
'                chkModule(&H2).Value = VBA.IIf(VBA.InStr(";" & ![Software Modules] & ";", ";3;") <> &H0, vbChecked, vbUnchecked)
'
'            End If 'Close respective IF..THEN block statement
            
            'If the Record contains the School's Logo then...
            If Not VBA.IsNull(![Logo]) Then
                
                'Display School's Logo
                sAdditionalPhoto(&H0).vDataBytes = ![Logo]
                Set ImgVirtualLogo.DataSource = Nothing
                ImgVirtualLogo.DataField = "Logo"
                Set ImgVirtualLogo.DataSource = vRs
                ImgVirtualLogo.Refresh
                ImgVirtualLogo.ToolTipText = txtSchoolName.Text & "'s Logo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualLogo, ImgDBLogo, Fra_Logo)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_DisplayRecord:
    
    myRecDisplayON = False 'Denote that Record display process is complete
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_DisplayRecord_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Displaying Record - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplayRecord
    
End Function

Private Sub chkDiscontinued_Click()
On Local Error GoTo Handle_chkDiscontinued_Click_Error
    
    If myRecDisplayON Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If chkDiscontinued.Value = vbChecked Then If vMsgBox("Ticking this option will disable the School's Record and will not be available in other Modules. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then chkDiscontinued.Value = vbUnchecked
    
Exit_chkDiscontinued_Click:
    
    myRecDisplayON = False 'Denote that Record display process is complete
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_chkDiscontinued_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Discontinuing Record - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_chkDiscontinued_Click
    
End Sub

Private Sub chkModule_Click(Index As Integer)
On Local Error GoTo Handle_chkModule_Click_Error
    
    If myRecDisplayON Then Exit Sub
    
    If Not ValidUserAccess(Me, iSetNo, &H5 + Index) Then myRecDisplayON = True: chkModule(Index).Value = vbChecked: myRecDisplayON = False: Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If chkModule(Index).Value = vbUnchecked Then If vMsgBox("Ticking this option will disable the " & chkModule(Index).Caption & " and will not be available in the System. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then myRecDisplayON = True: chkModule(Index).Value = vbChecked: myRecDisplayON = False
    
Exit_chkModule_Click:
    
    myRecDisplayON = False 'Denote that Record display process is complete
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_chkModule_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Discontinuing Record - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_chkModule_Click
    
End Sub

'Executed when this Form becomes the Active Form
Private Sub Form_Activate()
On Local Error GoTo Handle_Form_Activate_Error
    
    If Not IsLoading Then Exit Sub 'If the Form has already loaded then Quit this Procedure
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    IsLoading = False 'Denote that the Form has been loaded on the Memory
    
    IsNewRecord = True 'Denote that a new Record is to be entered
    
    'Set defaults
    iSetNo = &H1: chkModule(&H2).Visible = (User.Hierarchy = &H0)
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to clear all entries in Input Boxes
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records in the specified database table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'Get the total number of Records already saved
        lblRecords.Tag = .RecordCount
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        'If a Record is to be displayed then...
        If vEditRecordID <> VBA.vbNullString Then
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            'Retrieve the Record with the specified ID
            .Filter = "[School ID] = " & vArrayList(&H0)
            
            'If the record Exists then Call Procedure in this Form to display it
            If Not (.BOF And .EOF) Then DisplayRecord VBA.CLng(![School ID])
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            vEditRecordID = VBA.vbNullString  'Initialize variable
            
            If Not (.BOF And .EOF) Then If UBound(vArrayList) = &H1 Then Call MnuEdit_Click 'Call click event of Edit Menu
            If Not (.BOF And .EOF) Then If UBound(vArrayList) = &H2 Then vMsgBox "Only editting allowed for the existing School. Operation Aborted."
            
            'Reinitialize the elements of the fixed-size array and release dynamic-array storage space.
            Erase vArrayList
            
        Else 'If a new Record is to be entered then...
            
            'If there is an
            If .RecordCount > &H0 Then DisplayRecord VBA.CLng(![School ID]) Else Call NewRecord
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_Form_Activate:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Form_Activate_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Activating Form - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Form_Activate
    
End Sub

Private Sub Form_Load()
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    IsNewRecord = True: IsLoading = True
    
    myTableFixedFldName(&H0) = "School": myTableFixedFldName(&H1) = "School"
    myTableFixedFldName(&H2) = "School": myTableFixedFldName(&H3) = "School"
    myTableFixedFldName(&H4) = "Tbl_School": myTableFixedFldName(&H5) = "Tbl_Students:Students"
    
    myTable = "Tbl_" & myTableFixedFldName(&H1)
    
    Me.Caption = App.Title & " : " & myTableFixedFldName(&H3) & " Information"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub Fra_Logo_Click()
    Call ImgDBLogo_Click
End Sub

Private Sub ImgDBLogo_Click()
    Call PhotoClicked(ImgDBLogo, ImgVirtualLogo, Fra_Logo, Not (Fra_School.Enabled))
End Sub

Private Sub MnuEdit_Click()
    
    If Not ValidUserAccess(Me, iSetNo, &H2) Then Exit Sub
    
    'If the User has not selected any existing Record then...
    If Fra_School.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & "No Record Displayed"
        If Fra_School.Enabled Then txtSchoolName.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    LockEntries False 'Call Procedure in this Form to UnLock Input Controls
    
    IsNewRecord = False 'Denote that the displayed Record exists in the database
    
    If Fra_School.Enabled Then If txtSchoolName.Enabled Then txtSchoolName.SetFocus Else txtAbbreviation.SetFocus   'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Function NewRecord() As Boolean
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    Call LockEntries(False)  'Call Procedure in this Form to UnLock Input Controls
    
    IsNewRecord = True 'Denote that the displayed Record does not exist in the database
    
    txtSchoolName.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Function

Private Sub MnuSave_Click()
On Local Error GoTo Handle_MnuSave_Click_Error
    
    'If the User has not entered the School's Code then...
    If VBA.LenB(VBA.Trim$(txtSchoolCode.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the School's Code", vbExclamation, App.Title & " : Code not entered"
        txtSchoolCode.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the School's Name abbreviation then...
    If VBA.LenB(VBA.Trim$(txtAbbreviation.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the School's Name Abbreviation", vbExclamation, App.Title & " : Abbreviation not entered"
        txtAbbreviation.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the School's Name then...
    If VBA.LenB(VBA.Trim$(txtSchoolName.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the School's Name", vbExclamation, App.Title & " : Name not entered"
        txtSchoolName.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered numeric School's Phone Numbers then...
    If Not VBA.IsNumeric(VBA.Replace(VBA.Replace(txtPhoneNo.Text, VBA.vbCrLf, VBA.vbNullString), " ", VBA.vbNullString)) And txtPhoneNo.Text <> VBA.vbNullString Then
        
        'Warn User
        vMsgBox "Please enter numeric Phone Numbers", vbExclamation, App.Title & " : Invalid Phone Number"
        txtPhoneNo.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Start the saving process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    '---------------------------------------------------------------------------------------------------
    'Format entry appropriately
    '---------------------------------------------------------------------------------------------------
    
    txtSchoolName.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtSchoolName.Text, "  ", " ")))
    txtSchoolMotto.Text = VBA.Trim$(VBA.Replace(txtSchoolMotto.Text, "  ", " "))
    txtPostalAddress.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtPostalAddress.Text, "  ", " ")))
    txtPhoneNo.Text = VBA.Replace(txtPhoneNo.Text, " ", VBA.vbNullString)
    txtLocation.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtLocation.Text, "  ", " ")))
    txtEmailAddress.Text = VBA.LCase$(VBA.Trim$(VBA.Replace(txtEmailAddress.Text, " ", VBA.vbNullString)))
    txtWebsite.Text = VBA.LCase$(VBA.Trim$(VBA.Replace(txtWebsite.Text, " ", VBA.vbNullString)))
    
    '---------------------------------------------------------------------------------------------------
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[School ID] = " & VBA.Val(Fra_School.Tag): .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        ![School Code] = txtSchoolCode.Text
        ![Abbreviation] = txtAbbreviation.Text
        ![School Name] = txtSchoolName.Text
        ![School Motto] = txtSchoolMotto.Text
'        ![School Type] = cboSchoolType.ListIndex
'        ![Software Modules] = VBA.IIf(chkModule(&H0).Value = vbChecked, &H1, VBA.vbNullString) & VBA.IIf(chkModule(&H1).Value = vbChecked, VBA.IIf(chkModule(&H0).Value = vbChecked, ";", VBA.vbNullString) & &H2, VBA.vbNullString) & VBA.IIf(chkModule(&H2).Value = vbChecked, VBA.IIf(chkModule(&H0).Value = vbChecked Or chkModule(&H1).Value = vbChecked, ";", VBA.vbNullString) & &H3, VBA.vbNullString)
'        ![Student Gender] = cboStudentGender.ListIndex
        ![Postal Address] = VBA.IIf(VBA.Trim(txtPostalAddress.Text) = "P.O Box", VBA.vbNullString, txtPostalAddress.Text)
        ![E-mail Address] = txtEmailAddress.Text
        ![Phone No] = VBA.IIf(VBA.Trim(txtPhoneNo.Text) = "254", VBA.vbNullString, txtPhoneNo.Text)
        ![Website] = txtWebsite.Text
        ![Location] = txtLocation.Text
        ![NSSF No] = txtNSSFNo.Text
        ![NHIF No] = txtNHIFNo.Text
        
        ![Logo] = Null
        
        'If the School has a Logo then Assign the Logo to its field
        If ImgDBLogo.Picture <> &H0 Then .Fields("Logo").AppendChunk sAdditionalPhoto(&H0).vDataBytes
        
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![Discontinued] = chkDiscontinued.Value
'        ![Date Last Modified] = VBA.FormatDateTime(VBA.Date, vbShortDate)
        ![User ID] = User.User_ID
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        Fra_School.Tag = ![School ID]
        
        VBA.SaveSetting App.Title, "Settings", "School Name", SmartEncrypt(txtSchoolName.Text)
        
        'If the saved details are for the School in session then...
        If ![School ID] = School.ID Then
            
            'Update details in the System
            
            If Not VBA.IsNull(![School ID]) Then School.ID = ![School ID]
            If Not VBA.IsNull(![School Code]) Then School.code = ![School Code]
            If Not VBA.IsNull(![School Name]) Then School.Name = ![School Name]
'            If Not VBA.IsNull(![School Type]) Then School.Type = ![School Type]
'            If Not VBA.IsNull(![Software Modules]) Then School.Modules = ![Software Modules]
'            If Not VBA.IsNull(![Student Gender]) Then School.Gender = ![Student Gender]
            If Not VBA.IsNull(![School Motto]) Then School.Motto = ![School Motto]
            If Not VBA.IsNull(![Phone No]) Then School.PhoneNo = ![Phone No]
            If Not VBA.IsNull(![Location]) Then School.Location = ![Location]
            If Not VBA.IsNull(![E-mail Address]) Then School.EmailAddress = ![E-mail Address]
            If Not VBA.IsNull(![Postal Address]) Then School.PostalAddress = ![Postal Address]
            If Not VBA.IsNull(![Website]) Then School.Website = ![Website]
            If Not VBA.IsNull(![Logo]) Then Set School.Logo = ImgDBLogo.Picture Else Set School.Logo = Nothing
            
            Frm_Main.lblSchoolInfo(&H0).Caption = School.Name
            Frm_Main.lblSchoolInfo(&H1).Caption = School.PostalAddress & "    Phone: " & VBA.Replace(School.PhoneNo, VBA.vbCrLf, " / ")
            Frm_Main.lblSchoolInfo(&H2).Caption = "E-mail: " & VBA.Replace(School.EmailAddress, VBA.vbCrLf, " / ")
            Frm_Main.lblSchoolInfo(&H3).Caption = "Website: " & School.Website
            Set Frm_Main.ImgVirtualPhoto.Picture = School.Logo
            
            'Reset last selected Category to default
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", "0:0"
            
            If chkModule(&H2).Value = vbChecked And User.Hierarchy <> &H0 Then
                
'                Frm_Main.MnuExaminations.Visible = (VBA.InStr(";" & School.Modules & ";", ";1;") <> &H0)
'                Frm_Main.MnuSchoolFees.Visible = (VBA.InStr(";" & School.Modules & ";", ";2;") <> &H0)
                
            Else
                
'                Frm_Main.MnuExaminations.Visible = True: Frm_Main.MnuSchoolFees.Visible = True
                
            End If 'Close respective IF..THEN block statement
            
            Frm_Main.Caption = App.Title & " : " & School.Name & " - " & App.FileDescription
            
            'Call Procedure to Fit image to image holder
            Call FitPicTo(Frm_Main.ImgVirtualPhoto, Frm_Main.imgLogo, Frm_Main.ShpOutline)
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
        'Denote that the database table has been altered
        vDatabaseAltered = True
        
        'Get the total number of Records already saved
        lblRecords.Tag = VBA.Val(lblRecords.Tag) + VBA.IIf(IsNewRecord, &H1, &H0)
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        mySchoolCodeChange = False: mySchoolTypeChange = False
        LockEntries True 'Call Procedure in this Form to Lock Input Controls
        
    End With 'Close the WITH block statements
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "The Record has successfully been " & VBA.IIf(IsNewRecord, "Saved", "Modified"), vbInformation, App.Title & " : Saving Report"
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    IsNewRecord = True 'Cancel Edit Mode
    
Exit_MnuSave_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_MnuSave_Click_Error:
    
    'If Err is 'Invalid pointer error' while saving Logo then...
    If Err.Number = -2147287031# Then
        
        'Clear Picture
        ImgVirtualLogo.Picture = Nothing
        ImgVirtualLogo.ToolTipText = VBA.vbNullString
        ImgDBLogo.Picture = Nothing
        ImgDBLogo.ToolTipText = VBA.vbNullString
        Erase sAdditionalPhoto(&H0).vDataBytes
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Warn User
        vMsgBox "Unable to Attach the image to the Record. {" & Err.Description & "}", vbExclamation, App.Title & " : Error - " & Err.Number, Me
        
        'Indicate that a process or operation is in progress.
        Screen.MousePointer = vbHourglass
        
        Resume Next 'Execute the next line of code
        
    End If 'Close respective IF..THEN block statement
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Saving Record - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuSave_Click
    
End Sub

Private Sub ShpBttnAttachLogo_Click()
    Call PhotoClicked(ImgDBLogo, ImgVirtualLogo, Fra_Logo, Not (Fra_School.Enabled), "0")
End Sub

Private Sub ShpBttnBriefNotes_Click()
    'Call Function in Mdl_Stadmis to display Notes Input Form
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_School.Enabled, myTableFixedFldName(&H2))
End Sub

Private Sub ShpBttnClearLogo_Click()
    Call PhotoClicked(ImgDBLogo, ImgVirtualLogo, Fra_Logo, Not (Fra_School.Enabled), "1")
End Sub

Private Sub txtAbbreviation_KeyPress(KeyAscii As Integer)
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))) 'Force Capital Letters
    If KeyAscii = vbKeySpace Then KeyAscii = Empty 'Negate spaces
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii))) 'Force Small Letters
    If KeyAscii = vbKeySpace Then KeyAscii = Empty 'Negate spaces
End Sub

Private Sub txtLocation_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtLocation.Text = VBA.vbNullString Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtLocation, "[Tbl_School]", "Location", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
    If txtLocation.Text <> txtLocation.SelText Then myRecDisplayON = True: txtLocation.Text = VBA.Replace(txtLocation.Text, txtLocation.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtNHIFNo_KeyPress(KeyAscii As Integer)
    'Force CAPS and discard spaces
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))): If KeyAscii = vbKeySpace Then KeyAscii = Empty
End Sub

Private Sub txtNSSFNo_KeyPress(KeyAscii As Integer)
    'Force CAPS and discard spaces
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))): If KeyAscii = vbKeySpace Then KeyAscii = Empty
End Sub

Private Sub txtPhoneNo_GotFocus()
    txtPhoneNo.SelStart = VBA.LenB(txtPhoneNo.Text)
End Sub

Private Sub txtPostalAddress_GotFocus()
    txtPostalAddress.SelStart = VBA.LenB(txtPostalAddress.Text)
End Sub

Private Sub txtPostalAddress_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtPostalAddress.Text = VBA.vbNullString Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtPostalAddress, "[Tbl_School]", "Postal Address", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtPostalAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtPostalAddress_Validate(Cancel As Boolean)
    If txtPostalAddress.Text <> txtPostalAddress.SelText Then myRecDisplayON = True: txtPostalAddress.Text = VBA.Replace(txtPostalAddress.Text, txtPostalAddress.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtWebsite_KeyPress(KeyAscii As Integer)
    KeyAscii = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii)))
    If KeyAscii = vbKeySpace Then KeyAscii = Empty
End Sub
