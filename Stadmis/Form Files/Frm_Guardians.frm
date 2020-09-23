VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Guardians 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Guardians"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6255
   Icon            =   "Frm_Guardians.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Guardians.frx":09EA
   ScaleHeight     =   5670
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   2400
      TabIndex        =   35
      Top             =   5100
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
      Picture         =   "Frm_Guardians.frx":0D2C
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
   Begin VB.CommandButton CmdMoveRec 
      BackColor       =   &H00CFCFCF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5760
      Picture         =   "Frm_Guardians.frx":10C6
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "XY"
      ToolTipText     =   "Move Last"
      Top             =   5100
      Width           =   375
   End
   Begin VB.CommandButton CmdMoveRec 
      BackColor       =   &H00CFCFCF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5400
      Picture         =   "Frm_Guardians.frx":1408
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "XY"
      ToolTipText     =   "Move Next"
      Top             =   5100
      Width           =   375
   End
   Begin VB.CommandButton CmdMoveRec 
      BackColor       =   &H00CFCFCF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5040
      Picture         =   "Frm_Guardians.frx":174A
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "XY"
      ToolTipText     =   "Move Previous"
      Top             =   5100
      Width           =   375
   End
   Begin VB.CommandButton CmdMoveRec 
      BackColor       =   &H00CFCFCF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4680
      Picture         =   "Frm_Guardians.frx":1A8C
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "XY"
      ToolTipText     =   "Move First"
      Top             =   5100
      Width           =   375
   End
   Begin VB.TextBox txtOccupation 
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
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Tag             =   "WY"
      Top             =   4440
      Width           =   3135
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
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Tag             =   "WY"
      Top             =   3240
      Width           =   3495
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
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Tag             =   "WY"
      Top             =   3840
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
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00CFE1E2&
      Height          =   1935
      Left            =   240
      TabIndex        =   40
      Tag             =   "XY"
      Top             =   480
      Width           =   1935
      Begin VB.Image ImgDBPhoto 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Tag             =   "HW"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image ImgVirtualPhoto 
         Height          =   255
         Left            =   240
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Fra_Guardian 
      BackColor       =   &H00CFE1E2&
      Height          =   4575
      Left            =   120
      TabIndex        =   41
      Tag             =   "HW"
      Top             =   360
      Width           =   6015
      Begin VB.CheckBox chkDiscontinued 
         BackColor       =   &H00CFE1E2&
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
         Left            =   4680
         TabIndex        =   34
         Tag             =   "Y"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CheckBox chkAutoComplete 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   4080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   2880
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   2880
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkDeceased 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Deceased"
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
         Left            =   3600
         TabIndex        =   33
         Tag             =   "Y"
         Top             =   4080
         Width           =   1095
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   24
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox cboMaritalStatus 
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
         Height          =   315
         ItemData        =   "Frm_Guardians.frx":1DCE
         Left            =   3120
         List            =   "Frm_Guardians.frx":1DD8
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtNationalID 
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
         MaxLength       =   8
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cboGender 
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
         Height          =   315
         ItemData        =   "Frm_Guardians.frx":1DED
         Left            =   2160
         List            =   "Frm_Guardians.frx":1DF7
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtTitle 
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
         Left            =   2400
         MaxLength       =   5
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtOtherNames 
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
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtSurname 
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
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtBirthDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   29
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   58720259
         CurrentDate     =   31480
      End
      Begin MSComCtl2.DTPicker dtEntryDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   58720259
         CurrentDate     =   31480
      End
      Begin Stadmis.ShapeButton ShpBttnClearPhoto 
         Height          =   375
         Left            =   120
         TabIndex        =   13
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
         Picture         =   "Frm_Guardians.frx":1E09
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
      Begin Stadmis.ShapeButton ShpBttnAttachPhoto 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
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
         Picture         =   "Frm_Guardians.frx":21A3
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
      Begin VB.Label lblRegistrationDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Date:"
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
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblOccupation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
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
         TabIndex        =   30
         Tag             =   "Y"
         Top             =   3840
         Width           =   870
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
         Left            =   3960
         TabIndex        =   23
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
         Left            =   4320
         TabIndex        =   18
         Top             =   2040
         Width           =   750
      End
      Begin VB.Label lblMaritalStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marital Status:"
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
         Left            =   3120
         TabIndex        =   16
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label lblBirthDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
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
         Left            =   3960
         TabIndex        =   28
         Top             =   3240
         Width           =   780
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
         Left            =   120
         TabIndex        =   26
         Tag             =   "Y"
         Top             =   3240
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
         Left            =   120
         TabIndex        =   20
         Tag             =   "Y"
         Top             =   2640
         Width           =   1125
      End
      Begin VB.Label lblNationalID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "National ID:"
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
         Width           =   855
      End
      Begin VB.Label lblGender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
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
         TabIndex        =   14
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
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
         TabIndex        =   4
         Top             =   840
         Width           =   360
      End
      Begin VB.Label lblOtherNames 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Names:"
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
         TabIndex        =   10
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lblSurname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   840
         Width           =   690
      End
   End
   Begin VB.Image ImgHeader 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Guardians.frx":253D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
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
      TabIndex        =   42
      Tag             =   "Y"
      Top             =   5160
      Width           =   1395
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_Guardians.frx":2D33
      Stretch         =   -1  'True
      Tag             =   "WY"
      Top             =   4920
      Width           =   6255
   End
   Begin VB.Menu MnuNew 
      Caption         =   "&New"
   End
   Begin VB.Menu MnuSave 
      Caption         =   "&Save"
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu MnuDelete 
      Caption         =   "Delete"
   End
   Begin VB.Menu MnuSearch 
      Caption         =   "Sea&rch"
   End
End
Attribute VB_Name = "Frm_Guardians"
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

Public IsNewRecord As Boolean
Public FrmDefinitions$, SetPrivileges$

Private myTable$, myTablePryKey$
Private myRecIndex&, TxtKeyBack&, iSetNo&
Private myTableFixedFldName(&H5) As String
Private myRecDisplayON, IsLoading As Boolean

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
    
    txtNationalID.Text = VBA.vbNullString
    txtSurname.Tag = VBA.vbNullString
    txtSurname.Text = VBA.vbNullString
    txtOtherNames.Text = VBA.vbNullString
    txtPhoneNo.Text = VBA.IIf(myRecDisplayON, VBA.vbNullString, "254")
    txtPostalAddress.Text = "P.O Box "
    txtEmailAddress.Text = VBA.vbNullString
    ShpBttnBriefNotes.TagExtra = VBA.vbNullString
    dtBirthDate.Value = Null
    chkDeceased.Value = vbUnchecked
    chkDiscontinued.Value = vbUnchecked
    myRecDisplayON = myRecDisplayState
    
    'Clear Picture
    ImgVirtualPhoto.Picture = Nothing
    ImgVirtualPhoto.ToolTipText = VBA.vbNullString
    ImgDBPhoto.Picture = Nothing
    ImgDBPhoto.ToolTipText = VBA.vbNullString
    Erase sAdditionalPhoto(&H0).vDataBytes
    
Exit_ClearEntries:
    
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
    
    Fra_Guardian.Enabled = Not State
    txtPhoneNo.Locked = State
    txtPostalAddress.Locked = State
    txtEmailAddress.Locked = State
    
    MnuSave.Visible = Not State
    MnuEdit.Visible = State: MnuDelete.Visible = State
    
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
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Qry_Guardians] WHERE [Guardian ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'Assign the Record's primary key value
            If Not VBA.IsNull(![Guardian ID]) Then txtSurname.Tag = ![Guardian ID]
            dtEntryDate.Value = ![Entry Date]
            If Not VBA.IsNull(![National ID]) Then txtNationalID.Text = ![National ID]
            If Not VBA.IsNull(![Title]) Then txtTitle.Text = ![Title]
            If Not VBA.IsNull(![Surname]) Then txtSurname.Text = ![Surname]
            If Not VBA.IsNull(![Other Names]) Then txtOtherNames.Text = ![Other Names]
            If Not VBA.IsNull(![Gender]) Then cboGender.Text = ![Gender]
            If Not VBA.IsNull(![Marital Status]) Then cboMaritalStatus.Text = ![Marital Status]
            If Not VBA.IsNull(![Phone No]) Then txtPhoneNo.Text = ![Phone No]
            If Not VBA.IsNull(![Postal Address]) Then txtPostalAddress.Text = ![Postal Address] Else txtPostalAddress.Text = VBA.vbNullString
            If Not VBA.IsNull(![Location]) Then txtLocation.Text = ![Location]
            If Not VBA.IsNull(![E-mail Address]) Then txtEmailAddress.Text = ![E-mail Address]
            If Not VBA.IsNull(![Birth Date]) Then dtBirthDate.Value = ![Birth Date] Else dtBirthDate.Value = Null
            If Not VBA.IsNull(![Occupation]) Then txtOccupation.Text = ![Occupation]
            chkDeceased.Value = VBA.IIf(![Deceased], vbChecked, vbUnchecked)
            If Not VBA.IsNull(![Brief Notes]) Then ShpBttnBriefNotes.TagExtra = ![Brief Notes]
            chkDiscontinued.Value = VBA.IIf(![Discontinued], vbChecked, vbUnchecked)
            
            'If the Record contains the Guardian's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display Guardian's Photo
                sAdditionalPhoto(&H0).vDataBytes = ![Photo]
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = vRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = txtSurname.Text & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
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

Private Sub chkDeceased_Click()
On Local Error GoTo Handle_chkDeceased_Click_Error
    
    If myRecDisplayON Then Exit Sub
    
    If Not ValidUserAccess(Me, iSetNo, &H6) Then myRecDisplayON = True: chkDeceased.Value = vbUnchecked: myRecDisplayON = False: Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If chkDeceased.Value = vbChecked Then If vMsgBox("Ticking this option will denote that the displayed Guardian is no longer alive. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then myRecDisplayON = True: chkDeceased.Value = vbUnchecked: myRecDisplayON = False:
    
Exit_chkDeceased_Click:
    
    myRecDisplayON = False 'Denote that Record display process is complete
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_chkDeceased_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Discontinuing Record - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_chkDeceased_Click
    
End Sub

Private Sub chkDiscontinued_Click()
On Local Error GoTo Handle_chkDiscontinued_Click_Error
    
    If myRecDisplayON Then Exit Sub
    
    If Not ValidUserAccess(Me, iSetNo, &H4) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If chkDiscontinued.Value = vbChecked Then If vMsgBox("Ticking this option will disable the Record and will not be available in other Modules. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then chkDiscontinued.Value = vbUnchecked
    
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

Private Sub CmdMoveRec_Click(Index As Integer)
On Local Error GoTo Handle_CmdMoveRec_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Navigate through Records in the specified table
    myRecIndex = NavigateToRec(Me, "SELECT * FROM [Qry_Guardians] ORDER BY [Registered Name] ASC", "Guardian ID", Index, myRecIndex)
    
Exit_CmdMoveRec_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_CmdMoveRec_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Navigation Error - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_CmdMoveRec_Click
    
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
    cboMaritalStatus.ListIndex = &H0
    dtEntryDate.Value = VBA.Date
    cboGender.ListIndex = &H0: iSetNo = &HA
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to clear all entries in Input Boxes
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records in the specified database table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'Get the total number of Records already saved
        lblRecords.Tag = .RecordCount
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        'If a Record is to be displayed then...
        If VBA.LenB(VBA.Trim$(vEditRecordID)) <> &H0 Then
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            'Retrieve the Record with the specified ID
            .Filter = "[Guardian ID] = " & vArrayList(&H0)
            
            'If the record Exists then Call Procedure in this Form to display it
            If Not (.BOF And .EOF) Then DisplayRecord VBA.CLng(![Guardian ID])
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            vEditRecordID = VBA.vbNullString  'Initialize variable
            
            If Not (.BOF And .EOF) Then If UBound(vArrayList) = &H1 Then Call MnuEdit_Click 'Call click event of Edit Menu
            If Not (.BOF And .EOF) Then If UBound(vArrayList) = &H2 Then Call MnuDelete_Click 'Call click event of Delete Menu
            
            'Reinitialize the elements of the fixed-size array and release dynamic-array storage space.
            Erase vArrayList
            
        Else 'If a new Record is to be entered then...
            
            Call MnuNew_Click 'Call click event of New Menu
            
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
    
    myTableFixedFldName(&H0) = "Guardian": myTableFixedFldName(&H1) = "Guardians"
    myTableFixedFldName(&H2) = "Guardian": myTableFixedFldName(&H3) = "Guardians"
    myTableFixedFldName(&H4) = "Tbl_Guardians": myTableFixedFldName(&H5) = "Tbl_ChurchMembers:Church Members"
    
    myTable = "Tbl_" & VBA.Replace(myTableFixedFldName(&H1), " ", VBA.vbNullString)
    
    Me.Caption = App.Title & " : " & myTableFixedFldName(&H3)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub Fra_Photo_Click()
    Call ShpBttnAttachPhoto_Click
End Sub

Private Sub ImgDBPhoto_Click()
    Call ShpBttnAttachPhoto_Click
End Sub

Private Sub MnuDelete_Click()
On Local Error GoTo Handle_MnuDelete_Click_Error
    
    If Not ValidUserAccess(Me, iSetNo, &H3) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the User has not selected any existing Record then...
    If VBA.LenB(VBA.Trim$(txtSurname.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed"
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    If vMsgBox("Are you sure you want to DELETE the displayed " & myTableFixedFldName(&H2) & "'s Record?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    Dim iDependency As Boolean
    
    'Check if there are Records in other tables depending on the displayed Record
    iDependency = CheckForRecordDependants(Me, "[Tbl_ChurchMembers] WHERE [Guardian ID] = " & txtSurname.Tag)
    
    'If the Records exist then...
    If iDependency Then
        
        'Warn User
        vMsgBox "The displayed Guardian has other Records {Church Member Records} depending on it. Delete operation aborted", vbExclamation, App.Title & " : Operation Aborted"
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM " & myTable & " WHERE [Guardian ID] = " & txtSurname.Tag, vAdoCNN, adOpenKeyset, adLockPessimistic
    
    'Denote that the database table has been altered
    vDatabaseAltered = True
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    'Get the total number of Records already saved
    lblRecords.Tag = VBA.Val(lblRecords.Tag) - &H1
    lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "The displayed " & myTableFixedFldName(&H2) & " Record has successfully been deleted.", vbInformation, App.Title & " : Delete"
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
Exit_MnuDelete_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_MnuDelete_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Deleting Record - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuDelete_Click
    
End Sub

Private Sub MnuEdit_Click()
    
    If Not ValidUserAccess(Me, iSetNo, &H2) Then Exit Sub
    
    'If the User has not selected any existing Record then...
    If VBA.LenB(VBA.Trim$(txtSurname.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed"
        If Fra_Guardian.Enabled Then txtSurname.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Call LockEntries(False) 'Call Procedure in this Form to UnLock Input Controls
    
    IsNewRecord = False 'Denote that the displayed Record exists in the database
    
    If Fra_Guardian.Enabled Then txtNationalID.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuNew_Click()
    
    If Not ValidUserAccess(Me, iSetNo, &H1) Then Exit Sub
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    Call LockEntries(False)  'Call Procedure in this Form to UnLock Input Controls
    
    IsNewRecord = True 'Denote that the displayed Record does not exist in the database
    
    txtNationalID.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuSave_Click()
On Local Error GoTo Handle_MnuSave_Click_Error
    
    'If the User has not entered the Guardian's Name then...
    If VBA.LenB(VBA.Trim$(txtSurname.Text)) = &H0 And VBA.LenB(VBA.Trim$(txtOtherNames.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the Guardian's Name", vbExclamation, App.Title & " : Name not entered"
        txtSurname.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the Guardian's Name then...
    If Not VBA.IsNumeric(VBA.Replace(txtPhoneNo.Text, VBA.vbCrLf, VBA.vbNullString)) And txtPhoneNo.Text <> VBA.vbNullString Then
        
        'Warn User
        vMsgBox "Please enter numeric Phone Numbers", vbExclamation, App.Title & " : Name not entered"
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
    
    txtTitle.Text = VBA.Trim$(VBA.Replace(CapAllWords(VBA.Trim$(VBA.Replace(txtTitle.Text, " ", VBA.vbNullString))), ".", VBA.vbNullString)) & "."
    txtTitle.Text = VBA.IIf(VBA.LenB(txtTitle.Text) = &H1, VBA.vbNullString, txtTitle.Text)
    txtNationalID.Text = VBA.Format$(VBA.Trim$(txtNationalID.Text), "00000000")
    
    'If the User has entered the surname then...
    If VBA.LenB(VBA.Trim$(txtSurname.Text)) <> &H0 Then
        
        'Ensure only one name is entered for the surname
        vArrayList = VBA.Split(VBA.Trim$(VBA.Replace(txtSurname.Text, "  ", " ")))
        txtSurname.Text = CapAllWords(vArrayList(&H0))
        
    End If 'Close respective IF..THEN block statement
    
    myRecDisplayON = True
    
    txtOtherNames.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtOtherNames.Text, "  ", " ")))
    txtPostalAddress.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtPostalAddress.Text, "  ", " ")))
    txtPhoneNo.Text = VBA.Replace(txtPhoneNo.Text, " ", VBA.vbNullString)
    txtLocation.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtLocation.Text, "  ", " ")))
    txtEmailAddress.Text = VBA.LCase$(VBA.Trim$(VBA.Replace(txtEmailAddress.Text, " ", VBA.vbNullString)))
    txtOccupation.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtOccupation.Text, "  ", " ")))
    
    myRecDisplayON = False
    
    '---------------------------------------------------------------------------------------------------
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If the Guardian's National ID has been entered then...
        If VBA.LenB(VBA.Trim$(txtNationalID.Text)) <> &H0 Then
            
            'Check if the entered National ID already exists in the database
            .Filter = "[Guardian ID] <> " & VBA.Val(txtSurname.Tag) & " AND [National ID] = '" & txtNationalID.Text & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                vBuffer(&H0) = VBA.vbNullString
                
                If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
                If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A Guardian {" & vBuffer(&H0) & "} with the entered National ID {" & txtNationalID.Text & "} had already been saved. Please enter a different National ID.", vbExclamation, App.Title & " : Duplicate National ID Entry"
                
                txtNationalID.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the Guardian's Full Name has been entered then...
        If VBA.LenB(VBA.Trim$(txtSurname.Text)) <> &H0 And VBA.LenB(VBA.Trim$(txtOtherNames.Text)) <> &H0 Then
            
            'Check if the entered Name already exists in the database
            .Filter = "[Guardian ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Surname] = '" & VBA.Replace(txtSurname.Text, "'", "''") & "' AND [Other Names] = '" & VBA.Replace(txtOtherNames.Text, "'", "''") & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                vBuffer(&H0) = VBA.vbNullString
                
                If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
                If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User and get feedback. If the User decides to abort saving then...
                If vMsgBox("A Guardian {" & vBuffer(&H0) & "} with the entered Name had already been saved. Proceed with saving?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Duplicate Name Entry") = vbNo Then
                    
                    txtSurname.SetFocus 'Move the focus to the specified control
                    GoTo Exit_MnuSave_Click 'Quit this Procedure
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the Guardian's Phone No has been entered then...
        If VBA.LenB(VBA.Trim$(txtPhoneNo.Text)) <> &H0 And VBA.Trim$(txtPhoneNo.Text) <> "254" Then
            
            vArrayList = VBA.Split(txtPhoneNo.Text, VBA.vbCrLf)
            
            'For each entered Phone No...
            For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
                
                'Check if the entered Phone No already exists in the database
                .Filter = "[Guardian ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Phone No] LIKE '%" & VBA.Trim$(vArrayList(vIndex(&H0))) & "%'"
                
                'If the National ID already exists then...
                If Not (.BOF And .EOF) Then
                    
                    vBuffer(&H0) = VBA.vbNullString
                    
                    If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
                    If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
                    
                    Dim Ans%
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Warn User and get feedback. If the User decides to abort saving then...
                    Ans = vMsgBox("A Guardian {" & vBuffer(&H0) & "} with the entered Phone No {" & VBA.Trim$(vArrayList(vIndex(&H0))) & "} had already been saved. Proceed with saving? {No to display the Record}", vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title & " : Duplicate Phone No Entry")
                    
                    If Ans = vbCancel Then
                        
                        txtPhoneNo.SetFocus 'Move the focus to the specified control
                        GoTo Exit_MnuSave_Click 'Quit this Procedure
                        
                    ElseIf Ans = vbNo Then
                        
                        Call DisplayRecord(![Guardian ID])  'Display the Record
                        txtPhoneNo.SetFocus 'Move the focus to the specified control
                        GoTo Exit_MnuSave_Click 'Quit this Procedure
                        
                    Else
                        'Do nothing
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
            Next vIndex(&H0) 'Move to the next entered Phone No
            
        End If 'Close respective IF..THEN block statement
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[Guardian ID] = " & txtSurname.Tag: .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        ![Entry Date] = VBA.FormatDateTime(dtEntryDate.Value, vbShortDate)
        ![National ID] = txtNationalID.Text
        ![Title] = txtTitle.Text
        ![Surname] = txtSurname.Text
        ![Other Names] = txtOtherNames.Text
        ![Gender] = cboGender.Text
        ![Marital Status] = cboMaritalStatus.Text
        ![Phone No] = VBA.IIf(VBA.Trim(txtPhoneNo.Text) = "254", VBA.vbNullString, txtPhoneNo.Text)
        ![Postal Address] = VBA.IIf(VBA.Trim(txtPostalAddress.Text) = "P.O Box", VBA.vbNullString, txtPostalAddress.Text)
        ![Location] = txtLocation.Text
        ![E-mail Address] = txtEmailAddress.Text
        If VBA.IsNull(dtBirthDate.Value) Then ![Birth Date] = Null Else ![Birth Date] = VBA.FormatDateTime(dtBirthDate.Value, vbShortDate)
        
        ![Photo] = Null
        
        'If the Guardian has a Photo then Assign the Photo to its field
        If ImgDBPhoto.Picture <> &H0 Then .Fields("Photo").AppendChunk sAdditionalPhoto(&H0).vDataBytes
        
        ![Occupation] = txtOccupation.Text
        ![Deceased] = chkDeceased.Value
        
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![Discontinued] = chkDiscontinued.Value
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        txtSurname.Tag = ![Guardian ID]
        
        .Close 'Close the opened object and any dependent objects
        
        'Denote that the database table has been altered
        vDatabaseAltered = True
        
        'Get the total number of Records already saved
        lblRecords.Tag = VBA.Val(lblRecords.Tag) + VBA.IIf(IsNewRecord, &H1, &H0)
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
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
    
    'If Err is 'Invalid pointer error' while saving Photo then...
    If Err.Number = -2147287031# Then
        
        'Clear Picture
        ImgVirtualPhoto.Picture = Nothing
        ImgVirtualPhoto.ToolTipText = VBA.vbNullString
        ImgDBPhoto.Picture = Nothing
        ImgDBPhoto.ToolTipText = VBA.vbNullString
        Erase sAdditionalPhoto(&H0).vDataBytes
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Warn User
        vMsgBox "Unable to Attach the image to the Record. {" & Err.Description & "}", vbExclamation, App.Title & " : Error" & Err.Number, Me
        
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

Private Sub MnuSearch_Click()
On Local Error GoTo Handle_MnuSearch_Click_Error
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Guardian ID], [Guardian Name], [Gender], [National ID], [Occupation], [Phone No] FROM [Qry_Guardians] ORDER BY [Registered Name] ASC", txtSurname.Tag, , "1", , , myTableFixedFldName(&H3), , "[Tbl_Guardians]; WHERE [Guardian ID] = $", , 9300)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(VBA.Trim$(vBuffer(&H0))) = &H0 Then GoTo Exit_MnuSearch_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Call procedure to display the selected Record
    Call DisplayRecord(VBA.CLng(vArrayList(&H0)))
    
Exit_MnuSearch_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_MnuSearch_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Searching - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuSearch_Click
    
End Sub

Private Sub ShpBttnAttachPhoto_Click()
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, Not (Fra_Guardian.Enabled), "0")
End Sub

Private Sub ShpBttnBriefNotes_Click()
    'Call Function in Mdl_Stadmis to display Notes Input Form
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_Guardian.Enabled, myTableFixedFldName(&H2))
End Sub

Private Sub ShpBttnClearPhoto_Click()
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, Not (Fra_Guardian.Enabled), "1")
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii)))
    If KeyAscii = 32 Then KeyAscii = Empty
End Sub

Private Sub txtLocation_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtLocation.Text = VBA.vbNullString Or chkAutoComplete(&H3).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtLocation, "[Tbl_Guardians]", "Location", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
    If txtLocation.Text <> txtLocation.SelText And Not txtLocation.Locked Then myRecDisplayON = True: txtLocation.Text = VBA.Replace(txtLocation.Text, txtLocation.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtNationalID_KeyPress(KeyAscii As Integer)
    'Discard non-numeric entries
    KeyAscii = VBA.IIf((((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126))) And KeyAscii <> VBA.Asc("."), KeyAscii = Empty, KeyAscii)
End Sub

Private Sub txtOccupation_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtOccupation.Text = VBA.vbNullString Or chkAutoComplete(&H4).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtOccupation, "[Tbl_Guardians]", "Occupation", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtOccupation_Validate(Cancel As Boolean)
    If txtOccupation.Text <> txtOccupation.SelText And Not txtOccupation.Locked Then myRecDisplayON = True: txtOccupation.Text = VBA.Replace(txtOccupation.Text, txtOccupation.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtPhoneNo_GotFocus()
    txtPhoneNo.SelStart = VBA.LenB(txtPhoneNo.Text)
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
    'Discard non-numeric entries
    KeyAscii = VBA.IIf((((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126))) And KeyAscii <> vbKeyReturn, KeyAscii = Empty, KeyAscii)
End Sub

Private Sub txtPostalAddress_GotFocus()
    txtPostalAddress.SelStart = VBA.LenB(txtPostalAddress.Text)
End Sub

Private Sub txtPostalAddress_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtPostalAddress.Text = VBA.vbNullString Or chkAutoComplete(&H2).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtPostalAddress, "[Tbl_Guardians]", "Postal Address", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtPostalAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtPostalAddress_Validate(Cancel As Boolean)
    If txtPostalAddress.Text <> txtPostalAddress.SelText And Not txtPostalAddress.Locked Then myRecDisplayON = True: txtPostalAddress.Text = VBA.Replace(txtPostalAddress.Text, txtPostalAddress.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtSurname_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtSurname.Text = VBA.vbNullString Or chkAutoComplete(&H1).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtSurname, "[Tbl_Guardians]", "Surname", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtSurname_GotFocus()
    'Highlight contents
    txtSurname.SelStart = &H0: txtSurname.SelLength = VBA.LenB(txtSurname.Text)
End Sub

Private Sub txtSurname_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtSurname_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then KeyAscii = Empty
    If KeyAscii = vbKeyReturn Then txtOtherNames.SetFocus: KeyAscii = Empty
End Sub

Private Sub txtTitle_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtTitle.Text = VBA.vbNullString Or chkAutoComplete(&H0).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtTitle, "[Tbl_Guardians]", "Title", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtTitle_GotFocus()
    'Highlight contents
    txtTitle.SelStart = &H0: txtTitle.SelLength = VBA.LenB(txtTitle.Text)
End Sub

Private Sub txtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtTitle_Validate(Cancel As Boolean)
    'If txtTitle.Text <> txtTitle.SelText And Not txtTitle.Locked Then If TxtKeyBack <> vbKeyTab Then myRecDisplayON = True: txtTitle.Text = VBA.Replace(txtTitle.Text, txtTitle.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

