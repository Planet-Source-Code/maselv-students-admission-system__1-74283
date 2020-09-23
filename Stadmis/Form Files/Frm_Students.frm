VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60D73138-4A06-4DA5-838D-E17FF732D00B}#1.0#0"; "prjXTab.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Students 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Students"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6495
   Icon            =   "Frm_Students.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   1320
      TabIndex        =   68
      Tag             =   "AutoSizer:X"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Shape shpFraPhotoBorder 
         BorderColor     =   &H009FB9C8&
         BorderWidth     =   3
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image ImgDBPhoto 
         Height          =   1215
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
   End
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   2640
      TabIndex        =   31
      Top             =   5820
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
      Picture         =   "Frm_Students.frx":0ECA
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
      Index           =   0
      Left            =   4920
      Picture         =   "Frm_Students.frx":1264
      Style           =   1  'Graphical
      TabIndex        =   55
      Tag             =   "XY"
      ToolTipText     =   "Move First"
      Top             =   5820
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
      Left            =   5280
      Picture         =   "Frm_Students.frx":15A6
      Style           =   1  'Graphical
      TabIndex        =   56
      Tag             =   "XY"
      ToolTipText     =   "Move Previous"
      Top             =   5820
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
      Left            =   5640
      Picture         =   "Frm_Students.frx":18E8
      Style           =   1  'Graphical
      TabIndex        =   57
      Tag             =   "XY"
      ToolTipText     =   "Move Next"
      Top             =   5820
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
      Index           =   3
      Left            =   6000
      Picture         =   "Frm_Students.frx":1C2A
      Style           =   1  'Graphical
      TabIndex        =   58
      Tag             =   "XY"
      ToolTipText     =   "Move Last"
      Top             =   5820
      Width           =   375
   End
   Begin prjXTab.XTab XTab 
      Height          =   5055
      Left            =   120
      TabIndex        =   61
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8916
      TabCaption(0)   =   "Personal Details"
      TabContCtrlCnt(0)=   6
      Tab(0)ContCtrlCap(1)=   "Fra_Photo0"
      Tab(0)ContCtrlCap(2)=   "txtPhoneNo"
      Tab(0)ContCtrlCap(3)=   "txtEmailAddress"
      Tab(0)ContCtrlCap(4)=   "txtPostalAddress"
      Tab(0)ContCtrlCap(5)=   "txtOccupation"
      Tab(0)ContCtrlCap(6)=   "Fra_Student0"
      TabCaption(1)   =   "School Details"
      TabContCtrlCnt(1)=   7
      Tab(1)ContCtrlCap(1)=   "cmdAddDormitory"
      Tab(1)ContCtrlCap(2)=   "cmdAddRelationship"
      Tab(1)ContCtrlCap(3)=   "txtGuardianName"
      Tab(1)ContCtrlCap(4)=   "Lv"
      Tab(1)ContCtrlCap(5)=   "CmdAddGuardian"
      Tab(1)ContCtrlCap(6)=   "Fra_Student1"
      Tab(1)ContCtrlCap(7)=   "Fra_Student2"
      TabCaption(2)   =   "Medical Info"
      TabContCtrlCnt(2)=   3
      Tab(2)ContCtrlCap(1)=   "cmdAddMedicalDoctor"
      Tab(2)ContCtrlCap(2)=   "txtMedicalDoctor"
      Tab(2)ContCtrlCap(3)=   "Fra_Student3"
      TabTheme        =   1
      ActiveTabBackStartColor=   16777215
      ActiveTabBackEndColor=   16777215
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      ActiveTabForeColor=   16711680
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   16777215
      DisabledTabForeColor=   10526880
      Begin VB.CommandButton cmdAddDormitory 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   -71760
         Picture         =   "Frm_Students.frx":1F6C
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Add Dormitory"
         Top             =   1440
         Width           =   345
      End
      Begin VB.CommandButton cmdAddMedicalDoctor 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   -74760
         Picture         =   "Frm_Students.frx":22F6
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Add Parent/Guardian"
         Top             =   840
         Width           =   345
      End
      Begin VB.TextBox txtMedicalDoctor 
         BackColor       =   &H8000000F&
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
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   840
         Width           =   4695
      End
      Begin VB.Frame Fra_Student 
         BackColor       =   &H00FBFDFB&
         Height          =   4575
         Index           =   3
         Left            =   -74880
         TabIndex        =   70
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txtDisabilities 
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
            Height          =   1005
            Left            =   3120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Tag             =   "WY"
            Top             =   2400
            Width           =   2775
         End
         Begin VB.TextBox txtMedicalHistory 
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
            Height          =   3285
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   79
            Tag             =   "WY"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtAllergies 
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
            Height          =   1005
            Left            =   3120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Tag             =   "WY"
            Top             =   1080
            Width           =   2775
         End
         Begin VB.CommandButton cmdClearMedicalDoctor 
            Height          =   315
            Left            =   480
            Picture         =   "Frm_Students.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Clear selected Parent/Guardian"
            Top             =   480
            Width           =   345
         End
         Begin VB.CommandButton cmdSelectMedicalDoctor 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   840
            Picture         =   "Frm_Students.frx":2A0A
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Select a Parent/Guardian"
            Top             =   480
            Width           =   345
         End
         Begin VB.Label lblDisabilities 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disabilities:"
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
            TabIndex        =   80
            Top             =   2160
            Width           =   795
         End
         Begin VB.Label lblMedicalHistory 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Medical History:"
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
            TabIndex        =   78
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label lblAllergy 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allergies:"
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
            TabIndex        =   76
            Top             =   840
            Width           =   660
         End
         Begin VB.Label lblMedicalDoctor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Physician:"
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
            TabIndex        =   75
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdAddRelationship 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   -73680
         Picture         =   "Frm_Students.frx":2F94
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Add Relationship"
         Top             =   2520
         Width           =   345
      End
      Begin VB.TextBox txtGuardianName 
         BackColor       =   &H8000000F&
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
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3495
      End
      Begin MSComctlLib.ListView Lv 
         Height          =   1920
         Left            =   -74760
         TabIndex        =   54
         Tag             =   "AutoSizer:WH"
         Top             =   2880
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   3387
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         Icons           =   "ImgLst"
         SmallIcons      =   "ImgLst"
         ForeColor       =   128
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Frame Fra_Photo 
         BackColor       =   &H00FBFDFB&
         Height          =   1935
         Index           =   0
         Left            =   240
         TabIndex        =   64
         Tag             =   "XY"
         Top             =   480
         Width           =   1935
         Begin VB.Image ImgVirtualPhoto 
            Height          =   255
            Left            =   240
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Image ImgDBPhoto 
            Height          =   1575
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Tag             =   "HW"
            Top             =   240
            Width           =   1695
         End
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
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2640
         Width           =   1695
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
         TabIndex        =   23
         Tag             =   "WY"
         Top             =   3840
         Width           =   3735
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
         TabIndex        =   19
         Tag             =   "WY"
         Top             =   3240
         Width           =   3495
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
         TabIndex        =   28
         Tag             =   "WY"
         Top             =   4440
         Width           =   3135
      End
      Begin VB.CommandButton CmdAddGuardian 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   -74760
         Picture         =   "Frm_Students.frx":331E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Add Parent/Guardian"
         Top             =   2160
         Width           =   345
      End
      Begin VB.Frame Fra_Student 
         BackColor       =   &H00FBFDFB&
         Height          =   4575
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Tag             =   "HW"
         Top             =   360
         Width           =   6015
         Begin VB.CheckBox chkDiscontinued 
            BackColor       =   &H00FBFDFB&
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
            TabIndex        =   30
            Tag             =   "Y"
            Top             =   4080
            Width           =   1215
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
            Left            =   2400
            MaxLength       =   15
            TabIndex        =   2
            Top             =   480
            Width           =   1695
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
            TabIndex        =   4
            Top             =   1080
            Width           =   3735
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
            TabIndex        =   7
            Top             =   1680
            Width           =   495
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
            ItemData        =   "Frm_Students.frx":36A8
            Left            =   3000
            List            =   "Frm_Students.frx":36B2
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1680
            Width           =   855
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
            Left            =   3960
            MaxLength       =   8
            TabIndex        =   21
            Top             =   2880
            Width           =   1935
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
            ItemData        =   "Frm_Students.frx":36C4
            Left            =   3960
            List            =   "Frm_Students.frx":36CE
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   3480
            Width           =   1935
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
            TabIndex        =   16
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CheckBox chkDeceased 
            BackColor       =   &H00FBFDFB&
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
            TabIndex        =   29
            Tag             =   "Y"
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CheckBox chkAutoComplete 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   1680
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkAutoComplete 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   1
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   480
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkAutoComplete 
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   2880
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkAutoComplete 
            Caption         =   "Check1"
            Height          =   255
            Index           =   3
            Left            =   3960
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   2280
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkAutoComplete 
            Caption         =   "Check1"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   4080
            Value           =   1  'Checked
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtBirthDate 
            Height          =   285
            Left            =   3960
            TabIndex        =   11
            Top             =   1680
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   16580611
            CurrentDate     =   31480
         End
         Begin Stadmis.ShapeButton ShpBttnClearPhoto 
            Height          =   375
            Left            =   120
            TabIndex        =   60
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
            Picture         =   "Frm_Students.frx":36E3
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
            TabIndex        =   59
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
            Picture         =   "Frm_Students.frx":3A7D
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
            Left            =   2160
            TabIndex        =   0
            Top             =   240
            Width           =   690
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
            TabIndex        =   3
            Top             =   840
            Width           =   1005
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
            TabIndex        =   5
            Top             =   1440
            Width           =   360
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
            Left            =   3000
            TabIndex        =   8
            Top             =   1440
            Width           =   585
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
            Left            =   3960
            TabIndex        =   20
            Top             =   2640
            Width           =   855
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
            TabIndex        =   17
            Tag             =   "Y"
            Top             =   2640
            Width           =   1125
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
            TabIndex        =   22
            Tag             =   "Y"
            Top             =   3240
            Width           =   1110
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
            TabIndex        =   10
            Top             =   1440
            Width           =   780
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
            Left            =   3960
            TabIndex        =   24
            Top             =   3240
            Width           =   1050
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
            Left            =   2160
            TabIndex        =   12
            Top             =   2040
            Width           =   750
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
            TabIndex        =   14
            Top             =   2040
            Width           =   660
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
            TabIndex        =   26
            Tag             =   "Y"
            Top             =   3840
            Width           =   870
         End
      End
      Begin VB.Frame Fra_Student 
         BackColor       =   &H00FBFDFB&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1455
         Index           =   1
         Left            =   -74880
         TabIndex        =   63
         Top             =   360
         Width           =   6015
         Begin VB.TextBox txtDormitoryName 
            BackColor       =   &H8000000F&
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
            Left            =   4200
            TabIndex        =   82
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdSelectDormitory 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   3840
            Picture         =   "Frm_Students.frx":3E17
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Select a Dormitory"
            Top             =   1080
            Width           =   345
         End
         Begin VB.CommandButton cmdClearDormitory 
            Height          =   315
            Left            =   3480
            Picture         =   "Frm_Students.frx":43A1
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Clear selected Dormitory"
            Top             =   1080
            Width           =   345
         End
         Begin VB.CommandButton cmdGenerateAdmNo 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   2640
            Picture         =   "Frm_Students.frx":472B
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Select a Class"
            Top             =   480
            Width           =   345
         End
         Begin VB.CommandButton cmdClearClass 
            Height          =   315
            Left            =   120
            Picture         =   "Frm_Students.frx":4CB5
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Clear selected Class"
            Top             =   1080
            Width           =   345
         End
         Begin VB.CommandButton cmdSelectClass 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   480
            Picture         =   "Frm_Students.frx":503F
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Select a Class"
            Top             =   1080
            Width           =   345
         End
         Begin VB.TextBox txtClassName 
            BackColor       =   &H8000000F&
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
            Left            =   840
            TabIndex        =   69
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtAdmNo 
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
            MaxLength       =   15
            TabIndex        =   36
            Top             =   480
            Width           =   735
         End
         Begin MSComCtl2.DTPicker dtAdmissionDate 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   16580611
            CurrentDate     =   31480
         End
         Begin VB.Label lblDormitory 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Dormitory Assigned:"
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
            TabIndex        =   40
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblClassName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Class Admitted:"
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
            TabIndex        =   37
            Top             =   840
            Width           =   1125
         End
         Begin VB.Label lblAdmissionDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Date:"
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
            TabIndex        =   32
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblAdmNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adm No:"
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
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Fra_Student 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Parents/Guardians:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3135
         Index           =   2
         Left            =   -74880
         TabIndex        =   44
         Top             =   1800
         Width           =   6015
         Begin Stadmis.ShapeButton shpBttnAddGuardian 
            Height          =   375
            Left            =   5040
            TabIndex        =   53
            Top             =   660
            Width           =   855
            _ExtentX        =   1508
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
            Caption         =   "      Add"
            Picture         =   "Frm_Students.frx":55C9
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
         Begin VB.TextBox txtRelationshipName 
            BackColor       =   &H8000000F&
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   720
            Width           =   2655
         End
         Begin VB.CommandButton cmdClearRelationship 
            Height          =   315
            Left            =   1560
            Picture         =   "Frm_Students.frx":5963
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Clear selected Relationship"
            Top             =   720
            Width           =   345
         End
         Begin VB.CommandButton cmdSelectRelationship 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   1920
            Picture         =   "Frm_Students.frx":5CED
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Select a Relationship"
            Top             =   720
            Width           =   345
         End
         Begin VB.CheckBox chkPaysFees 
            BackColor       =   &H00FBFDFB&
            Caption         =   "Pays Fees"
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
            Left            =   4800
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSelectGuardian 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   840
            Picture         =   "Frm_Students.frx":6277
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Select a Parent/Guardian"
            Top             =   360
            Width           =   345
         End
         Begin VB.CommandButton cmdClearGuardian 
            Height          =   315
            Left            =   480
            Picture         =   "Frm_Students.frx":6801
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Clear selected Parent/Guardian"
            Top             =   360
            Width           =   345
         End
         Begin VB.Label lblRelationshipName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship:"
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
            TabIndex        =   49
            Top             =   780
            Width           =   930
         End
      End
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   -720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Students.frx":6B8B
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   62
      Tag             =   "Y"
      Top             =   5880
      Width           =   1395
   End
   Begin VB.Image ImgHeader 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Students.frx":6F25
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6615
   End
   Begin VB.Image ImgFooter 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Students.frx":771B
      Stretch         =   -1  'True
      Tag             =   "WY"
      Top             =   5520
      Width           =   6615
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
      Caption         =   "Search"
   End
   Begin VB.Menu MnuLvs 
      Caption         =   "LvMenus"
      Visible         =   0   'False
      Begin VB.Menu MnuLv 
         Caption         =   "Remove from List"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Frm_Students"
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

Public iSelColPos&
Public iTargetFields$, iLvwItemDataType$, iPhotoSpecifications$, iSearchIDs$

Private myTable$, myTablePryKey$
Private myRecIndex&, TxtKeyBack&, iSetNo&
Private myTableFixedFldName(&H5) As String
Private myRecDisplayON, IsLoading As Boolean

Private Xaxis&
Private iShiftKey%
Private DragPhoto, iSelectionComplete, PhotoDragged As Boolean

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
    
    cmdClearClass.Tag = VBA.vbNullString
    cmdSelectClass.Tag = VBA.vbNullString
    cmdClearDormitory.Tag = VBA.vbNullString
    cmdSelectDormitory.Tag = VBA.vbNullString
    
    ShpBttnBriefNotes.TagExtra = VBA.vbNullString
    dtBirthDate.Value = Null
    chkDeceased.Value = vbUnchecked
    chkDiscontinued.Value = vbUnchecked
    myRecDisplayON = myRecDisplayState
    
    Lv.ListItems.Clear
    
    'If no headers have been created then...
    If Lv.ColumnHeaders.Count = &H0 Then
        
        Lv.ColumnHeaders.Clear
        Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "    Guardian ID", 550
        Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Guardian Name", 2200
        Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Relationship", 1100
        Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Pays Fees", 1000
        Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "National ID", 1100
        Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Location", 1450
        
    End If
    
    Call cmdClearGuardian_Click
    
    'Clear Picture
    ImgVirtualPhoto.Picture = Nothing
    ImgVirtualPhoto.ToolTipText = VBA.vbNullString
    ImgDBPhoto(&H0).Picture = Nothing
    ImgDBPhoto(&H0).ToolTipText = VBA.vbNullString
    Erase sAdditionalPhoto(&H0).vDataBytes
    
    'Clear Picture
    ImgDBPhoto(&H1).Picture = Nothing
    ImgDBPhoto(&H1).ToolTipText = VBA.vbNullString
    
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
    
    Fra_Student(&H0).Enabled = Not State
    Fra_Student(&H1).Enabled = Fra_Student(&H0).Enabled
    Fra_Student(&H2).Enabled = Fra_Student(&H0).Enabled
    txtPhoneNo.Locked = State
    txtPostalAddress.Locked = State
    txtEmailAddress.Locked = State
    Lv.Checkboxes = Not State
    If Not State Then Fra_Student(&H1).Enabled = ValidUserAccess(Me, iSetNo, &H7)
    
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
        
        .Open "SELECT * FROM [Qry_Students] WHERE [Student ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'Assign the Record's primary key value
            dtAdmissionDate.Value = ![Admission Date]
            If Not VBA.IsNull(![Student ID]) Then txtSurname.Tag = ![Student ID]
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
            
            If Not VBA.IsNull(![Physician ID]) Then txtMedicalDoctor.Tag = ![Physician ID]
            If Not VBA.IsNull(![Physician Name]) Then txtMedicalDoctor.Text = ![Physician Name]
            If Not VBA.IsNull(![Medical History]) Then txtMedicalHistory.Text = ![Medical History]
            If Not VBA.IsNull(![Allergies]) Then txtAllergies.Text = ![Allergies]
            If Not VBA.IsNull(![Disabilities]) Then txtDisabilities.Text = ![Disabilities]
            
            If Not VBA.IsNull(![Occupation]) Then txtOccupation.Text = ![Occupation]
            If Not VBA.IsNull(![Adm No]) Then txtAdmNo.Text = ![Adm No]
            If Not VBA.IsNull(![Brief Notes]) Then ShpBttnBriefNotes.TagExtra = ![Brief Notes]
            chkDeceased.Value = VBA.IIf(![Deceased], vbChecked, vbUnchecked)
            chkDiscontinued.Value = VBA.IIf(![Discontinued], vbChecked, vbUnchecked)
            
            'If the Record contains the Student's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display Student's Photo
                sAdditionalPhoto(&H0).vDataBytes = ![Photo]
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = vRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = txtSurname.Text & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto(&H0), Fra_Photo(&H0))
                
            End If 'Close respective IF..THEN block statement
            
            .Close 'Close the opened object and any dependent objects
            
            .Open "SELECT * FROM [Qry_StudentClasses] WHERE [Student ID] = " & VBA.Val(txtSurname.Tag) & " AND YEAR([Date Assigned]) = " & VBA.Year(dtAdmissionDate.Value), vAdoCNN, adOpenKeyset, adLockReadOnly
            
            'If there are records in the table then...
            If Not (.BOF And .EOF) Then
                
                If Not VBA.IsNull(![Student Class ID]) Then cmdClearClass.Tag = ![Student Class ID]
                If Not VBA.IsNull(![Stream ID]) Then txtClassName.Tag = ![Stream ID]: cmdSelectClass.Tag = ![Stream ID]
                If Not VBA.IsNull(![Class]) Then txtClassName.Text = ![Class]
                
            End If 'Close respective IF..THEN block statement
            
            .Close 'Close the opened object and any dependent objects
            
            .Open "SELECT * FROM [Qry_StudentDormitories] WHERE [Student ID] = " & VBA.Val(txtSurname.Tag) & " AND YEAR([Date Assigned]) = " & VBA.Year(dtAdmissionDate.Value), vAdoCNN, adOpenKeyset, adLockReadOnly
            
            'If there are records in the table then...
            If Not (.BOF And .EOF) Then
                
                If Not VBA.IsNull(![Student Dormitory ID]) Then cmdClearDormitory.Tag = ![Student Dormitory ID]
                If Not VBA.IsNull(![Dormitory ID]) Then txtDormitoryName.Tag = ![Dormitory ID]: cmdSelectDormitory.Tag = ![Dormitory ID]
                If Not VBA.IsNull(![Dormitory Name]) Then txtDormitoryName.Text = ![Dormitory Name]
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
        Call FillListView(Lv, "SELECT [Guardian ID], [Guardian Name], ([Relationship Name]) AS [Relationship], [Pays Fees], [National ID], [Location] FROM [Qry_StudentGuardians] WHERE [Student ID] = " & VBA.Val(txtSurname.Tag) & " ORDER BY [Student Name] ASC", "1", , "Student Guardian ID||Relationship ID", "Pays Fees:=YES", &HC000C0)
        
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
    
    If chkDeceased.Value = vbChecked Then If vMsgBox("Ticking this option will denote that the displayed Student is no longer alive. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then myRecDisplayON = True: chkDeceased.Value = vbUnchecked: myRecDisplayON = False:
    
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

Private Sub cmdAddDormitory_Click()
    
    If VBA.LenB(VBA.Trim$(txtDormitoryName.Tag)) <> &H0 Then vEditRecordID = txtDormitoryName.Tag & "||||"
    FrmDefinitions = "Dormitory:Dormitories:Dormitory:Dormitories:Tbl_Dormitories:Tbl_StudentDormitories:Students"
    Frm_Dormitories.Show vbModal, Me
    If vDatabaseAltered And Fra_Student(&H0).Enabled Then vDatabaseAltered = False: Call cmdClearDormitory_Click
    
End Sub

Private Sub CmdAddGuardian_Click()
    
    If VBA.LenB(VBA.Trim$(txtGuardianName.Tag)) <> &H0 Then vEditRecordID = txtGuardianName.Tag & "||||"
    Frm_Guardians.Show vbModal, Me
    If vDatabaseAltered And Fra_Student(&H0).Enabled Then vDatabaseAltered = False: Call cmdClearGuardian_Click
    
End Sub

Private Sub cmdAddMedicalDoctor_Click()
    
    If VBA.LenB(VBA.Trim$(txtMedicalDoctor.Tag)) <> &H0 Then vEditRecordID = txtMedicalDoctor.Tag & "||||"
    Frm_Guardians.Show vbModal, Me
    If vDatabaseAltered And Fra_Student(&H0).Enabled Then vDatabaseAltered = False: Call cmdClearMedicalDoctor_Click
    
End Sub

Private Sub cmdAddRelationship_Click()
On Local Error GoTo Handle_CmdAddRelationship_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(txtRelationshipName.Tag)) <> &H0 Then vEditRecordID = txtRelationshipName.Tag & "|||"
    Call LoadUnloadedFormViaStringName(Me, Frm_FamilyRelationships.Name, "Relationship:Relationships:Relationship:Relationships:Tbl_Relationships:Tbl_StudentRelatives:Students")
    If vDatabaseAltered And Fra_Student(&H0).Enabled Then vDatabaseAltered = False: Call cmdClearRelationship_Click
    
Exit_CmdAddRelationship_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_CmdAddRelationship_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Loading Relationship Form - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_CmdAddRelationship_Click
    
End Sub

Private Sub cmdClearClass_Click()
    
    'If the User is not allowed to alter the Student's allocated class then quit this procedure
    If cmdSelectClass.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &H9, &H2) Then Exit Sub
    
    txtClassName.Tag = VBA.vbNullString: txtClassName.Text = VBA.vbNullString
    
End Sub

Private Sub cmdClearDormitory_Click()
    
    'If the User is not allowed to alter the Student's allocated Dormitory then quit this procedure
    If cmdSelectDormitory.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &HB, &H2) Then Exit Sub
    
    txtDormitoryName.Tag = VBA.vbNullString: txtDormitoryName.Text = VBA.vbNullString
    
End Sub

Private Sub cmdClearGuardian_Click()
    txtGuardianName.Tag = VBA.vbNullString: txtGuardianName.Text = VBA.vbNullString: shpBttnAddGuardian.Tag = VBA.vbNullString: Call cmdClearRelationship_Click
End Sub

Private Sub cmdClearMedicalDoctor_Click()
    txtMedicalDoctor.Tag = VBA.vbNullString: txtMedicalDoctor.Text = VBA.vbNullString
End Sub

Private Sub cmdClearRelationship_Click()
    txtRelationshipName.Tag = VBA.vbNullString: txtRelationshipName.Text = VBA.vbNullString
End Sub

Private Sub cmdGenerateAdmNo_Click()
On Local Error GoTo Handle_cmdGenerateAdmNo_Click_Error
    
    Dim MousePointerState%
    Dim nRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB(, False) 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set nRs = New ADODB.Recordset
    With nRs 'Execute a series of statements on vRs recordset
        
        vIndex(&H0) = &H0
        
        .Open "SELECT * FROM [Qry_Students] ORDER BY [Adm No] DESC", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then If Not VBA.IsNull(![Adm No]) Then vIndex(&H0) = VBA.Val(![Adm No])
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_cmdGenerateAdmNo_Click:
    
    txtAdmNo.Text = VBA.Format$(VBA.Val(vIndex(&H0) + &H1), "00000")
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdGenerateAdmNo_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Generating Adm No - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_cmdGenerateAdmNo_Click
    
End Sub

Private Sub CmdMoveRec_Click(Index As Integer)
On Local Error GoTo Handle_CmdMoveRec_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Navigate through Records in the specified table
    myRecIndex = NavigateToRec(Me, "SELECT [Student ID] FROM [Qry_Students] ORDER BY [Admission Date] DESC, [Registered Name] ASC", "Student ID", Index, myRecIndex)
    
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

Private Sub cmdSelectClass_Click()
On Local Error GoTo Handle_cmdSelectClass_Click_Error
    
    'If the User is not allowed to alter the Student's allocated class then quit this procedure
    If cmdSelectClass.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &H9, &H2) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Stream ID], ([Class Name] & ' ' & [Stream Name]) AS [Class], [Capacity] FROM [Qry_Streams] WHERE [Discontinued] = FALSE ORDER BY [Class Level] ASC, [Stream Name] ASC", cmdSelectClass.Tag, "1;2;3;4", "1", , , "Classes", , , , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectClass_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtClassName.Tag = vArrayList(&H0) 'Assign Stream ID
    txtClassName.Text = vArrayList(&H1) 'Assign Class Name
    
    cmdSelectDormitory.SetFocus
    
Exit_cmdSelectClass_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectClass_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Class - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectClass_Click
    
End Sub

Private Sub cmdSelectDormitory_Click()
On Local Error GoTo Handle_cmdSelectDormitory_Click_Error
    
    'If the User is not allowed to alter the Student's allocated Dormitory then quit this procedure
    If cmdSelectDormitory.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &HB, &H2) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Dormitory ID], ([Dormitory Name]) AS [Dormitory], [Capacity] FROM [Tbl_Dormitories] WHERE [Discontinued] = FALSE ORDER BY [Dormitory Name] ASC", cmdSelectDormitory.Tag, "1;2;3", "1", , , "Dormitories", , , , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectDormitory_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtDormitoryName.Tag = vArrayList(&H0) 'Assign Dormitory ID
    txtDormitoryName.Text = vArrayList(&H1) 'Assign Dormitory Name
    
    cmdSelectDormitory.SetFocus
    
Exit_cmdSelectDormitory_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectDormitory_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Dormitory - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectDormitory_Click
    
End Sub

Private Sub cmdSelectGuardian_Click()
On Local Error GoTo Handle_cmdSelectGuardian_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Guardian ID], [Guardian Name], [National ID], [Location] FROM [Qry_Guardians] WHERE [Discontinued] = FALSE AND [Deceased] = FALSE ORDER BY [Guardian Name] ASC", cmdSelectGuardian.Tag, "1;2;3;4", "1", , , "Guardians", , "[Qry_Guardians]; WHERE [Guardian ID] = $;1;Guardian Name", , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectGuardian_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtGuardianName.Tag = vArrayList(&H0)  'Assign Guardian ID
    txtGuardianName.Text = vArrayList(&H1) 'Assign Guardian Name
    If UBound(vArrayList) >= &H2 Then txtGuardianName.Text = txtGuardianName.Text & VBA.IIf(VBA.LenB(vArrayList(&H2)) <> &H0, " -> " & vArrayList(&H2), VBA.vbNullString) 'Assign Record Name
    cmdSelectGuardian.Tag = vArrayList(&H2) & "|" & vArrayList(&H3)   'Assign National ID & Location
    
    chkPaysFees.SetFocus
    
Exit_cmdSelectGuardian_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectGuardian_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Guardian - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectGuardian_Click
    
End Sub

Private Sub cmdSelectMedicalDoctor_Click()
On Local Error GoTo Handle_cmdSelectMedicalDoctor_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Guardian ID], ([Guardian Name]) AS [Doctor Name], [National ID], [Location] FROM [Qry_Guardians] WHERE [Discontinued] = FALSE AND [Deceased] = FALSE ORDER BY [Registered Name] ASC", cmdSelectMedicalDoctor.Tag, "1;2;3;4", "1", , , "Guardians", , "[Qry_Guardians]; WHERE [Guardian ID] = $;1;Guardian Name", , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectMedicalDoctor_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtMedicalDoctor.Tag = vArrayList(&H0)  'Assign Guardian ID
    txtMedicalDoctor.Text = vArrayList(&H1) 'Assign Guardian Name
    
    chkPaysFees.SetFocus
    
Exit_cmdSelectMedicalDoctor_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectMedicalDoctor_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Guardian - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectMedicalDoctor_Click
    
End Sub

Private Sub cmdSelectRelationship_Click()
On Local Error GoTo Handle_cmdSelectRelationship_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Relationship ID], [Relationship Name] FROM [Tbl_Relationships] WHERE [Discontinued] = FALSE ORDER BY [Hierarchical Level] ASC", txtRelationshipName.Tag, "1;2", "1", , , "Relationships", , , , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectRelationship_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtRelationshipName.Tag = vArrayList(&H0)  'Assign Relationship ID
    txtRelationshipName.Text = vArrayList(&H1) 'Assign Relationship Name
    
    shpBttnAddGuardian.SetFocus
    
Exit_cmdSelectRelationship_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectRelationship_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Relationship - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectRelationship_Click
    
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
    dtAdmissionDate.Value = VBA.Date
    cboGender.ListIndex = &H0: iSetNo = &H3
    
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
            .Filter = "[Student ID] = " & vArrayList(&H0)
            
            Dim nHasRecords As Boolean
            
            nHasRecords = Not (.BOF And .EOF)
            
            'If the record Exists then Call Procedure in this Form to display it
            If nHasRecords Then DisplayRecord VBA.CLng(![Student ID])
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            vEditRecordID = VBA.vbNullString  'Initialize variable
            
            If nHasRecords Then If UBound(vArrayList) = &H1 Then Call MnuEdit_Click 'Call click event of Edit Menu
            If nHasRecords Then If UBound(vArrayList) = &H2 Then Call MnuDelete_Click 'Call click event of Delete Menu
            
            'Reinitialize the elements of the fixed-size array and release dynamic-array storage space.
            Erase vArrayList
            
        Else 'If a new Record is to be entered then...
            
            Call MnuNew_Click 'Call click event of New Menu
            
        End If 'Close respective IF..THEN block statement
        
        If .State = adStateOpen Then .Close 'Close the opened object and any dependent objects
        
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
    
    myTableFixedFldName(&H0) = "Student": myTableFixedFldName(&H1) = "Students"
    myTableFixedFldName(&H2) = "Student": myTableFixedFldName(&H3) = "Students"
    myTableFixedFldName(&H4) = "Tbl_Students": myTableFixedFldName(&H5) = "Tbl_ChurchMembers:Church Members"
    
    myTable = "Tbl_" & VBA.Replace(myTableFixedFldName(&H1), " ", VBA.vbNullString)
    
    Me.Caption = App.Title & " : " & myTableFixedFldName(&H3)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub Fra_Photo_Click(Index As Integer)
    ImgVirtualPhoto.Picture = ImgDBPhoto(Index).Picture
    ImgVirtualPhoto.ToolTipText = ImgDBPhoto(Index).ToolTipText
    Call PhotoClicked(ImgDBPhoto(Index), ImgVirtualPhoto, Fra_Photo(Index), VBA.IIf(Index = &H1, True, Not (Fra_Student(&H0).Enabled)), "0")
End Sub

Private Sub ImgDBPhoto_Click(Index As Integer)
    If PhotoDragged And Index = &H1 Then Exit Sub
    Call Fra_Photo_Click(Index): PhotoDragged = False
End Sub

Private Sub Lv_DblClick()
    
    'If there are no existing guardians then quit this Procedure
    If Lv.ListItems.Count = &H0 Or Not Fra_Student(&H0).Enabled Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Dim iSelItem As ListItem
    
    Set iSelItem = Lv.SelectedItem
    
    txtGuardianName.Tag = iSelItem.Text
    shpBttnAddGuardian.Tag = iSelItem.Tag
    txtGuardianName.Text = iSelItem.ListSubItems(&H1).Text
    txtRelationshipName.Tag = iSelItem.ListSubItems(&H2).Tag
    txtRelationshipName.Text = iSelItem.ListSubItems(&H2).Text
    chkPaysFees.Value = VBA.IIf(VBA.LCase$(iSelItem.ListSubItems(&H3).Text) = "yes", vbChecked, vbUnchecked)
    cmdSelectGuardian.Tag = iSelItem.ListSubItems(&H4).Text & "|"
    cmdSelectGuardian.Tag = cmdSelectGuardian.Tag & iSelItem.ListSubItems(&H5).Text
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub Lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Local Error GoTo Handle_DisplayStudentDetails_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo(&H1).Visible = False
    
    Dim nRs As New ADODB.Recordset
    
    ConnectDB , False 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT [Photo], [Guardian Name] FROM [Qry_Guardians] WHERE [Guardian ID] = " & VBA.Val(Item.Text), vAdoCNN, adOpenKeyset, adLockReadOnly
        
        If Not (.BOF And .EOF) Then
            
            'If the Record contains the Student's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display Student's Photo
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = nRs
                ImgVirtualPhoto.Refresh
                If Not VBA.IsNull(![Guardian Name]) Then ImgVirtualPhoto.ToolTipText = ![Guardian Name] & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto(&H1), Fra_Photo(&H1))
                
                'Position the Photo Frame directly below the selected item
                Fra_Photo(&H1).Top = XTab.Top + Lv.Top + Item.Top + Item.Height + 50
                
                Fra_Photo(&H1).Visible = True
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_DisplayStudentDetails:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Function
    
Handle_DisplayStudentDetails_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplayStudentDetails
    
End Sub

Private Sub Lv_LostFocus()
    If Me.ActiveControl.Name <> Fra_Photo(&H1).Name Then Fra_Photo(&H1).Visible = False
End Sub

Private Sub Lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Fra_Student(&H0).Enabled And Not Nothing Is Lv.HitTest(X, Y) Then Fra_Photo(&H1).Visible = False: Me.PopupMenu MnuLvs
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
    iDependency = CheckForRecordDependants(Me, "[Tbl_ChurchMembers] WHERE [Student ID] = " & txtSurname.Tag)
    
    'If the Records exist then...
    If iDependency Then
        
        'Warn User
        vMsgBox "The displayed Student has other Records {Church Member Records} depending on it. Delete operation aborted", vbExclamation, App.Title & " : Operation Aborted"
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM " & myTable & " WHERE [Student ID] = " & txtSurname.Tag, vAdoCNN, adOpenKeyset, adLockPessimistic
    
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
        If Fra_Student(&H0).Enabled Then txtSurname.SetFocus 'Move the focus to the specified control
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
    
    If Fra_Student(&H0).Enabled Then txtNationalID.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuLv_Click(Index As Integer)
    
    Select Case Index
        
        Case &H0: 'Remove from List
            Lv.ListItems.Remove Lv.SelectedItem.Index
        Case Else:
        
    End Select
    
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
    Call cmdGenerateAdmNo_Click
    
    IsNewRecord = True 'Denote that the displayed Record does not exist in the database
    
    txtNationalID.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuSave_Click()
On Local Error GoTo Handle_MnuSave_Click_Error
    
    'If the User has not entered the Student's Name then...
    If VBA.LenB(VBA.Trim$(txtSurname.Text)) = &H0 And VBA.LenB(VBA.Trim$(txtOtherNames.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the Student's Name", vbExclamation, App.Title & " : Name not entered"
        txtSurname.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the Student's Adm No then...
    If VBA.LenB(VBA.Trim$(txtAdmNo.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the Student's Adm No", vbExclamation, App.Title & " : Adm No not entered"
        txtAdmNo.SetFocus 'Move focus to the specified control
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
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If the Student's National ID has been entered then...
        If VBA.LenB(VBA.Trim$(txtNationalID.Text)) <> &H0 Then
            
            'Check if the entered National ID already exists in the database
            .Filter = "[Student ID] <> " & VBA.Val(txtSurname.Tag) & " AND [National ID] = '" & txtNationalID.Text & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                vBuffer(&H0) = VBA.vbNullString
                
                If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
                If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A Student {" & vBuffer(&H0) & "} with the entered National ID {" & txtNationalID.Text & "} had already been saved. Please enter a different National ID.", vbExclamation, App.Title & " : Duplicate National ID Entry"
                
                txtNationalID.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'Check if the entered Adm No already exists in the database
        .Filter = "[Student ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Adm No] = " & VBA.Val(txtAdmNo.Text)
        
        'If the National ID already exists then...
        If Not (.BOF And .EOF) Then
            
            vBuffer(&H0) = VBA.vbNullString
            
            If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
            If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Warn User and get feedback. If the User decides to abort saving then...
            vMsgBox "A Student {" & vBuffer(&H0) & "} with the entered Adm No had already been saved. Saving Aborted.", vbExclamation, App.Title & " : Duplicate Adm No", Me
            
            txtAdmNo.SetFocus 'Move the focus to the specified control
            GoTo Exit_MnuSave_Click 'Quit this Procedure
            
        End If 'Close respective IF..THEN block statement
        
        'If the Student's Full Name has been entered then...
        If VBA.LenB(VBA.Trim$(txtSurname.Text)) <> &H0 And VBA.LenB(VBA.Trim$(txtOtherNames.Text)) <> &H0 Then
            
            'Check if the entered Name already exists in the database
            .Filter = "[Student ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Surname] = '" & VBA.Replace(txtSurname.Text, "'", "''") & "' AND [Other Names] = '" & VBA.Replace(txtOtherNames.Text, "'", "''") & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                vBuffer(&H0) = VBA.vbNullString
                
                If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
                If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User and get feedback. If the User decides to abort saving then...
                If vMsgBox("A Student {" & vBuffer(&H0) & "} with the entered Name had already been saved. Proceed with saving?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Duplicate Name Entry") = vbNo Then
                    
                    txtSurname.SetFocus 'Move the focus to the specified control
                    GoTo Exit_MnuSave_Click 'Quit this Procedure
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the Student's Phone No has been entered then...
        If VBA.LenB(VBA.Trim$(txtPhoneNo.Text)) <> &H0 And VBA.Trim$(txtPhoneNo.Text) <> "254" Then
            
            vArrayList = VBA.Split(txtPhoneNo.Text, VBA.vbCrLf)
            
            'For each entered Phone No...
            For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
                
                'Check if the entered Phone No already exists in the database
                .Filter = "[Student ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Phone No] LIKE '%" & VBA.Trim$(vArrayList(vIndex(&H0))) & "%'"
                
                'If the National ID already exists then...
                If Not (.BOF And .EOF) Then
                    
                    vBuffer(&H0) = VBA.vbNullString
                    
                    If Not VBA.IsNull(![Other Names]) Then vBuffer(&H0) = ![Other Names]
                    If Not VBA.IsNull(![Surname]) Then vBuffer(&H0) = VBA.Trim$(vBuffer(&H0) & " " & ![Surname])
                    
                    Dim Ans%
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Warn User and get feedback. If the User decides to abort saving then...
                    Ans = vMsgBox("A Student {" & vBuffer(&H0) & "} with the entered Phone No {" & VBA.Trim$(vArrayList(vIndex(&H0))) & "} had already been saved. Proceed with saving? {No to display the Record}", vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title & " : Duplicate Phone No Entry")
                    
                    If Ans = vbCancel Then
                        
                        txtPhoneNo.SetFocus 'Move the focus to the specified control
                        GoTo Exit_MnuSave_Click 'Quit this Procedure
                        
                    ElseIf Ans = vbNo Then
                        
                        Call DisplayRecord(![Student ID])  'Display the Record
                        txtPhoneNo.SetFocus 'Move the focus to the specified control
                        GoTo Exit_MnuSave_Click 'Quit this Procedure
                        
                    Else
                        'Do nothing
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
            Next vIndex(&H0) 'Move to the next entered Phone No
            
        End If 'Close respective IF..THEN block statement
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[Student ID] = " & txtSurname.Tag: .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        ![Adm No] = VBA.UCase$(VBA.Trim$(txtAdmNo.Text))
        ![Admission Date] = VBA.FormatDateTime(dtAdmissionDate.Value, vbShortDate)
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
        
        'If the Student has a Photo then Assign the Photo to its field
        If ImgDBPhoto(&H0).Picture <> &H0 Then .Fields("Photo").AppendChunk sAdditionalPhoto(&H0).vDataBytes
        
        If txtMedicalDoctor.Tag <> VBA.vbNullString Then ![Physician ID] = VBA.Val(txtMedicalDoctor.Tag) Else ![Physician ID] = Null
        ![Medical History] = txtMedicalHistory.Text
        ![Allergies] = txtAllergies.Text
        ![Disabilities] = txtDisabilities.Text
        
        ![Occupation] = txtOccupation.Text
        ![Deceased] = chkDeceased.Value
        
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![Discontinued] = chkDiscontinued.Value
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        ![School ID] = School.ID
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        txtSurname.Tag = ![Student ID]
        
        .Close 'Close the opened object and any dependent objects
        
        .Open "SELECT * FROM [Tbl_StudentClasses] WHERE YEAR([Date Assigned]) = " & VBA.Year(dtAdmissionDate.Value) & " AND [Student ID] = " & VBA.Val(txtSurname.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
        If cmdClearClass.Tag <> VBA.vbNullString Then .Filter = "Student Class ID] = " & VBA.Val(cmdClearClass.Tag) Else .Filter = "[Stream ID] = " & VBA.Val(txtClassName.Tag)
        If Not (.BOF And .EOF) Then .Update Else .AddNew
        
        ![Date Assigned] = VBA.FormatDateTime(dtAdmissionDate.Value, vbShortDate)
        ![Student ID] = VBA.Val(txtSurname.Tag)
        ![Stream ID] = VBA.Val(txtClassName.Tag)
        ![Date Last Modified] = VBA.Now
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        .Close 'Close the opened object and any dependent objects
        
        .Open "SELECT * FROM [Tbl_StudentDormitories] WHERE YEAR([Date Assigned]) = " & VBA.Year(dtAdmissionDate.Value) & " AND [Student ID] = " & VBA.Val(txtSurname.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
        If cmdClearDormitory.Tag <> VBA.vbNullString Then .Filter = "Student Dormitory ID] = " & VBA.Val(cmdClearDormitory.Tag) Else .Filter = "[Dormitory ID] = " & VBA.Val(txtDormitoryName.Tag)
        If Not (.BOF And .EOF) Then .Update Else .AddNew
        
        ![Date Assigned] = VBA.FormatDateTime(dtAdmissionDate.Value, vbShortDate)
        ![Student ID] = VBA.Val(txtSurname.Tag)
        ![Dormitory ID] = VBA.Val(txtDormitoryName.Tag)
        ![Date Last Modified] = VBA.Now
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        If Lv.Checkboxes Then
            
            .Open "SELECT * FROM [Tbl_StudentGuardians]", vAdoCNN, adOpenKeyset, adLockPessimistic
            
            'For each Guardian in the list...
            For vIndex(&H0) = &H1 To Lv.ListItems.Count
                
                vArrayList = VBA.Split(Lv.ListItems(vIndex(&H0)).Tag, "|")
                
                .Filter = "[Student Guardian ID] = " & VBA.Val(vArrayList(&H0))
                
                'If the Guardian has been selected then...
                If Lv.ListItems(vIndex(&H0)).Checked Then
                    
                    'If the Guardian was already selected then open the record for update else create a new entry
                    If .RecordCount Then .Update Else .AddNew
                    
                    ![Date Assigned] = VBA.Now
                    ![Student ID] = VBA.Val(txtSurname.Tag)
                    ![Guardian ID] = VBA.Val(Lv.ListItems(vIndex(&H0)).Text)
                    ![Relationship ID] = VBA.Val(vArrayList(&H2))
                    ![School ID] = School.ID
                    ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
                    
                    .Update
                    .UpdateBatch adAffectAllChapters
                    
                    Lv.ListItems(vIndex(&H0)).ForeColor = &HC0&
                    Lv.ListItems(vIndex(&H0)).ListSubItems(&H1).ForeColor = Lv.ListItems(vIndex(&H0)).ForeColor
                    Lv.ListItems(vIndex(&H0)).ListSubItems(&H2).ForeColor = Lv.ListItems(vIndex(&H0)).ForeColor
                    Lv.ListItems(vIndex(&H0)).ListSubItems(&H3).ForeColor = Lv.ListItems(vIndex(&H0)).ForeColor
                    Lv.ListItems(vIndex(&H0)).ListSubItems(&H4).ForeColor = Lv.ListItems(vIndex(&H0)).ForeColor
                    Lv.ListItems(vIndex(&H0)).Checked = True
                    
                Else 'If the Guardian has not been selected then...
                    
                    'If it was saved then delete it
                    If .RecordCount And Lv.ListItems(vIndex(&H0)).Tag <> VBA.vbNullString Then .Delete
                    
                End If 'Close respective IF..THEN block statement
                
            Next vIndex(&H0) 'Move to the next Guardian in the list
            
            .Close 'Close the opened object and any dependent objects
            
        End If 'Close respective IF..THEN block statement
        
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
        ImgDBPhoto(&H0).Picture = Nothing
        ImgDBPhoto(&H0).ToolTipText = VBA.vbNullString
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [Student ID], YEAR([Admission Date]) AS [Reg Year], [Adm No], ([Registered Name]) AS [Student Name], [Gender] FROM [Qry_Students] ORDER BY [Admission Date] DESC, [Registered Name] ASC", txtSurname.Tag, , "1", , , "Students", , "[Qry_Students]; WHERE [Student ID] = $;4", , 6700)
    
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

Private Sub shpBttnAddGuardian_Click()
    
    If txtGuardianName.Tag = VBA.vbNullString Then
        
        'Warn User
        vMsgBox "Please select a Guardian.", vbExclamation, App.Title & " : Guardian not Selected", Me
        cmdSelectGuardian.SetFocus
        Exit Sub 'Quit this Procedure
        
    End If
    
    If txtRelationshipName.Tag = VBA.vbNullString Then
        
        'Warn User
        vMsgBox "Please select the relationship between the Student and the Guardian.", vbExclamation, App.Title & " : Relationship not Selected", Me
        Call cmdSelectRelationship_Click
        If txtRelationshipName.Tag = VBA.vbNullString Then Exit Sub 'Quit this Procedure
        
    End If
    
    Dim iLst As ListItem
    
    For vIndex(&H0) = &H1 To Lv.ListItems.Count Step &H1
        If VBA.Val(Lv.ListItems(vIndex(&H0)).Text) = VBA.Val(txtGuardianName.Tag) Then Set iLst = Lv.ListItems(vIndex(&H0)): Exit For
    Next vIndex(&H0)
    
    'If the Guardian had not been added before then...
    If Nothing Is iLst Then
        
        If Not ValidUserAccess(Me, &HF, &H1) Then Exit Sub
        
        'Create an entry for the Guardian Record
        Set iLst = Lv.ListItems.Add(Lv.ListItems.Count + &H1, , "", &H1, &H1)
        
        'Allocation ID
        iLst.Tag = Fra_Student(&H2).Tag
        
        'Guardian ID
        iLst.Text = VBA.Val(txtGuardianName.Tag)
        
        'Guardian Name
        vArrayList = VBA.Split(txtGuardianName.Text, "->")
        iLst.ListSubItems.Add iLst.ListSubItems.Count + &H1, , VBA.Trim$(vArrayList(&H0))
        
        'Relationship
        iLst.ListSubItems.Add iLst.ListSubItems.Count + &H1, , txtRelationshipName.Text
        
        'Relationship ID
        iLst.ListSubItems(iLst.ListSubItems.Count).Tag = VBA.Val(txtRelationshipName.Tag)
        
        'Pays Fees"
        iLst.ListSubItems.Add iLst.ListSubItems.Count + &H1, , VBA.IIf(chkPaysFees.Value = vbChecked, "Yes", "No")
        
        vArrayList = VBA.Split(cmdSelectGuardian.Tag, "|")
        'National ID
        iLst.ListSubItems.Add iLst.ListSubItems.Count + &H1, , vArrayList(&H0)
        
        'Location
        iLst.ListSubItems.Add iLst.ListSubItems.Count + &H1, , vArrayList(&H1)
        
    Else 'If the Guardian had been added before then...
        
        If iLst.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &HF, &H2) Then Exit Sub
        
        If vMsgBox("The selected Guardian has already been added. Replace existing Record?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Duplicate Entry", Me) = vbNo Then Exit Sub
        
        'Allocation ID
        iLst.Tag = Fra_Student(&H2).Tag
        
        'Guardian ID
        iLst.Text = VBA.Val(txtGuardianName.Tag)
        
        'Guardian Name
        vArrayList = VBA.Split(txtGuardianName.Text, "->")
        iLst.ListSubItems(&H1).Text = VBA.Trim$(vArrayList(&H0))
        
        'Relationship
        iLst.ListSubItems(&H2).Text = txtRelationshipName.Text
        
        'Relationship ID
        iLst.ListSubItems(&H2).Tag = VBA.Val(txtRelationshipName.Tag)
        
        'Pays Fees"
        iLst.ListSubItems(&H3).Text = VBA.IIf(chkPaysFees.Value = vbChecked, "Yes", "No")
        
        vArrayList = VBA.Split(cmdSelectGuardian.Tag, "|")
        
        'National ID
        iLst.ListSubItems(&H4).Text = vArrayList(&H0)
        
        'Location
        iLst.ListSubItems(&H5).Text = vArrayList(&H1)
        
    End If
    
    iLst.ForeColor = vbBlue
    iLst.ListSubItems(&H1).ForeColor = iLst.ForeColor
    iLst.ListSubItems(&H2).ForeColor = iLst.ForeColor
    iLst.ListSubItems(&H3).ForeColor = iLst.ForeColor
    iLst.ListSubItems(&H4).ForeColor = iLst.ForeColor
    iLst.Checked = True
    
    Call cmdClearGuardian_Click
    
End Sub

Private Sub ShpBttnAttachPhoto_Click()
    Call Fra_Photo_Click(&H0)
End Sub

Private Sub ShpBttnBriefNotes_Click()
    'Call Function in Mdl_Stadmis to display Notes Input Form
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_Student(&H0).Enabled, myTableFixedFldName(&H2))
End Sub

Private Sub ShpBttnClearPhoto_Click()
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, Not (Fra_Student(&H0).Enabled), "1")
End Sub

Private Sub txtAdmNo_KeyPress(KeyAscii As Integer)
    'Discard non-numeric entries
    KeyAscii = VBA.IIf((((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126))) And KeyAscii <> vbKeyReturn, KeyAscii = Empty, KeyAscii)
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii)))
    If KeyAscii = 32 Then KeyAscii = Empty
End Sub

Private Sub txtLocation_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtLocation.Text = VBA.vbNullString Or chkAutoComplete(&H3).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtLocation, "[Tbl_Students]", "Location", (TxtKeyBack = vbKeyBack))
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
    Call AutoComplete(txtOccupation, "[Tbl_Students]", "Occupation", (TxtKeyBack = vbKeyBack))
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
    Call AutoComplete(txtPostalAddress, "[Tbl_Students]", "Postal Address", (TxtKeyBack = vbKeyBack))
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
    Call AutoComplete(txtSurname, "[Tbl_Students]", "Surname", (TxtKeyBack = vbKeyBack))
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
    Call AutoComplete(txtTitle, "[Tbl_Students]", "Title", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtTitle_GotFocus()
    'Highlight contents
    txtTitle.SelStart = &H0: txtTitle.SelLength = VBA.LenB(txtTitle.Text)
End Sub

Private Sub txtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub


'------------------------------------------------------------------------------------------
'The Following Codes enable the User to Drag the displayed Photo across the Listview Item

Private Sub DragMe(CurrX!)
    
    PhotoDragged = False
    
    If DragPhoto = True Then 'Check if it is in drag mode
    
        'If the Photo is dragged past the Listview left position
        If Fra_Photo(&H1).Left < XTab.Left + Lv.Left + VBA.IIf(Not Nothing Is Lv.SmallIcons, 300, &H0) Then
        
            'Set the Photo to the left position of the sheet
            Fra_Photo(&H1).Left = XTab.Left + Lv.Left + VBA.IIf(Not Nothing Is Lv.SmallIcons, 300, &H0)
            DragPhoto = False 'Disable drag mode
            Exit Sub 'Quit this procedure
            
        Else 'If the Photo is dragged within the Listview then...
        
            'If the Photo is dragged past the Listview right position
            If Fra_Photo(&H1).Left > XTab.Left + (Lv.Width + Lv.Left) - Fra_Photo(&H1).Width Then
                
                'Set the Photo to the right position of the Listview
                Fra_Photo(&H1).Left = XTab.Left + (Lv.Width + Lv.Left) - Fra_Photo(&H1).Width
                DragPhoto = False 'Disable drag mode
                Exit Sub 'Quit this procedure
                
            Else 'If the Photo is dragged within the Listview dimensions
                
                'If the Photo frame is dragged within the Listview
                'Move the Photo to the mouse pointer position
                Fra_Photo(&H1).Move (Fra_Photo(&H1).Left + CurrX - Xaxis)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        PhotoDragged = True
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub Fra_Photo_MouseDown(Index As Integer, Button%, Shift%, X!, Y!)
    If Index <> &H1 Then Exit Sub
    DragPhoto = True 'Enable dragging of the Photo
    Screen.MousePointer = vbSizeAll 'Set the mouse pointer to cross to show drag mode
    Xaxis = X 'Capture the current X-axis position
End Sub

Private Sub Fra_Photo_MouseMove(Index As Integer, Button%, Shift%, X!, Y!)
    If Index = &H1 Then DragMe X  'Call procedure
End Sub

Private Sub Fra_Photo_MouseUp(Index As Integer, Button%, Shift%, X!, Y!)
    If Index <> &H1 Then Exit Sub
    DragPhoto = False 'Disable drag mode
    Screen.MousePointer = vbDefault 'Set mouse pointer to default pointer to portray end of processing
End Sub

Private Sub ImgDBPhoto_MouseDown(Index As Integer, Button%, Shift%, X!, Y!)
    If Index <> &H1 Then Exit Sub
    DragPhoto = True 'Enable dragging of the Photo
    'Set the mouse pointer to cross to show drag mode
    Screen.MousePointer = vbSizeAll
    Xaxis = X 'Capture the current X-axis position
End Sub

Private Sub ImgDBPhoto_MouseMove(Index As Integer, Button%, Shift%, X!, Y!)
    If Index = &H1 Then DragMe X 'Call procedure
End Sub

Private Sub ImgDBPhoto_MouseUp(Index As Integer, Button%, Shift%, X!, Y!)
    If Index <> &H1 Then Exit Sub
    DragPhoto = False 'Disable drag mode
    Screen.MousePointer = vbDefault 'Set mouse pointer to default pointer to portray end of processing
End Sub

