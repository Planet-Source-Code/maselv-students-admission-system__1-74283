VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{60D73138-4A06-4DA5-838D-E17FF732D00B}#1.0#0"; "prjXTab.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Users 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Users"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6495
   Icon            =   "Frm_Users.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Width           =   1815
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00CFE1E2&
      Height          =   1935
      Left            =   240
      TabIndex        =   83
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
      Picture         =   "Frm_Users.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   77
      Tag             =   "XY"
      ToolTipText     =   "Move First"
      Top             =   7260
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
      Picture         =   "Frm_Users.frx":0D2C
      Style           =   1  'Graphical
      TabIndex        =   78
      Tag             =   "XY"
      ToolTipText     =   "Move Previous"
      Top             =   7260
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
      Picture         =   "Frm_Users.frx":106E
      Style           =   1  'Graphical
      TabIndex        =   79
      Tag             =   "XY"
      ToolTipText     =   "Move Next"
      Top             =   7260
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
      Picture         =   "Frm_Users.frx":13B0
      Style           =   1  'Graphical
      TabIndex        =   80
      Tag             =   "XY"
      ToolTipText     =   "Move Last"
      Top             =   7260
      Width           =   375
   End
   Begin prjXTab.XTab XTab 
      Height          =   3495
      Left            =   240
      TabIndex        =   37
      Top             =   3360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      TabCount        =   4
      TabCaption(0)   =   "Addresses"
      TabContCtrlCnt(0)=   15
      Tab(0)ContCtrlCap(1)=   "txtBriefNotes"
      Tab(0)ContCtrlCap(2)=   "dtBirthDate"
      Tab(0)ContCtrlCap(3)=   "txtLocation"
      Tab(0)ContCtrlCap(4)=   "chkAutoComplete2"
      Tab(0)ContCtrlCap(5)=   "chkAutoComplete3"
      Tab(0)ContCtrlCap(6)=   "chkAutoComplete4"
      Tab(0)ContCtrlCap(7)=   "txtEmailAddress"
      Tab(0)ContCtrlCap(8)=   "txtPostalAddress"
      Tab(0)ContCtrlCap(9)=   "txtOccupation"
      Tab(0)ContCtrlCap(10)=   "lblBriefNotes"
      Tab(0)ContCtrlCap(11)=   "lblPostalAddress"
      Tab(0)ContCtrlCap(12)=   "lblEmailAddress"
      Tab(0)ContCtrlCap(13)=   "lblBirthDate"
      Tab(0)ContCtrlCap(14)=   "lblLocation"
      Tab(0)ContCtrlCap(15)=   "lblOccupation"
      TabCaption(1)   =   "Account Details"
      TabContCtrlCnt(1)=   3
      Tab(1)ContCtrlCap(1)=   "txtPosition"
      Tab(1)ContCtrlCap(2)=   "txtReportsTo"
      Tab(1)ContCtrlCap(3)=   "Fra_AccountDetails"
      TabCaption(2)   =   "Software Details"
      TabContCtrlCnt(2)=   1
      Tab(2)ContCtrlCap(1)=   "Fra_SoftwareDetails"
      TabCaption(3)   =   "Privileges"
      TabContCtrlCnt(3)=   4
      Tab(3)ContCtrlCap(1)=   "Tv"
      Tab(3)ContCtrlCap(2)=   "ShpBttnCheck0"
      Tab(3)ContCtrlCap(3)=   "ShpBttnCheck1"
      Tab(3)ContCtrlCap(4)=   "Label1"
      TabTheme        =   1
      ActiveTabBackStartColor=   16514555
      ActiveTabBackEndColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      ActiveTabForeColor=   16711680
      InActiveTabForeColor=   128
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.Frame Fra_SoftwareDetails 
         BackColor       =   &H00FFFFFF&
         Height          =   3015
         Left            =   -74880
         TabIndex        =   90
         Top             =   360
         Width           =   5775
         Begin MSComCtl2.UpDown udRank 
            Height          =   315
            Left            =   5400
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtRank"
            BuddyDispid     =   196627
            OrigLeft        =   5400
            OrigTop         =   360
            OrigRight       =   5655
            OrigBottom      =   675
            Max             =   3
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Frame Fra_ParentalControl 
            BackColor       =   &H00FFFFFF&
            Caption         =   "{Parental Control}"
            Height          =   975
            Left            =   120
            TabIndex        =   92
            Top             =   1920
            Width           =   3135
            Begin VB.CheckBox chkParentalControlON 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check1"
               Height          =   255
               Left            =   192
               TabIndex        =   73
               TabStop         =   0   'False
               ToolTipText     =   "Turn parental Control ON/OFF"
               Top             =   480
               Width           =   200
            End
            Begin MSComCtl2.DTPicker dtStartTime 
               Height          =   285
               Left            =   720
               TabIndex        =   74
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "hh:mm tt"
               Format          =   58720259
               UpDown          =   -1  'True
               CurrentDate     =   31480
            End
            Begin MSComCtl2.DTPicker dtEndTime 
               Height          =   285
               Left            =   1920
               TabIndex        =   75
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "hh:mm tt"
               Format          =   58720259
               UpDown          =   -1  'True
               CurrentDate     =   31480
            End
            Begin VB.Label lblParentalControlON 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ON:"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lblEndTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "End Time:"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   1920
               TabIndex        =   94
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblStartTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Start Time:"
               BeginProperty Font 
                  Name            =   "Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   225
               Left            =   720
               TabIndex        =   93
               Top             =   240
               Width           =   780
            End
         End
         Begin VB.TextBox txtSecurityQuestionAns 
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
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   4200
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   71
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtSecurityQuestionAns 
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
            Index           =   0
            Left            =   4200
            MaxLength       =   30
            TabIndex        =   66
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chkShowSecurityAns 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   1560
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkShowSecurityAns 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   67
            TabStop         =   0   'False
            ToolTipText     =   "Tick to Auto-Complete entries"
            Top             =   960
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkRequestPasswordOnLogon 
            BackColor       =   &H00FBFDFB&
            Caption         =   "User must change password at next logon"
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
            Height          =   495
            Left            =   3600
            TabIndex        =   76
            Top             =   2160
            Width           =   2055
         End
         Begin VB.ComboBox cboSecurityQuestion2 
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
            ItemData        =   "Frm_Users.frx":16F2
            Left            =   120
            List            =   "Frm_Users.frx":16F4
            TabIndex        =   69
            Top             =   1560
            Width           =   3735
         End
         Begin VB.ComboBox cboSecurityQuestion1 
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
            ItemData        =   "Frm_Users.frx":16F6
            Left            =   120
            List            =   "Frm_Users.frx":16F8
            TabIndex        =   64
            Top             =   960
            Width           =   3735
         End
         Begin VB.TextBox txtUserName 
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
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtPassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   3600
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   60
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtRank 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   5040
            MaxLength       =   30
            TabIndex        =   62
            Text            =   "1"
            Top             =   360
            Width           =   420
         End
         Begin VB.TextBox txtAccountName 
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
            Left            =   120
            MaxLength       =   30
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblSecurityQuestion1Ans 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your Answer:"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   3960
            TabIndex        =   65
            Top             =   720
            Width           =   990
         End
         Begin VB.Label lblSecurityQuestion2Ans 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your Answer:"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   3960
            TabIndex        =   70
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label lblSecurityQuestion2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security Question 2:"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   1485
         End
         Begin VB.Label lblSecurityQuestion1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Security Question 1:"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Name:"
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
            TabIndex        =   57
            Top             =   120
            Width           =   840
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
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
            Left            =   3600
            TabIndex        =   59
            Top             =   120
            Width           =   750
         End
         Begin VB.Label lblUserLevel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rank:"
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
            Left            =   5040
            TabIndex        =   61
            Top             =   120
            Width           =   420
         End
         Begin VB.Label lblAccountName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
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
            TabIndex        =   55
            Top             =   120
            Width           =   1095
         End
      End
      Begin MSComctlLib.TreeView Tv 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   88
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4471
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgLst"
         Appearance      =   1
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
      Begin Stadmis.ShapeButton ShpBttnCheck 
         Height          =   375
         Index           =   0
         Left            =   -71040
         TabIndex        =   87
         Top             =   480
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
         Caption         =   "Check All"
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
      Begin Stadmis.ShapeButton ShpBttnCheck 
         Height          =   375
         Index           =   1
         Left            =   -70080
         TabIndex        =   86
         Top             =   480
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
         Caption         =   "UnCheck All"
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
      Begin VB.TextBox txtBriefNotes 
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
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Tag             =   "WY"
         Top             =   2400
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtBirthDate 
         Height          =   285
         Left            =   3960
         TabIndex        =   31
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   58720259
         CurrentDate     =   31480
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
         MaxLength       =   20
         TabIndex        =   26
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoComplete 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         Caption         =   "Check1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   1800
         Value           =   1  'Checked
         Width           =   255
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
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Tag             =   "WY"
         Top             =   1200
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
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Tag             =   "WY"
         Top             =   600
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
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Tag             =   "WY"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtPosition 
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
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   39
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtReportsTo 
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
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   960
         Width           =   3030
      End
      Begin VB.Frame Fra_AccountDetails 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Account Details:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74880
         TabIndex        =   84
         Top             =   360
         Width           =   5775
         Begin VB.CommandButton cmdClearReportsTo 
            Height          =   315
            Left            =   1920
            Picture         =   "Frm_Users.frx":16FA
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Clear selected Person"
            Top             =   600
            Width           =   345
         End
         Begin VB.CommandButton cmdSelectReportsTo 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   2280
            Picture         =   "Frm_Users.frx":1A84
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Select a Person"
            Top             =   600
            Width           =   345
         End
         Begin VB.ComboBox cboUserStatus 
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
            ItemData        =   "Frm_Users.frx":200E
            Left            =   120
            List            =   "Frm_Users.frx":2021
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2400
            Width           =   1815
         End
         Begin VB.TextBox txtPINNo 
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
            TabIndex        =   48
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtBankAccNo 
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
            TabIndex        =   44
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtBankBranchName 
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
            MaxLength       =   100
            TabIndex        =   46
            Top             =   1200
            Width           =   3750
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
            Left            =   3840
            MaxLength       =   30
            TabIndex        =   52
            Top             =   1800
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
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   50
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblPosition 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position:"
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
            TabIndex        =   38
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblReportsTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reports To:"
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
            TabIndex        =   40
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblEmployeeStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Status:"
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
            TabIndex        =   53
            Top             =   2160
            Width           =   900
         End
         Begin VB.Label lblPINNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PIN No:"
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
            TabIndex        =   47
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label lblBankAccNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Acc No:"
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
            TabIndex        =   43
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblBankBranchName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Branch Name:"
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
            TabIndex        =   45
            Top             =   960
            Width           =   1395
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
            Left            =   3840
            TabIndex        =   51
            Top             =   1560
            Width           =   660
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
            Left            =   1920
            TabIndex        =   49
            Top             =   1560
            Width           =   675
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE: Double-Click an item or select and use Space bar to allow or disallow privilege"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   390
         Left            =   -74880
         TabIndex        =   89
         Top             =   360
         Width           =   3375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblBriefNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brief Notes:"
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
         TabIndex        =   35
         Tag             =   "Y"
         Top             =   2160
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
         TabIndex        =   22
         Tag             =   "Y"
         Top             =   360
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
         TabIndex        =   28
         Tag             =   "Y"
         Top             =   960
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
         TabIndex        =   30
         Top             =   960
         Width           =   780
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
         TabIndex        =   25
         Top             =   360
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
         TabIndex        =   32
         Tag             =   "Y"
         Top             =   1560
         Width           =   870
      End
   End
   Begin VB.Frame Fra_User 
      BackColor       =   &H00CFE1E2&
      Height          =   6615
      Left            =   120
      TabIndex        =   82
      Tag             =   "HW"
      Top             =   360
      Width           =   6255
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
         Picture         =   "Frm_Users.frx":2052
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
         Picture         =   "Frm_Users.frx":23EC
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
      Begin VB.CheckBox chkAutoComplete 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Tick to Auto-Complete entries"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkAutoComplete 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   5
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
         Left            =   120
         TabIndex        =   20
         Tag             =   "Y"
         Top             =   2640
         Width           =   1095
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
         ItemData        =   "Frm_Users.frx":2786
         Left            =   3120
         List            =   "Frm_Users.frx":2790
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
         Left            =   3960
         MaxLength       =   8
         TabIndex        =   3
         Top             =   480
         Width           =   2175
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
         ItemData        =   "Frm_Users.frx":27A5
         Left            =   2160
         List            =   "Frm_Users.frx":27AF
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
         TabIndex        =   6
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
         Width           =   3975
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
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
         Left            =   1320
         TabIndex        =   21
         Tag             =   "Y"
         Top             =   2640
         Width           =   1332
      End
      Begin MSComCtl2.DTPicker dtDateEmployed 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "ddd dd MMM yyyy"
         Format          =   58720259
         CurrentDate     =   31480
      End
      Begin VB.Label lblDateEmployed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Employed:"
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
         Width           =   1140
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
         TabIndex        =   2
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
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   -600
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
            Picture         =   "Frm_Users.frx":27C1
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
      TabIndex        =   81
      Tag             =   "Y"
      Top             =   7320
      Width           =   1395
   End
   Begin VB.Image ImgHeader 
      Height          =   375
      Left            =   0
      Picture         =   "Frm_Users.frx":2B5B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6375
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_Users.frx":3351
      Stretch         =   -1  'True
      Tag             =   "WY"
      Top             =   7080
      Width           =   6375
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
End
Attribute VB_Name = "Frm_Users"
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
Private myAccountChange As Boolean
Private myRecDisplayON, IsLoading, tvPopulated As Boolean

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
    
    Tv.Tag = VBA.vbNullString
    
    txtNationalID.Text = VBA.vbNullString
    txtSurname.Tag = VBA.vbNullString
    txtSurname.Text = VBA.vbNullString
    txtOtherNames.Text = VBA.vbNullString
    txtPhoneNo.Text = VBA.IIf(myRecDisplayON, VBA.vbNullString, "254")
    txtPostalAddress.Text = "P.O Box "
    txtEmailAddress.Text = VBA.vbNullString
    txtRank.Text = User.Hierarchy + &H1
    txtBriefNotes.Text = VBA.vbNullString
    dtBirthDate.Value = Null
    chkDeceased.Value = vbUnchecked
    chkDiscontinued.Value = vbUnchecked
    
    txtAccountName.Enabled = True
    TxtUserName.Text = VBA.vbNullString
    TxtPassword.Text = VBA.vbNullString
    txtPosition.Text = VBA.vbNullString
    cboSecurityQuestion1.ListIndex = -&H1
    cboSecurityQuestion2.ListIndex = -&H1
    txtSecurityQuestionAns(&H0).Text = VBA.vbNullString
    txtSecurityQuestionAns(&H1).Text = VBA.vbNullString
    txtRank.Enabled = True: udRank.Enabled = True
    
    txtBankAccNo.Text = VBA.vbNullString
    txtBankBranchName.Text = VBA.vbNullString
    txtPINNo.Text = VBA.vbNullString
    txtNSSFNo.Text = VBA.vbNullString
    txtNHIFNo.Text = VBA.vbNullString
    TxtPassword.Enabled = True
    
    cmdClearReportsTo_Click
    
    chkParentalControlON.Value = vbChecked
    chkRequestPasswordOnLogon.Enabled = True
    Fra_ParentalControl.Enabled = True
    
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
    
    Fra_User.Enabled = Not State
    
    txtPhoneNo.Locked = State
    txtPosition.Locked = State
    txtLocation.Locked = State
    txtBriefNotes.Locked = State
    txtEmailAddress.Locked = State
    txtPostalAddress.Locked = State
    Fra_SoftwareDetails.Enabled = Not State
    ShpBttnCheck(&H0).Visible = Not State
    ShpBttnCheck(&H1).Visible = Not State
    
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
    
    tvPopulated = False 'Denote that the Treeview has not been populated with selected User's privileges
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        chkShowSecurityAns(&H0).Value = VBA.IIf(VBA.Val(VBA.GetSetting(App.Title, "Settings", User.User_ID & " : Show Security Qst " & &H0, &H0)), vbChecked, vbUnchecked)
        chkShowSecurityAns(&H1).Value = VBA.IIf(VBA.Val(VBA.GetSetting(App.Title, "Settings", User.User_ID & " : Show Security Qst " & &H1, &H0)), vbChecked, vbUnchecked)
        
        .Open "SELECT * FROM [Qry_Users] WHERE [User ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'Assign the Record's primary key value
            If Not VBA.IsNull(![User ID]) Then txtSurname.Tag = ![User ID]
            dtDateEmployed.Value = ![Date Employed]
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
            
            If Not VBA.IsNull(![Account Name]) Then txtAccountName.Text = ![Account Name]
            If Not VBA.IsNull(![User Name]) Then TxtUserName.Text = ![User Name]
            If Not VBA.IsNull(![Password]) Then TxtPassword.Text = SmartDecrypt(![Password])
            If Not VBA.IsNull(![Hierarchy]) Then txtRank.Text = ![Hierarchy]: txtRank.Tag = ![Hierarchy]
            
            If Not VBA.IsNull(![Position]) Then txtPosition.Text = ![Position]
            
            If Not VBA.IsNull(![Security Question 1]) Then cboSecurityQuestion1.Text = ![Security Question 1]
            If Not VBA.IsNull(![Security Ans 1]) Then txtSecurityQuestionAns(&H0).Text = SmartDecrypt(![Security Ans 1])
            If Not VBA.IsNull(![Security Question 2]) Then cboSecurityQuestion2.Text = ![Security Question 2]
            If Not VBA.IsNull(![Security Ans 2]) Then txtSecurityQuestionAns(&H1).Text = SmartDecrypt(![Security Ans 2])
            
            Fra_ParentalControl.Enabled = (User.Hierarchy = &H0 Or (User.User_ID <> VBA.Val(txtSurname.Tag) And User.Hierarchy < VBA.Val(txtRank.Tag)))
            chkParentalControlON.Value = VBA.IIf(![Parental Control], vbChecked, vbUnchecked)
            dtStartTime.Enabled = VBA.IIf(chkParentalControlON.Value = vbChecked, True, False)
            dtEndTime.Enabled = dtStartTime.Enabled
            
            If Not VBA.IsNull(![Start Time]) Then dtStartTime.Value = ![Start Time]
            If Not VBA.IsNull(![End Time]) Then dtEndTime.Value = ![End Time]
            
            If Not VBA.IsNull(![Report To]) Then
                
                txtReportsTo.Tag = ![Report To]
                If Not VBA.IsNull(![Reports To]) Then txtReportsTo.Text = ![Reports To]
                
            End If
            
            If Not VBA.IsNull(![Bank Acc No]) Then txtBankAccNo.Text = ![Bank Acc No]
            If Not VBA.IsNull(![Bank Name]) Then txtBankBranchName.Text = ![Bank Name]
            If Not VBA.IsNull(![PIN No]) Then txtPINNo.Text = ![PIN No]
            If Not VBA.IsNull(![NSSF No]) Then txtNSSFNo.Text = ![NSSF No]
            If Not VBA.IsNull(![NHIF No]) Then txtNHIFNo.Text = ![NHIF No]
            
            chkDeceased.Value = VBA.IIf(![Deceased], vbChecked, vbUnchecked)
            If Not VBA.IsNull(![Brief Notes]) Then txtBriefNotes.Text = ![Brief Notes]
            If Not VBA.IsNull(![User Status]) Then If VBA.Trim$(![User Status]) <> VBA.vbNullString Then cboUserStatus.Text = VBA.Trim$(![User Status])
            cboUserStatus.Tag = cboUserStatus.Text
            chkRequestPasswordOnLogon.Enabled = (VBA.Val(User.User_ID) <> VBA.Val(txtSurname.Tag))
            
            TxtPassword.Enabled = ((ValidUserAccess(Me, iSetNo, &H6, , False) And User.User_ID = VBA.Val(txtSurname.Tag)) Or User.Hierarchy = &H0)
            txtAccountName.Enabled = VBA.IIf(User.Hierarchy = &H0, True, VBA.IIf(User.User_ID = VBA.Val(txtSurname.Tag), False, ValidUserAccess(Me, iSetNo, &HC, , False)))
            
            'Prevent Users from setting their own hierarchy levels
            txtRank.Enabled = Not (User.Hierarchy <> &H0 And User.User_ID = VBA.Val(txtSurname.Tag))
            
            cboSecurityQuestion1.Locked = (User.Hierarchy <> &H0 And User.User_ID <> VBA.Val(txtSurname.Tag))
            cboSecurityQuestion2.Locked = cboSecurityQuestion1.Locked
            txtSecurityQuestionAns(&H0).Locked = cboSecurityQuestion1.Locked
            txtSecurityQuestionAns(&H1).Locked = cboSecurityQuestion1.Locked
            
            udRank.Enabled = txtRank.Enabled
            
            'If the Record contains the User's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display User's Photo
                sAdditionalPhoto(&H0).vDataBytes = ![Photo]
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = vRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = txtSurname.Text & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
            End If 'Close respective IF..THEN block statement
            
            If Not VBA.IsNull(![User Privileges]) Then Tv.Tag = SmartDecrypt(![User Privileges]): Call DisplayUserPrivileges
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
    'If there are privileges displayed then...
    If Tv.Nodes.Count >= &H1 Then
        
        'If User privileges have been displayed then...
        If Tv.Nodes(&H1).Tag <> VBA.vbNullString Then
            
            'For each node...
            For vIndex(&H0) = &H1 To Tv.Nodes.Count Step &H1
                Tv.Nodes(vIndex(&H0)).Checked = (VBA.Val(VBA.Left$(Tv.Nodes(vIndex(&H0)).Tag, &H1)) <> &H0)
            Next vIndex(&H0)
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
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

Private Function DisplayPrivileges() As Boolean
On Local Error GoTo Handle_DisplayPrivileges_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim mNode(&H1) As Node
    
    With Tv
        
        .Visible = False: .Nodes.Clear
        
        Set mNode(&H0) = Tv.Nodes.Add(, , , "School", &H1, &H1) '1
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Change School Code", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Change School Type", &H1, &H1)
            If User.Hierarchy = &H0 Then Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Change School Name", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Users", &H1, &H1) '2
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Change Password", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Clear Password", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Set Life Status", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Privileges", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Own Privileges", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Assign Privileges", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Change Account Name", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Login Report", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Application Log", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Access Software Settings", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Students", &H1, &H1) '3
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Set Life Status", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Change School Details", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Parents/Guardians", &H1, &H1) '4
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Set Life Status", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Relationships", &H1, &H1) '5
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Set Hierarchy", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Student Guardian Allocation", &H1, &H1) '6
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Classes", &H1, &H1) '7
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Streams", &H1, &H1) '8
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Student Class Allocation", &H1, &H1) '9
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Dormitories", &H1, &H1) '10
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Student Dormitory Allocation", &H1, &H1) '11
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Sports", &H1, &H1) '12
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Student Sport Allocation", &H1, &H1) '13
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Clubs", &H1, &H1) '14
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Student Club Allocation", &H1, &H1) '15
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Societies", &H1, &H1) '16
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Student Society Allocation", &H1, &H1) '17
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Prefect Posts", &H1, &H1) '18
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
        Set mNode(&H0) = Tv.Nodes.Add(, , , "Prefects", &H1, &H1) '19
            
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Add", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Edit", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Delete", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can Discontinue", &H1, &H1)
            Set mNode(&H1) = Tv.Nodes.Add(mNode(&H0), tvwChild, , mNode(&H0).Text & " - Can View Report", &H1, &H1)
            mNode(&H0).Expanded = True
            
            .Nodes(&H1).Selected = True: .Nodes(&H1).EnsureVisible: .Visible = True
            
            Dim nNum&, nNum1&
            Dim AllRedChecked As Boolean
            Dim AllBlackChecked As Boolean
            Dim nArray() As String
            Dim nNode(&H1) As Node
            
            Set nNode(&H0) = .Nodes(&H1)
            
            'User.Privileges = def_Privileges
            nArray = VBA.Split(User.Privileges, "|")
            
            For nNum1 = &H1 To nNode(&H0).LastSibling.Index Step &H1
                
                If UBound(nArray) < nNum1 - &H1 Then Exit For
                
                AllBlackChecked = True: AllRedChecked = True
                Set nNode(&H1) = nNode(&H0).Child.FirstSibling
                
                'Number the privilege categories
                nNode(&H0).Text = VBA.Format$(nNum1, "00") & ". " & nNode(&H0).Text
                
                For nNum = &H1 To nNode(&H0).Children Step &H1
                    
                    If nNum > VBA.Len(nArray(nNum1 - &H1)) Then Exit For
                    nNode(&H1).ForeColor = VBA.IIf(VBA.Val(VBA.Mid$(nArray(nNum1 - &H1), nNum, &H1)) = 1, &H0&, &HFF&)
                    If AllBlackChecked And nNode(&H1).ForeColor <> &H0& Then AllBlackChecked = False
                    If AllRedChecked And nNode(&H1).ForeColor <> &HFF& Then AllRedChecked = False
                    
                    'Number the privileges
                    nNode(&H1).Text = VBA.Format$(nNum, "00") & ". " & nNode(&H1).Text
                    
                    Set nNode(&H1) = nNode(&H1).Next
                    
                Next nNum
                
                nNode(&H0).ForeColor = VBA.IIf(AllRedChecked, &HFF&, &H0&)
                nNode(&H0).ForeColor = VBA.IIf(AllBlackChecked, &H0&, &HFF0000)
                If Nothing Is nNode(&H0).Next Then Exit For
                Set nNode(&H0) = nNode(&H0).Next
                
            Next nNum1
            
    End With
    
Exit_DisplayPrivileges:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Sub-Procedure
    
Handle_DisplayPrivileges_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Diplaying Privileges - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplayPrivileges
    
End Function

Private Function DisplayUserPrivileges() As Boolean
On Local Error GoTo Handle_DisplayUserPrivileges_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim tNode As Node
    Dim tNodeA As Node
    Dim AllChecked As Boolean
    Dim tNumA&, tNumB&, tNumC&
    
    Set tNode = Tv.Nodes(&H1)
    
    'Split assigned privileges into sets
    vArrayList = VBA.Split(Tv.Tag, "|")
    vArrayListTmp = VBA.Split(User.Privileges, "|")
    
    'For each set of privileges
    For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
        
        AllChecked = True 'Default: All child nodes have been checked
        
        Set tNodeA = tNode.Child.FirstSibling
        
        'For each child node of the selected node...
        For tNumA = &H1 To tNode.Children Step &H1
            
            If VBA.Len(vArrayList(vIndex(&H0))) < tNumA Then Exit For
            
            tNodeA.ForeColor = VBA.IIf(tNodeA.ForeColor = &HFF&, &HFF&, VBA.IIf(VBA.Val(VBA.Mid$(vArrayList(vIndex(&H0)), tNumA, &H1)) = 1, &HC00000, &H0&))
            
            'Tick/Untick the child node according to the assigned privilege
            tNodeA.Tag = VBA.IIf((tNodeA.ForeColor <> &H0& And VBA.Val(VBA.Mid$(vArrayList(vIndex(&H0)), tNumA, &H1))) = 1, &H1, &H0)
            
            'If at least one child not has not been checked, hinder node from being checked
            If AllChecked And tNodeA.Tag <> &H1 Then AllChecked = False
            Set tNodeA = tNodeA.Next 'Move to the next sibling Node of a TreeView
            
        Next tNumA 'Move to the next child node
        
        'Check node when all child nodes have been checked
        tNode.ForeColor = VBA.IIf(tNode.ForeColor = &HFF&, &HFF&, VBA.IIf(AllChecked, &HC00000, &H0&))
        tNode.Tag = VBA.IIf(AllChecked, &H1, &H0)
        
        If Nothing Is tNode.Next Then Exit For
        Set tNode = tNode.Next 'Move to the next sibling Node of a TreeView
        
    Next vIndex(&H0) 'Increment counter variable by value in the Step Option
    
    If Fra_User.Enabled Then tvPopulated = True 'Denote that the Treeview has been populated with selected User's privileges
    
Exit_DisplayUserPrivileges:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Sub-Procedure
    
Handle_DisplayUserPrivileges_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Diplaying Privileges - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplayUserPrivileges
    
End Function

Private Sub chkDeceased_Click()
On Local Error GoTo Handle_chkDeceased_Click_Error
    
    If myRecDisplayON Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If chkDeceased.Value = vbChecked Then If vMsgBox("Ticking this option will denote that the displayed User is no longer alive. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then chkDeceased.Value = vbUnchecked
    
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
    
    If chkDiscontinued.Value = vbChecked Then If vMsgBox("Ticking this option will disable the User details and will not be available in other Modules. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation") = vbNo Then chkDiscontinued.Value = vbUnchecked
    
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

Private Sub chkParentalControlON_Click()
    dtStartTime.Enabled = (chkParentalControlON.Value = vbChecked)
    dtEndTime.Enabled = dtStartTime.Enabled
End Sub

Private Sub chkShowSecurityAns_Click(Index As Integer)
    txtSecurityQuestionAns(Index).PasswordChar = VBA.IIf(chkShowSecurityAns(Index).Value = vbChecked, "*", VBA.vbNullString)
    VBA.SaveSetting App.Title, "Settings", User.User_ID & " : Show Security Qst " & Index, chkShowSecurityAns(Index).Value
End Sub

Private Sub ShpBttnCheck_Click(Index As Integer)
On Local Error GoTo Handle_ShpBttnCheck_Click_Error
    
    If Not Fra_User.Enabled Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    For vIndex(&H0) = &H1 To Tv.Nodes.Count Step &H1
        Tv.Nodes(vIndex(&H0)).ForeColor = VBA.IIf(Tv.Nodes(vIndex(&H0)).ForeColor = &HFF&, Tv.Nodes(vIndex(&H0)).ForeColor, VBA.IIf(Index = &H0, &HC00000, &H0&))
    Next vIndex(&H0)
    
Exit_ShpBttnCheck_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_ShpBttnCheck_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & VBA.IIf(Index = &H0, "Un", VBA.vbNullString) & "Ticking Privileges - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnCheck_Click
    
End Sub

Private Sub cmdClearReportsTo_Click()
    txtReportsTo.Tag = VBA.vbNullString: txtReportsTo.Text = VBA.vbNullString:
End Sub

Private Sub CmdMoveRec_Click(Index As Integer)
On Local Error GoTo Handle_CmdMoveRec_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Navigate through Records in the specified table
    myRecIndex = NavigateToRec(Me, "SELECT [User ID] FROM [Qry_Users] WHERE [Hierarchy] >= " & User.Hierarchy & " ORDER BY [Hierarchy] ASC, [Registered Name] ASC", "User ID", Index, myRecIndex)
    
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

Private Sub cmdSelectReportsTo_Click()
On Local Error GoTo Handle_cmdSelectReportsTo_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    vBuffer(&H0) = VBA.IIf(txtSurname.Tag <> VBA.vbNullString, " WHERE [User ID] <> " & VBA.Val(txtSurname.Tag), VBA.vbNullString)
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [User ID], [Registered Name], [Gender], [Position] FROM [Qry_Users]" & vBuffer(&H0) & " ORDER BY [Hierarchy] ASC, [Registered Name] ASC;", TxtUserName.Tag, "1;2;4", "1", , , "Users", , "[Qry_Users]; WHERE [User ID] = $;1;User Name", , 8500)
    
    'If a Record has not been selected then Branch to the specified Label
    If vBuffer(&H0) = VBA.vbNullString Then GoTo Exit_cmdSelectReportsTo_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    txtReportsTo.Tag = vArrayList(&H0)
    txtReportsTo.Text = vArrayList(&H1) & " - " & vArrayList(&H2)
    
    If Fra_User.Enabled Then txtBankAccNo.SetFocus 'Move the focus to the specified control
    
Exit_cmdSelectReportsTo_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectReportsTo_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectReportsTo_Click
    
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
    
    Call DisplayPrivileges
    
    'Set defaults
    iSetNo = &H2
    cboGender.ListIndex = &H0
    cboMaritalStatus.ListIndex = &H0
    dtDateEmployed.Value = VBA.Date
    udRank.Min = User.Hierarchy + &H1
    chkParentalControlON.Value = vbChecked
    dtStartTime.Enabled = (chkParentalControlON.Value = vbChecked)
    dtEndTime.Enabled = dtStartTime.Enabled
    
    'Fill Security Questions
    
    vArrayList = VBA.Split(def_SecurityQuestions1, "|")
    
    For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
        cboSecurityQuestion1.AddItem vArrayList(vIndex(&H0))
    Next vIndex(&H0)
    
    vArrayList = VBA.Split(def_SecurityQuestions2, "|")
    
    For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
        cboSecurityQuestion2.AddItem vArrayList(vIndex(&H0))
    Next vIndex(&H0)
    
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
            .Filter = "[User ID] = " & vArrayList(&H0)
            
            'If the record Exists then Call Procedure in this Form to display it
            If Not (.BOF And .EOF) Then DisplayRecord VBA.CLng(![User ID])
            
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
    
    myTableFixedFldName(&H0) = "User": myTableFixedFldName(&H1) = "Users"
    myTableFixedFldName(&H2) = "User": myTableFixedFldName(&H3) = "Users"
    myTableFixedFldName(&H4) = "Tbl_Users": myTableFixedFldName(&H5) = "Tbl_Login:Login"
    
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

Private Sub lblPassword_DblClick()
    If User.Hierarchy <> &H0 Then Exit Sub
    TxtPassword.PasswordChar = VBA.vbNullString: TxtPassword.SetFocus
End Sub

Private Sub MnuDelete_Click()
On Local Error GoTo Handle_MnuDelete_Click_Error
    
    'If the User is not allowed to execute this operation then quit this procedure
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
    
    Dim iDependencyTitle As String
    Dim iDependency As Boolean
    Dim iDependencyArray() As String
    Dim iDependencyArrayTmp() As String
    
    iDependencyArray = VBA.Split(myTableFixedFldName(&H5), "|")
    
    'For each entry...
    For vIndex(&H0) = &H0 To UBound(iDependencyArray) Step &H1
        
        iDependencyArrayTmp = VBA.Split(iDependencyArray(vIndex(&H0)), ";")
        
        'Check if there are Records in other tables depending on the displayed Record
        iDependency = CheckForRecordDependants(Me, iDependencyArrayTmp(&H0) & " WHERE [User ID] = " & txtSurname.Tag)
        
        'If the Record is depended upon then Quit this FOR..LOOP block statement
        If iDependency Then Exit For
        
    Next vIndex(&H0) 'Move to the next entry
    
    'If the Records exist then...
    If iDependency Then
        
        'Warn User
        vMsgBox "The displayed User has other Records {" & iDependencyArrayTmp(&H1) & "} depending on it. Delete operation aborted", vbExclamation, App.Title & " : Operation Aborted"
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM " & myTable & " WHERE [User ID] = " & txtSurname.Tag, vAdoCNN, adOpenKeyset, adLockPessimistic
    
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
    
    'If the User is not allowed to execute this operation then quit this procedure
    If Not ValidUserAccess(Me, iSetNo, &H2) Then Exit Sub
    
    'If the User has not selected any existing Record then...
    If VBA.LenB(VBA.Trim$(txtSurname.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed"
        If Fra_User.Enabled Then txtSurname.SetFocus 'Move the focus to the specified control
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
    Fra_SoftwareDetails.Enabled = (User.Hierarchy <= VBA.Val(txtRank.Text))
    
    IsNewRecord = False 'Denote that the displayed Record exists in the database
    
    If Fra_User.Enabled Then txtNationalID.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuNew_Click()
    
    'If the User is not allowed to execute this operation then quit this procedure
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
    
    Dim tNode As Node
    Dim tNodeA As Node
    Dim AllChecked As Boolean
    Dim tNumA&, tNumB&, tNumC&
    
    Set tNode = Tv.Nodes(&H1)
    
    'Split assigned privileges into sets
    vArrayListTmp = VBA.Split(User.Privileges, "|")
    
    'For each set of privileges
    For vIndex(&H0) = &H0 To UBound(vArrayListTmp) Step &H1
        
        Set tNodeA = tNode.Child.FirstSibling
        
        'For each child node of the selected node...
        For tNumA = &H1 To tNode.Children Step &H1
            
            tNodeA.Tag = &H0 & VBA.IIf(User.Hierarchy = &H0, &H1, &H0)
            If UBound(vArrayListTmp) >= vIndex(&H0) Then If VBA.Len(vArrayListTmp(vIndex(&H0))) >= tNumA Then tNodeA.Tag = VBA.IIf(User.Hierarchy = &H0, &H1, VBA.Val(VBA.Mid$(vArrayListTmp(vIndex(&H0)), tNumA, &H1)))
            tNodeA.Tag = &H0 & tNodeA.Tag: tNodeA.Checked = False
            tNodeA.ForeColor = VBA.IIf(VBA.Val(VBA.Right$(tNodeA.Tag, &H1)) = &H0 And User.Hierarchy <> &H0, &HFF&, tNodeA.ForeColor)
            Set tNodeA = tNodeA.Next 'Move to the next sibling Node of a TreeView
            
        Next tNumA 'Move to the next child node
        
        'Check node when all child nodes have been checked
        tNode.ForeColor = VBA.IIf(VBA.Replace(vArrayListTmp(vIndex(&H0)), "0", "") = "" And User.Hierarchy <> &H0, &HFF&, VBA.IIf(AllChecked, &HC00000, &H0&))
        tNode.Checked = False: tNode.Tag = VBA.CInt(tNode.Checked) & VBA.IIf(VBA.Replace(vArrayListTmp(vIndex(&H0)), "1", "") = "", "1", "0")
        
        Set tNode = tNode.Next 'Move to the next sibling Node of a TreeView
        
    Next vIndex(&H0) 'Increment counter variable by value in the Step Option
    
    IsNewRecord = True 'Denote that the displayed Record does not exist in the database
    
    txtNationalID.SetFocus 'Move focus to the specified control
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuSave_Click()
On Local Error GoTo Handle_MnuSave_Click_Error
    
    'If the User has not entered the User's Name then...
    If VBA.LenB(VBA.Trim$(txtSurname.Text)) = &H0 And VBA.LenB(VBA.Trim$(txtOtherNames.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the User's Name", vbExclamation, App.Title & " : Name not entered"
        txtSurname.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the User's Name then...
    If Not VBA.IsNumeric(VBA.Replace(txtPhoneNo.Text, VBA.vbCrLf, VBA.vbNullString)) And txtPhoneNo.Text <> VBA.vbNullString Then
        
        'Warn User
        vMsgBox "Please enter numeric Phone Numbers", vbExclamation, App.Title & " : Name not entered"
        txtPhoneNo.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not selected Group Account then...
    If txtAccountName.Text = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please specify the User's Account Name.", vbExclamation, App.Title & " : Record not Selected", Me
        txtAccountName.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not specified the User Name then...
    If TxtUserName.Text = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please enter the User Name.", vbExclamation, App.Title & " : Record not Selected", Me
        TxtUserName.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the entered User Name is less than the required length then...
    If VBA.Len(TxtUserName.Text) < SoftwareSetting.Min_Username_Characters Then
        
        'Warn User
        vMsgBox "Please specify the Account User Name of a minimum of " & SoftwareSetting.Min_Username_Characters & " characters.", vbInformation, App.Title & " : Blank UserName Entry", Me
        TxtUserName.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not specified the User Name then...
    If TxtPassword.Text = VBA.vbNullString Then
        
        If User.User_ID <> VBA.Val(txtSurname.Tag) And Not (User.Hierarchy = &H0) And Not ValidUserAccess(Me, iSetNo, &H2) Then
            
            'Confirm if entered User should specify the account password. If not then...
            If vMsgBox("User's Account Password has not been specified. Allow User to specify on next logon?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Blank Password Entry", Me) = vbNo Then
                
                'Inform User to specify on behalf of the other User
                vMsgBox "Please specify the User's Password.", vbInformation, App.Title & " : Blank Password Entry", Me
                TxtPassword.SetFocus 'Move the focus to the specified control
                Exit Sub 'Quit this saving procedure
                
            Else 'If the entered User should specify the account password
                
                'Enable password request when the entered User logs in
                chkRequestPasswordOnLogon.Value = vbChecked
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    'If the entered password length is less than the required value then...
    If TxtPassword.Text <> VBA.vbNullString And VBA.Len(TxtPassword.Text) < SoftwareSetting.Min_User_Password_Characters Then
        
        'Warn User
        vMsgBox "Please specify the Account Password of a minimum of " & SoftwareSetting.Min_User_Password_Characters & " characters.", vbInformation, App.Title & " : Blank Password Entry", Me
        TxtPassword.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not selected/entered the first security question then...
    If VBA.Len(VBA.Trim$(cboSecurityQuestion1.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please select the first security question.", vbInformation, App.Title & " : Blank Security Qst 1", Me
        cboSecurityQuestion1.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the first security answer then...
    If VBA.Len(VBA.Trim$(txtSecurityQuestionAns(&H0).Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the first security answer.", vbInformation, App.Title & " : Blank Security Ans 1", Me
        txtSecurityQuestionAns(&H0).SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not selected/entered the second question then...
    If VBA.Len(VBA.Trim$(cboSecurityQuestion2.Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please select the second security question.", vbInformation, App.Title & " : Blank Security Qst 2", Me
        cboSecurityQuestion2.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not entered the second security answer then...
    If VBA.Len(VBA.Trim$(txtSecurityQuestionAns(&H1).Text)) = &H0 Then
        
        'Warn User
        vMsgBox "Please enter the second security answer.", vbInformation, App.Title & " : Blank Security Ans 2", Me
        txtSecurityQuestionAns(&H1).SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If parental control has been set then...
    If chkParentalControlON.Value = vbChecked And User.User_ID <> VBA.Val(txtSurname.Tag) Then
        
        'If the End Time is earlier than Start Time then...
        If dtEndTime.Value <= dtStartTime.Value Then
            
            'Warn User
            vMsgBox "The Start Time cannot be later or equal to End Time. Please adjust to a valid time frame.", vbInformation, App.Title & " : Invalid Time Frame", Me
            dtStartTime.SetFocus 'Move the focus to the specified control
            Exit Sub 'Quit this saving procedure
            
        End If 'Close respective IF..THEN block statement
        
        'Confirm if entered User wants to set parental control on the displayed User. If not then...
        If vMsgBox("Time limits prevent Users from logging on during the specified hours. If they're logged on when their allotted time ends, they'll be automatically logged off. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Blank Password Entry", Me) = vbNo Then
            
            chkParentalControlON.SetFocus 'Move the focus to the specified control
            Exit Sub 'Quit this saving procedure
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Start the saving process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    '---------------------------------------------------------------------------------------------------
    'Format entry appropriately
    '---------------------------------------------------------------------------------------------------
    
    myRecDisplayON = True
    
    txtTitle.Text = VBA.Trim$(VBA.Replace(CapAllWords(VBA.Trim$(VBA.Replace(txtTitle.Text, " ", VBA.vbNullString))), ".", VBA.vbNullString)) & "."
    txtTitle.Text = VBA.IIf(VBA.LenB(txtTitle.Text) = &H1, VBA.vbNullString, txtTitle.Text)
    txtNationalID.Text = VBA.Format$(VBA.Trim$(txtNationalID.Text), "00000000")
    txtSurname.Text = CapAllWords(txtSurname.Text)
    
    txtOtherNames.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtOtherNames.Text, "  ", " ")))
    txtPostalAddress.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtPostalAddress.Text, "  ", " ")))
    txtPhoneNo.Text = VBA.Replace(txtPhoneNo.Text, " ", VBA.vbNullString)
    txtLocation.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtLocation.Text, "  ", " ")))
    txtEmailAddress.Text = VBA.LCase$(VBA.Trim$(VBA.Replace(txtEmailAddress.Text, " ", VBA.vbNullString)))
    txtOccupation.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtOccupation.Text, "  ", " ")))
    
    txtAccountName.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtAccountName.Text, "  ", " ")))
    TxtUserName.Text = CapAllWords(VBA.Trim$(VBA.Replace(TxtUserName.Text, "  ", " ")))
    txtPosition.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtPosition.Text, "  ", " ")))
    txtBankBranchName.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtBankBranchName.Text, "  ", " ")))
    txtPINNo.Text = CapAllWords(VBA.Replace(VBA.Trim$(VBA.Replace(txtPINNo.Text, "  ", " ")), " ", "-"))
    
    myRecDisplayON = False
    
    '---------------------------------------------------------------------------------------------------
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM [Qry_Users] WHERE [User ID] <> " & VBA.Val(txtSurname.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If the User's National ID has been entered then...
        If VBA.LenB(VBA.Trim$(txtNationalID.Text)) <> &H0 Then
            
            'Check if the entered National ID already exists in the database
            .Filter = "[National ID] = '" & txtNationalID.Text & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A User {" & ![Registered Name] & "} with the entered National ID {" & txtNationalID.Text & "} had already been saved. Please enter a different National ID.", vbExclamation, App.Title & " : Duplicate National ID Entry"
                
                txtNationalID.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's Account Name has been entered then...
        If VBA.Trim$(VBA.LenB(txtSurname.Text)) <> &H0 Then
            
            'Check if the entered Account Name already exists in the database
            .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Account Name] = '" & VBA.Replace(txtSurname.Text, "'", "''") & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User and get feedback. If the User decides to abort saving then...
                vMsgBox "A User {" & ![Registered Name] & "} with the entered Account Name had already been saved. Saving Aborted.", vbExclamation, App.Title & " : Duplicate Account Name", Me
                
                txtAccountName.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's Full Name has been entered then...
        If VBA.LenB(VBA.Trim$(txtSurname.Text)) <> &H0 And VBA.LenB(VBA.Trim$(txtOtherNames.Text)) <> &H0 Then
            
            'Check if the entered Name already exists in the database
            .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Surname] = '" & VBA.Replace(txtSurname.Text, "'", "''") & "' AND [Other Names] = '" & VBA.Replace(txtOtherNames.Text, "'", "''") & "'"
            
            'If the National ID already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User and get feedback. If the User decides to abort saving then...
                If vMsgBox("A User {" & ![Registered Name] & "} with the entered Name had already been saved. Proceed with saving?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Duplicate Name Entry") = vbNo Then
                    
                    txtSurname.SetFocus 'Move the focus to the specified control
                    GoTo Exit_MnuSave_Click 'Quit this Procedure
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's Phone No has been entered then...
        If VBA.LenB(VBA.Trim$(txtPhoneNo.Text)) <> &H0 And VBA.Trim$(txtPhoneNo.Text) <> "254" Then
            
            vArrayList = VBA.Split(txtPhoneNo.Text, VBA.vbCrLf)
            
            'For each entered Phone No...
            For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
                
                'Check if the entered Phone No already exists in the database
                .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Phone No] LIKE '%" & VBA.Trim$(vArrayList(vIndex(&H0))) & "%'"
                
                'If the National ID already exists then...
                If Not (.BOF And .EOF) Then
                    
                    Dim Ans%
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Warn User and get feedback. If the User decides to abort saving then...
                    Ans = vMsgBox("A User {" & ![Registered Name] & "} with the entered Phone No {" & VBA.Trim$(vArrayList(vIndex(&H0))) & "} had already been saved. Proceed with saving? {No to display the Record}", vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title & " : Duplicate Phone No Entry")
                    
                    If Ans = vbCancel Then
                        
                        txtPhoneNo.SetFocus 'Move the focus to the specified control
                        GoTo Exit_MnuSave_Click 'Quit this Procedure
                        
                    ElseIf Ans = vbNo Then
                        
                        Call DisplayRecord(![User ID])  'Display the Record
                        txtPhoneNo.SetFocus 'Move the focus to the specified control
                        GoTo Exit_MnuSave_Click 'Quit this Procedure
                        
                    Else
                        'Do nothing
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
            Next vIndex(&H0) 'Move to the next entered Phone No
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's Bank Account No has been entered then...
        If VBA.Trim$(txtBankAccNo.Text) <> VBA.vbNullString Then
            
            'Check if the entered Bank Account No already exists in the database
            .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [Bank Acc No] = '" & txtBankAccNo.Text & "'"
            
            'If the Bank Account No already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A User {" & ![Registered Name] & "} with the entered Bank Account No {" & txtBankAccNo.Text & "} had already been saved. Please enter a different Bank Account No.", vbExclamation, App.Title & " : Duplicate Bank Account No Entry", Me
                
                txtBankAccNo.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this saving procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's PIN No has been entered then...
        If VBA.Trim$(txtPINNo.Text) <> VBA.vbNullString Then
            
            'Check if the entered PIN No already exists in the database
            .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [PIN No] = '" & txtPINNo.Text & "'"
            
            'If the PIN No already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A User {" & ![Registered Name] & "} with the entered PIN No {" & txtPINNo.Text & "} had already been saved. Please enter a different PIN No.", vbExclamation, App.Title & " : Duplicate PIN No Entry", Me
                
                txtPINNo.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this saving procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's NSSF No has been entered then...
        If VBA.Trim$(txtNSSFNo.Text) <> VBA.vbNullString Then
            
            'Check if the entered NSSF No already exists in the database
            .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [NSSF No] = '" & txtNSSFNo.Text & "'"
            
            'If the NSSF No already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A User {" & ![Registered Name] & "} with the entered NFFS No {" & txtNSSFNo.Text & "} had already been saved. Please enter a different NSSF No.", vbExclamation, App.Title & " : Duplicate NSSF No Entry", Me
                
                txtNSSFNo.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this saving procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        'If the User's NHIF No has been entered then...
        If VBA.Trim$(txtNHIFNo.Text) <> VBA.vbNullString Then
            
            'Check if the entered NHIF No already exists in the database
            .Filter = "[User ID] <> " & VBA.Val(txtSurname.Tag) & " AND [NHIF No] = '" & txtNHIFNo.Text & "'"
            
            'If the NHIF No already exists then...
            If Not (.BOF And .EOF) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "A User {" & ![Registered Name] & "} with the entered NHIF No {" & txtNHIFNo.Text & "} had already been saved. Please enter a different NHIF No.", vbExclamation, App.Title & " : Duplicate NHIF No Entry", Me
                
                txtNHIFNo.SetFocus 'Move the focus to the specified control
                GoTo Exit_MnuSave_Click 'Quit this saving procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM [Tbl_Users]", vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If the no of Users for the selected Month exceeds the maximum No of Users as per the Software Licence then...
        If .RecordCount > SoftwareSetting.Licences.Max_Users And IsNewRecord And SoftwareSetting.Licences.Max_Users <> &H0 Then
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Warn User
            vMsgBox "The Software Licence allows a maximum of only " & SoftwareSetting.Licences.Max_Users & " Software Users. Please contact Administrator for Licence Renewal.", vbExclamation, App.Title & " : Limited Capabilities", Me
            GoTo Exit_MnuSave_Click 'Quit this saving procedure
            
        End If 'Close respective IF..THEN block statement
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[User ID] = " & txtSurname.Tag: .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        ![Account Name] = VBA.UCase$(VBA.Trim$(txtAccountName.Text))
        ![Date Employed] = VBA.FormatDateTime(dtDateEmployed.Value, vbShortDate)
        If txtNationalID.Text = VBA.vbNullString Then ![National ID] = Null Else ![National ID] = txtNationalID.Text
        ![Title] = txtTitle.Text
        ![Surname] = txtSurname.Text
        ![Other Names] = txtOtherNames.Text
        ![Gender] = cboGender.Text
        ![Marital Status] = cboMaritalStatus.Text
        ![Phone No] = VBA.IIf(VBA.Trim(txtPhoneNo.Text) = "254", VBA.vbNullString, txtPhoneNo.Text)
        ![Postal Address] = VBA.IIf(VBA.Trim(txtPostalAddress.Text) = "P.O Box", VBA.vbNullString, txtPostalAddress.Text)
        ![Location] = txtLocation.Text
        ![E-mail Address] = txtEmailAddress.Text
        ![Occupation] = txtOccupation.Text
        
        ![Security Question 1] = cboSecurityQuestion1.Text
        ![Security Ans 1] = SmartEncrypt(VBA.LCase$(VBA.Trim$(txtSecurityQuestionAns(&H0).Text)))
        ![Security Question 2] = cboSecurityQuestion2.Text
        ![Security Ans 2] = SmartEncrypt(VBA.LCase$(VBA.Trim$(txtSecurityQuestionAns(&H1).Text)))
        
        ![Account Name] = txtAccountName.Text
        ![User Name] = TxtUserName.Text
        ![Hierarchy] = txtRank.Text
        ![Password] = SmartEncrypt(TxtPassword.Text)
        
        ![Parental Control] = chkParentalControlON.Value
        
        If User.User_ID = VBA.Val(txtSurname.Tag) Or User.Hierarchy > VBA.Val(txtRank.Tag) Then
            'Do nothing
        Else
            
            If chkParentalControlON.Value = vbChecked Then
                ![Start Time] = VBA.FormatDateTime(dtStartTime.Value, vbShortTime)
                ![End Time] = VBA.FormatDateTime(dtEndTime.Value, vbShortTime)
            Else
                ![Start Time] = Null: ![End Time] = Null
            End If
            
        End If
        
        '================================================================================================================
        
        'Retrieve User Privileges
        
        vBuffer(&H0) = VBA.vbNullString 'Initialize variable
        
        Dim tNum&
        Dim tNode(&H1) As Node
        
        Set tNode(&H0) = Tv.Nodes(&H1)
        
        vBuffer(&H0) = VBA.vbNullString: vIndex(&H1) = &H0
        
        For tNum = &H1 To tNode(&H0).LastSibling.Index Step &H1
            
            Set tNode(&H1) = tNode(&H0).Child.FirstSibling
            
            For vIndex(&H0) = &H1 To tNode(&H0).Children Step &H1
                
                vBuffer(&H1) = VBA.IIf(tNode(&H1).ForeColor = &HFF&, VBA.Val(tNode(&H1).Tag), VBA.IIf(tNode(&H1).ForeColor = &HC00000, &H1, &H0))
                vBuffer(&H0) = vBuffer(&H0) & vBuffer(&H1): tNode(&H1).Tag = vBuffer(&H1)
                Set tNode(&H1) = tNode(&H1).Next
                
            Next vIndex(&H0)
            
            vBuffer(&H0) = vBuffer(&H0) & VBA.IIf(Not Nothing Is tNode(&H0).Next, "|", VBA.vbNullString)
            
            If Nothing Is tNode(&H0).Next Then Exit For
            Set tNode(&H0) = tNode(&H0).Next
            
        Next tNum
        
        ![User Privileges] = SmartEncrypt(vBuffer(&H0))
        
        '================================================================================================================
        
        ![Position] = txtPosition.Text
        If txtReportsTo.Tag = VBA.vbNullString Then ![Report To] = Null Else ![Report To] = txtReportsTo.Tag
        
        ![Bank Acc No] = txtBankAccNo.Text
        ![Bank Name] = txtBankBranchName.Text
        ![PIN No] = txtPINNo.Text
        ![NSSF No] = txtNSSFNo.Text
        ![NHIF No] = txtNHIFNo.Text
        
        If VBA.IsNull(dtBirthDate.Value) Then ![Birth Date] = Null Else ![Birth Date] = VBA.FormatDateTime(dtBirthDate.Value, vbShortDate)
        
        ![Photo] = Null
        
        'If the User has a Photo then Assign the Photo to its field
        If ImgDBPhoto.Picture <> &H0 Then .Fields("Photo").AppendChunk sAdditionalPhoto(&H0).vDataBytes
        
        ![Deceased] = chkDeceased.Value
        ![Brief Notes] = txtBriefNotes.Text
        ![User Status] = cboUserStatus.Text
        ![Saved By] = User.User_ID
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        If ![Hierarchy] = &H0 Then VBA.SaveSetting App.Title, "Settings", "Admin Lock", ![Password]
        
        'Assign the current Primary Key value of the saved Record
        txtSurname.Tag = ![User ID]
        User.Privileges = VBA.IIf(txtSurname.Tag = User.User_ID, vBuffer(&H0), User.Privileges)
        
        'Check settings if the User should change his/her password
        vIndex(&H0) = VBA.Val(VBA.GetSetting(App.Title, "Settings", VBA.Val(txtSurname.Tag) & " : Request Password change after Login", "0"))
        
        'If so then...
        If chkRequestPasswordOnLogon.Value = vbUnchecked Then
            If vIndex(&H0) = &H1 Then VBA.DeleteSetting App.Title, "Settings", VBA.Val(txtSurname.Tag) & " : Request Password change after Login"
        Else
            VBA.SaveSetting App.Title, "Settings", VBA.Val(txtSurname.Tag) & " : Request Password change after Login", &H1
        End If
        
        vBuffer(&H0) = VBA.vbNullString 'Initialize variable
        
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [User ID], [Account Name], [User Name], [Gender], [Hierarchy], [National ID], [Occupation], [Phone No] FROM [Qry_Users] WHERE [Hierarchy] >= " & User.Hierarchy & " ORDER BY [Hierarchy] DESC, [Password] ASC, [Registered Name] ASC", txtSurname.Tag, , "1", , , myTableFixedFldName(&H3), , "[Tbl_Users]; WHERE [User ID] = $", , 9300)
    
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
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, Not (Fra_User.Enabled), "0")
End Sub

Private Sub ShpBttnClearPhoto_Click()
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, Not (Fra_User.Enabled), "1")
End Sub

Private Sub Tv_DblClick()
    
    'If the Record is locked then quit this procedure
    If Not Fra_User.Enabled Then
        
        'Warn User
        vMsgBox "Please set the Record to Edit mode before allocating privileges", vbExclamation, App.Title & " : Record Locked", Me
        Exit Sub 'Quit this Sub procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User is not allowed to execute this operation then quit this procedure
    If Not ValidUserAccess(Me, iSetNo, &HB) Then Exit Sub
    
    'If no node has been selected then quit this procedure
    If Nothing Is Tv.SelectedItem Then Exit Sub
    
    Dim SelNode As Node
    
    Set SelNode = Tv.SelectedItem
    
    SelNode.Expanded = True
    
    'If the User is altering his/her own account privileges then...
    If User.User_ID = VBA.Val(txtSurname.Tag) And User.Hierarchy > &H0 Then
        
        'Warn User
        vMsgBox "You have insufficient rights to alter your own account privileges. Please contact System Administrator.", vbCritical, App.Title & " : User Privileges", Me
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User details have been locked then...
    If SelNode.ForeColor = &HFF& And User.Hierarchy > &H0 Then
        
        'Warn User
        vMsgBox "You have insufficient rights to alter the privileges of the selected item" & VBA.IIf(SelNode.Children <> &H0, " Category.", "."), vbExclamation, App.Title & " : User Privileges", Me
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User details have been locked then...
    If Not Fra_User.Enabled And User.Hierarchy > &H0 Then
        
        'Warn User
        vMsgBox "This Record has been locked. Please Edit the Record in order to define User privileges", vbExclamation, App.Title & " : User Privileges", Me
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    SelNode.ForeColor = VBA.IIf(SelNode.ForeColor = &H0&, &HC00000, &H0&)
    
    'If the selected node is a category then...
    If SelNode.Children <> &H0 Then
        
        Dim nNode(&H1) As Node
        
        Set nNode(&H0) = SelNode
        Set nNode(&H1) = nNode(&H0).Child.FirstSibling
        
        For vIndex(&H0) = &H1 To nNode(&H0).Children Step &H1
            
            nNode(&H1).ForeColor = VBA.IIf(nNode(&H1).ForeColor <> &HFF&, nNode(&H0).ForeColor, nNode(&H1).ForeColor)
            Set nNode(&H1) = nNode(&H1).Next
            
        Next vIndex(&H0)
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub Tv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call Tv_DblClick: KeyAscii = Empty
End Sub

Private Sub txtAccountName_Change()
    
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtAccountName.Text = VBA.vbNullString Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtAccountName, "[Tbl_Users]", "Account Name", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtAccountName_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtBankAccNo_KeyPress(KeyAscii As Integer)
    'Force CAPS and discard spaces
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))): If KeyAscii = vbKeySpace Then KeyAscii = Empty
End Sub

Private Sub txtBankBranchName_Change()
    
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtBankBranchName.Text = VBA.vbNullString Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtBankBranchName, "[Tbl_Users]", "Bank Name", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtBankBranchName_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = VBA.Asc(VBA.LCase$(VBA.Chr$(KeyAscii)))
    If KeyAscii = 32 Then KeyAscii = Empty
End Sub

Private Sub txtLocation_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtLocation.Text = VBA.vbNullString Or chkAutoComplete(&H3).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtLocation, "[Tbl_Users]", "Location", (TxtKeyBack = vbKeyBack))
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
    Call AutoComplete(txtOccupation, "[Tbl_Users]", "Occupation", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtOccupation_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtOccupation_Validate(Cancel As Boolean)
    If txtOccupation.Text <> txtOccupation.SelText And Not txtOccupation.Locked Then myRecDisplayON = True: txtOccupation.Text = VBA.Replace(txtOccupation.Text, txtOccupation.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtPassword_LostFocus()
    TxtPassword.PasswordChar = "*"
End Sub

Private Sub txtPhoneNo_GotFocus()
    txtPhoneNo.SelStart = VBA.LenB(txtPhoneNo.Text)
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
    'Discard non-numeric entries
    KeyAscii = VBA.IIf((((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126))) And KeyAscii <> vbKeyReturn, KeyAscii = Empty, KeyAscii)
End Sub

Private Sub txtPosition_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtPosition.Text = VBA.vbNullString Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtPosition, "[Tbl_Users]", "Position", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtPosition_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtPostalAddress_GotFocus()
    txtPostalAddress.SelStart = VBA.LenB(txtPostalAddress.Text)
End Sub

Private Sub txtPostalAddress_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtPostalAddress.Text = VBA.vbNullString Or chkAutoComplete(&H2).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtPostalAddress, "[Tbl_Users]", "Postal Address", (TxtKeyBack = vbKeyBack))
    vAutoCompleting = False
    
End Sub

Private Sub txtPostalAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    TxtKeyBack = KeyCode
End Sub

Private Sub txtPostalAddress_Validate(Cancel As Boolean)
    If txtPostalAddress.Text <> txtPostalAddress.SelText And Not txtPostalAddress.Locked Then myRecDisplayON = True: txtPostalAddress.Text = VBA.Replace(txtPostalAddress.Text, txtPostalAddress.SelText, VBA.vbNullString): myRecDisplayON = False
End Sub

Private Sub txtRank_Change()
    'Prevent empty content
    txtRank.Text = VBA.IIf(txtRank.Text = VBA.vbNullString, udRank.Min, txtRank.Text)
    txtRank.SelStart = VBA.Len(txtRank.Text)
End Sub

Private Sub txtRank_KeyPress(KeyAscii As Integer)
    'Discard non-numeric entries
    KeyAscii = VBA.IIf(((KeyAscii >= 33 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126)), KeyAscii = Empty, KeyAscii): If KeyAscii = Empty Then VBA.Beep
End Sub

Private Sub txtRank_Validate(Cancel As Boolean)
    
    txtRank.Tag = VBA.IIf(txtRank.Tag = VBA.vbNullString, User.Hierarchy + &H1, txtRank.Tag)
    
    'If the Displayed User is of a higher rank than the current User then...
    If VBA.Val(txtRank.Text) < User.Hierarchy Then
        
        'Warn User
        If vMsgBox("You cannot assign to a User a Rank higher than your own Rank. Do you want to make another entry?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title & " : Invalid Rank", Me) = vbYes Then Cancel = True
        txtRank.Text = txtRank.Tag 'Reset Rank
        
    End If 'Close respective IF..THEN block statement
    
    'If the Displayed User is of an equal rank as the current User then...
    If VBA.Val(txtRank.Text) = User.Hierarchy Then
        
        'Warn User
        If vMsgBox("You cannot " & VBA.IIf(User.User_ID = VBA.Val(txtSurname.Tag), "change your own Rank", "assign to a User with an equal Rank") & ". Do you want to make another entry?", vbExclamation + vbYesNo + vbDefaultButton2, App.Title & " : Invalid Rank", Me) = vbYes Then Cancel = True
        txtRank.Text = txtRank.Tag 'Reset Rank
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub txtSurname_Change()
     
    Static vAutoCompleting As Boolean
    
    If vAutoCompleting Or myRecDisplayON Or txtSurname.Text = VBA.vbNullString Or chkAutoComplete(&H1).Value = vbUnchecked Then Exit Sub
    vAutoCompleting = True
    Call AutoComplete(txtSurname, "[Tbl_Users]", "Surname", (TxtKeyBack = vbKeyBack))
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
    Call AutoComplete(txtTitle, "[Tbl_Users]", "Title", (TxtKeyBack = vbKeyBack))
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

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    'Allow entry of only one word
    If KeyAscii = vbKeySpace Then KeyAscii = Empty: VBA.Beep
End Sub

Private Sub txtNHIFNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then KeyAscii = Empty: Exit Sub 'Neglect spaces
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))) 'Force capital letters
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And Not (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyDelete Then KeyAscii = Empty: VBA.Beep: Exit Sub  'Neglect symbols
End Sub

Private Sub txtNSSFNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then KeyAscii = Empty: VBA.Beep: Exit Sub  'Neglect spaces
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))) 'Force capital letters
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And Not (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyDelete Then KeyAscii = Empty: VBA.Beep: Exit Sub 'Neglect symbols
End Sub

Private Sub txtPINNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then KeyAscii = Empty: VBA.Beep: Exit Sub  'Neglect spaces
    KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii))) 'Force capital letters
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And Not (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyDelete Then KeyAscii = Empty: VBA.Beep: Exit Sub  'Neglect symbols
End Sub

Private Sub XTab_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
    If iNewActiveTab <> &H3 Or tvPopulated Then Exit Sub
    bCancel = Not ValidUserAccess(Me, iSetNo, VBA.IIf(User.User_ID <> VBA.Val(txtSurname.Tag), &H9, &HA))
End Sub
