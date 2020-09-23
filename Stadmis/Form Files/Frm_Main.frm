VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "App Title : Main Switchboard"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11775
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_UserPhoto 
      BackColor       =   &H00CFE1E2&
      Height          =   1335
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   1095
      Begin VB.Image ImgUserPhoto 
         Height          =   975
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer TimerDateTime 
      Interval        =   100
      Left            =   -840
      Top             =   6720
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   5880
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Image ImgVirtualPhoto 
         Height          =   135
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image ImgDBPhoto 
         Height          =   1215
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape shpFraPhotoBorder 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   1215
      End
   End
   Begin Stadmis.AutoSizer AutoSizer 
      Left            =   -420
      Top             =   6720
      _ExtentX        =   661
      _ExtentY        =   661
      MinWidth        =   11895
      MinHeight       =   8505
   End
   Begin VB.Frame Fra_Main 
      BackColor       =   &H00FFFFFF&
      Height          =   4920
      Left            =   2760
      TabIndex        =   8
      Tag             =   "AutoSizer:WH"
      Top             =   1695
      Width           =   8895
      Begin VB.Frame Fra_Menu 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Tag             =   "AutoSizer:W"
         Top             =   600
         Width           =   8655
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   44
            ToolTipText     =   "Add new Records"
            Top             =   160
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":0ECA
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   45
            ToolTipText     =   "Edit selected Record"
            Top             =   160
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":1264
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   2
            Left            =   900
            TabIndex        =   46
            ToolTipText     =   "Delete selected Record"
            Top             =   160
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":15FE
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   3
            Left            =   1440
            TabIndex        =   47
            ToolTipText     =   "Search"
            Top             =   160
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":1B98
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   4
            Left            =   1860
            TabIndex        =   48
            ToolTipText     =   "Refresh"
            Top             =   160
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":1F32
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   5
            Left            =   2340
            TabIndex        =   49
            ToolTipText     =   "Export to MsExcel"
            Top             =   165
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":22CC
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   6
            Left            =   2760
            TabIndex        =   50
            ToolTipText     =   "Export to MsWord"
            Top             =   165
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":2666
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
         Begin Stadmis.ShapeButton ShpBttnMenu 
            Height          =   375
            Index           =   7
            Left            =   3240
            TabIndex        =   51
            ToolTipText     =   "Reports"
            Top             =   165
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   7
            ButtonStyleColors=   3
            BackColor       =   14211288
            BackColorPressed=   15715986
            BackColorHover  =   16243621
            BorderColor     =   9408398
            BorderColorPressed=   6045981
            BorderColorHover=   11632444
            Caption         =   ""
            Picture         =   "Frm_Main.frx":2A00
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
         Begin VB.Line LnMenu 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   3180
            X2              =   3180
            Y1              =   542
            Y2              =   152
         End
         Begin VB.Line LnMenu 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   2280
            X2              =   2280
            Y1              =   535
            Y2              =   145
         End
         Begin VB.Line LnMenu 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   1350
            X2              =   1350
            Y1              =   542
            Y2              =   152
         End
      End
      Begin VB.CheckBox chkShowTotals 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Show Summation Row"
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
         Height          =   204
         Left            =   1800
         TabIndex        =   22
         Tag             =   "AutoSizer:Y"
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox ChkDisplayPhoto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Display Photo"
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
         Height          =   204
         Left            =   120
         TabIndex        =   21
         Tag             =   "AutoSizer:Y"
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSComctlLib.ListView Lv 
         Height          =   3360
         Left            =   120
         TabIndex        =   9
         Tag             =   "AutoSizer:WH"
         Top             =   1320
         Width           =   8640
         _ExtentX        =   15240
         _ExtentY        =   5927
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         Icons           =   "ImgLst"
         SmallIcons      =   "ImgLst"
         ForeColor       =   12582912
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
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin VB.Label lblSelectedCategory 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Category"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Tag             =   "AutoSizer:W"
         Top             =   120
         Width           =   8700
      End
      Begin VB.Label lblSelectedCategory 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected category descriptions"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   52
         Tag             =   "AutoSizer:W"
         Top             =   360
         Width           =   8460
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3240
         TabIndex        =   23
         Tag             =   "AutoSizer:Y"
         Top             =   4680
         Width           =   1395
      End
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   -1440
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":2D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":3514
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":3F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":4DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":5562
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":58FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":67EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":6B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":7582
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":79D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":7E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_Main.frx":8278
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDeveloper 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Generated by SkulFee Software. Copyright©2012 Lexeme Kenya Ltd"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   3165
      TabIndex        =   20
      Tag             =   "AutoSizer:WY"
      Top             =   7320
      Width           =   8505
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "National ID:"
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
      Index           =   4
      Left            =   1320
      TabIndex        =   43
      Tag             =   "W"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblUserNationalID 
      BackStyle       =   0  'Transparent
      Caption         =   "National ID"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1440
      TabIndex        =   42
      Tag             =   "W"
      Top             =   1400
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
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
      TabIndex        =   41
      Tag             =   "W"
      Top             =   960
      Width           =   660
   End
   Begin VB.Label lblUserGender 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2040
      TabIndex        =   40
      Tag             =   "W"
      Top             =   960
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
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
      Index           =   2
      Left            =   1320
      TabIndex        =   39
      Tag             =   "W"
      Top             =   555
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   1320
      TabIndex        =   38
      Tag             =   "W"
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Tag             =   "W"
      Top             =   310
      Width           =   1335
   End
   Begin VB.Label lblUserFullName 
      BackStyle       =   0  'Transparent
      Caption         =   "Full name"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1320
      TabIndex        =   36
      Tag             =   "W"
      Top             =   750
      Width           =   2235
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgStretcher 
      Height          =   15
      Index           =   1
      Left            =   120
      Picture         =   "Frm_Main.frx":86CA
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:H"
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label lblSchoolInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "As eagles we soar high"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   33
      Tag             =   "AutoSizer:W"
      Top             =   480
      Width           =   7425
   End
   Begin VB.Shape ShpOutline 
      Height          =   1455
      Left            =   2760
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblCurrentUser 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User logged in at"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4440
      TabIndex        =   32
      Tag             =   "AutoSizer:WY"
      Top             =   6840
      Width           =   7155
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   31
      Tag             =   "AutoSizer:Y"
      Top             =   7320
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   6
      Left            =   1920
      TabIndex        =   30
      Tag             =   "AutoSizer:Y"
      Top             =   7320
      Width           =   930
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Tag             =   "AutoSizer:Y"
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Label LblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   1320
      TabIndex        =   28
      Tag             =   "AutoSizer:Y"
      Top             =   7320
      Width           =   420
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   27
      Tag             =   "AutoSizer:Y"
      Top             =   7080
      Width           =   1020
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   1320
      TabIndex        =   26
      Tag             =   "AutoSizer:Y"
      Top             =   7080
      Width           =   1020
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
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Tag             =   "AutoSizer:Y"
      Top             =   6840
      Width           =   1140
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
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   24
      Tag             =   "AutoSizer:Y"
      Top             =   6840
      Width           =   3060
   End
   Begin VB.Line Line1 
      BorderColor     =   &H009CC1C5&
      X1              =   2760
      X2              =   2760
      Y1              =   1560
      Y2              =   120
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Website: http://www.lexeme-kenya.com"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   4320
      TabIndex        =   19
      Tag             =   "AutoSizer:W"
      Top             =   1305
      Width           =   7305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: info@lexeme-kenya.com"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   4320
      TabIndex        =   18
      Tag             =   "AutoSizer:W"
      Top             =   1035
      Width           =   7305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.O Box 1234, Nairobi, 00100, KENYA    Tel: (254) - 724 688 172   Fax: 020 123456"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   17
      Tag             =   "AutoSizer:W"
      Top             =   780
      Width           =   7290
   End
   Begin VB.Image ImgLogo 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   2760
      Picture         =   "Frm_Main.frx":985F
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblSchoolInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lexeme High School"
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
      Index           =   0
      Left            =   4320
      TabIndex        =   16
      Tag             =   "AutoSizer:W"
      Top             =   120
      Width           =   7410
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWeekDay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day 3 of 7"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1575
      TabIndex        =   15
      Top             =   2115
      Width           =   720
   End
   Begin VB.Label lblDateToday 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wed 23 Sep 2010"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1245
      TabIndex        =   14
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Label LblAMPM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PM"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   990
      TabIndex        =   13
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label LblSecond 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   765
      TabIndex        =   12
      ToolTipText     =   "The current second"
      Top             =   2070
      Width           =   225
   End
   Begin VB.Label LblMinute 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   420
      Left            =   720
      TabIndex        =   11
      ToolTipText     =   "The current minute"
      Top             =   1770
      Width           =   315
   End
   Begin VB.Label LblHour 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   870
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "The hour of the day"
      Top             =   1680
      Width           =   645
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   6
      Left            =   360
      Picture         =   "Frm_Main.frx":1FEBE
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   5
      Left            =   360
      Picture         =   "Frm_Main.frx":20448
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   4
      Left            =   360
      Picture         =   "Frm_Main.frx":21312
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   3
      Left            =   360
      Picture         =   "Frm_Main.frx":21CFC
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   2
      Left            =   360
      Picture         =   "Frm_Main.frx":22BC6
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   1
      Left            =   360
      Picture         =   "Frm_Main.frx":22F50
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   375
   End
   Begin VB.Image imgMenu 
      Height          =   375
      Index           =   0
      Left            =   360
      Picture         =   "Frm_Main.frx":232DA
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software &Users"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   6
      Left            =   840
      TabIndex        =   6
      Top             =   6240
      Width           =   1080
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Classes"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   5640
      Width           =   1140
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student &Guardians"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   5040
      Width           =   1320
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Dormitories"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   4440
      Width           =   780
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Classes"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Parents/Guardians"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Students"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   630
   End
   Begin VB.Image imgBannerA 
      Height          =   375
      Index           =   2
      Left            =   12360
      Picture         =   "Frm_Main.frx":241A4
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgBannerA 
      Height          =   375
      Index           =   1
      Left            =   12360
      Picture         =   "Frm_Main.frx":2541E
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgBannerA 
      Height          =   375
      Index           =   0
      Left            =   12360
      Picture         =   "Frm_Main.frx":26698
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   6
      Left            =   120
      Picture         =   "Frm_Main.frx":2782D
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   5
      Left            =   120
      Picture         =   "Frm_Main.frx":289C2
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "Frm_Main.frx":29B57
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   3
      Left            =   120
      Picture         =   "Frm_Main.frx":2ACEC
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   2
      Left            =   120
      Picture         =   "Frm_Main.frx":2BE81
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      Height          =   615
      Index           =   1
      Left            =   120
      Picture         =   "Frm_Main.frx":2D016
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Image imgBanner 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "Frm_Main.frx":2E1AB
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2565
   End
   Begin VB.Image imgFooter 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_Main.frx":2F425
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   6720
      Width           =   11895
   End
   Begin VB.Image imgHeader 
      Height          =   1695
      Left            =   0
      Picture         =   "Frm_Main.frx":305BA
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   11895
   End
   Begin VB.Image imgStretcher 
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "Frm_Main.frx":3174F
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Menu MnuFiles 
      Caption         =   "&File"
      Begin VB.Menu MnuFile 
         Caption         =   "&Change Password"
         Index           =   0
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuFile 
         Caption         =   "Configurations"
         Index           =   2
         Begin VB.Menu MnuConfiguration 
            Caption         =   "&Users"
            Index           =   0
         End
         Begin VB.Menu MnuConfiguration 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuConfiguration 
            Caption         =   "&Configuration Tools"
            Index           =   2
            Begin VB.Menu MnuConfigurationTool 
               Caption         =   "&Backup Utility"
               Index           =   0
            End
            Begin VB.Menu MnuConfigurationTool 
               Caption         =   "-"
               Index           =   1
            End
            Begin VB.Menu MnuConfigurationTool 
               Caption         =   "Software Settings"
               Index           =   2
            End
            Begin VB.Menu MnuConfigurationTool 
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu MnuConfigurationTool 
               Caption         =   "School Info"
               Index           =   4
            End
         End
         Begin VB.Menu MnuConfiguration 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu MnuConfiguration 
            Caption         =   "&Software Settings"
            Index           =   4
         End
         Begin VB.Menu MnuConfiguration 
            Caption         =   "&Login Report"
            Index           =   5
         End
         Begin VB.Menu MnuConfiguration 
            Caption         =   "&Application Log"
            Index           =   6
         End
      End
      Begin VB.Menu MnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuFile 
         Caption         =   "&Log off"
         Index           =   4
      End
      Begin VB.Menu MnuFile 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu MnuAdministrations 
      Caption         =   "Administration"
      Begin VB.Menu MnuAdministration 
         Caption         =   "&Students"
         Index           =   0
         Begin VB.Menu MnuStudents 
            Caption         =   "1986 Students"
            Index           =   0
         End
         Begin VB.Menu MnuStudents 
            Caption         =   "All Students"
            Index           =   1
         End
         Begin VB.Menu MnuStudents 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuStudents 
            Caption         =   "Graduated Students"
            Index           =   3
         End
         Begin VB.Menu MnuStudents 
            Caption         =   "Transfered Students"
            Index           =   4
         End
         Begin VB.Menu MnuStudents 
            Caption         =   "Discontinued Students"
            Index           =   5
         End
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "Parents/Guardians"
         Index           =   1
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "Student Guardians"
         Index           =   2
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "Classes"
         Index           =   4
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "Student Classes"
         Index           =   5
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "Dormitories"
         Index           =   7
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "Student Dormitories"
         Index           =   8
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuAdministration 
         Caption         =   "&Relationships"
         Index           =   10
      End
   End
   Begin VB.Menu MnuActivities 
      Caption         =   "Activities"
      Begin VB.Menu MnuActivity 
         Caption         =   "Sports"
         Index           =   0
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Student Sports"
         Index           =   1
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Clubs"
         Index           =   3
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Student Clubs"
         Index           =   4
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Societies"
         Index           =   6
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Student Societies"
         Index           =   7
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Prefect Posts"
         Index           =   9
      End
      Begin VB.Menu MnuActivity 
         Caption         =   "Prefects"
         Index           =   10
      End
   End
   Begin VB.Menu MnuHelps 
      Caption         =   "&Help"
      Begin VB.Menu MnuHelp 
         Caption         =   "&Help"
         Index           =   0
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&About App Title"
         Index           =   1
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&Register App"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Frm_Main"
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

Public iLvwItemDataType$, iMaxColWidths$, iSearchIDs$

'Private Class with Events
Private WithEvents SystemTray As clsSystemTray
Attribute SystemTray.VB_VarHelpID = -1

Private IsFrmLoadingComplete As Boolean
Private iRefresh%, iShiftKey%, iWindoState%
Private FrmDefinitions$, iPhotoSpecifications$
Private DragPhoto, PhotoDragged, IsFrmLoading, iDisplayingForm As Boolean
Private vQuitState&, LstLastItem&, iLastimgBannerIndex&, iLastimgBannerSelIndex&
Private ItemCategoryNo&, ItemIndex&, iSetNo&, Xaxis&, iSelColPos&, vExportFreezeCol&, iItemX&, iItemY&

Private Function DisplaySchoolDetails() As Boolean
On Local Error GoTo Handle_DisplaySchoolDetails_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    School.ID = &H0: School.Name = VBA.vbNullString
    
    'Clear Picture
    ImgVirtualPhoto.Picture = Nothing
    ImgVirtualPhoto.ToolTipText = VBA.vbNullString
    ImgDBPhoto.Picture = Nothing
    ImgDBPhoto.ToolTipText = VBA.vbNullString
    Set School.Logo = Nothing
    Erase sAdditionalPhoto
    
    ConnectDB 'Call Function in Mdl_Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Tbl_School]", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            If Not VBA.IsNull(![School Name]) Then
                
                School.Name = ![School Name]
                
                vBuffer(&H0) = VBA.GetSetting(App.Title, "Settings", "School Name", VBA.vbNullString)
                vBuffer(&H1) = SmartDecrypt(vBuffer(&H0))
                
                If vBuffer(&H0) = VBA.vbNullString And ![School Name] <> "Lexeme High School" Then
                    'Do nothing
                ElseIf VBA.InStr(vBuffer(&H1), "School") = &H0 Then
                    'Do nothing
                ElseIf vBuffer(&H1) <> ![School Name] Then
                    'Do nothing
                Else
                    GoTo SchoolInfoVerified
                End If 'Close respective IF..THEN block statement
                
                'Check If the User has insufficient privileges to perform this Operation
                If Not ValidUserAccess(Me, &H1, &H7, False, "The specified School information has not been verified by " & App.CompanyName & " personnel. Allow another User with sufficient privileges to grant you access?", True) Then
                    School.Name = "Lexeme High School"
                End If 'Close respective IF..THEN block statement
                
SchoolInfoVerified:
                
                vBuffer(&H0) = VBA.vbNullString: vBuffer(&H1) = VBA.vbNullString
                
            End If 'Close respective IF..THEN block statement
            
            'Assign the Record's primary key value
            If Not VBA.IsNull(![School ID]) Then School.ID = ![School ID]
            If Not VBA.IsNull(![School Code]) Then School.Code = ![School Code]
'            If Not VBA.IsNull(![School Type]) Then School.Type = ![School Type]
'            If Not VBA.IsNull(![Software Modules]) Then School.Modules = ![Software Modules]
'            If Not VBA.IsNull(![Student Gender]) Then School.Gender = ![Student Gender]
            If Not VBA.IsNull(![School Motto]) Then School.Motto = ![School Motto]
            If Not VBA.IsNull(![Phone No]) Then School.PhoneNo = ![Phone No]
            If Not VBA.IsNull(![Location]) Then School.Location = ![Location]
            If Not VBA.IsNull(![NSSF No]) Then School.NSSFNo = ![NSSF No]
            If Not VBA.IsNull(![NHIF No]) Then School.NHIFNo = ![NHIF No]
            If Not VBA.IsNull(![E-mail Address]) Then School.EmailAddress = ![E-mail Address]
            If Not VBA.IsNull(![Postal Address]) Then School.PostalAddress = ![Postal Address]
            If Not VBA.IsNull(![Website]) Then School.Website = ![Website]
            
            'If the Record contains the School's Logo then...
            If Not VBA.IsNull(![Logo]) Then
                
                'Display School's Logo
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Logo"
                Set ImgVirtualPhoto.DataSource = vRs
                ImgVirtualPhoto.Refresh
                Set School.Logo = ImgVirtualPhoto.Picture
                ImgVirtualPhoto.ToolTipText = School.Name & "'s Logo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
    'Show them
'    ExamModuleActivated = (VBA.InStr(";" & School.Modules & ";", ";1;") <> &H0)
'    FeesModuleActivated = (VBA.InStr(";" & School.Modules & ";", ";2;") <> &H0)
    
    'If the Inactivated modules should be displayed then...
    If (VBA.InStr(";" & School.Modules & ";", ";3;") = &H0) Or User.Hierarchy = &H0 Then
        'Do nothing
    Else 'If the Inactivated modules should not be displayed then...
        
        'Hide them
'        MnuExaminations.Visible = (VBA.InStr(";" & School.Modules & ";", ";1;") <> &H0)
'        MnuSchoolFees.Visible = (VBA.InStr(";" & School.Modules & ";", ";2;") <> &H0)
        
    End If 'Close respective IF..THEN block statement
    
Exit_DisplaySchoolDetails:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Procedure
    
Handle_DisplaySchoolDetails_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplaySchoolDetails
    
End Function

Private Function DisplayUserDetails(UserID&) As Boolean
On Local Error GoTo Handle_DisplayUserDetails_Error
    
    Dim MousePointerState%
    Dim nRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    lblUserFullName.Caption = VBA.vbNullString
    lblUserGender.Caption = VBA.vbNullString
    lblUserNationalID.Caption = VBA.vbNullString
    
    If vAdoCNN.ConnectionString = VBA.vbNullString Then Call ConnectDB(, False)
    Set nRs = New ADODB.Recordset 'Create a new instance of the recordset object
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT * FROM [Qry_Users] WHERE [User ID] = " & UserID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If the User's info exists then...
        If Not (.BOF And .EOF) Then
            
            If Not VBA.IsNull(![User Name]) Then lblUserName.Caption = ![User Name]
            If Not VBA.IsNull(![Registered Name]) Then lblUserFullName.Caption = ![Registered Name]
            If Not VBA.IsNull(![Gender]) Then lblUserGender.Caption = ![Gender]
            If Not VBA.IsNull(![National ID]) Then lblUserNationalID.Caption = ![National ID]
            
            'If the Record contains the User's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display User's Photo
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = nRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = lblUserName.Caption & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgUserPhoto, Fra_UserPhoto)
                
            End If 'Close respective IF..THEN block statement
            
        End If  'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_DisplayUserDetails:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_DisplayUserDetails_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Display User Details Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplayUserDetails
    
End Function

Private Sub AutoSizer_AfterResize()
    lblDeveloper.ZOrder &H0
End Sub

Private Sub ChkDisplayPhoto_Click()
    VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Hide Item Photos", ChkDisplayPhoto.Value
End Sub

Private Sub chkShowTotals_Click()
    VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Hide Summation Row", chkShowTotals.Value
End Sub

Private Sub Form_Activate()

    'If the Form is not Loading the Quit this Procedure
    If Not IsFrmLoading Then Exit Sub
    
    'Denote that the Form is not Loading
    IsFrmLoading = False
    
    Call TimerDateTime_Timer 'Display current System's Time
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    '---------------------------------------------------------------------------
    'Hide critical software folders
    
    Dim iFldr
    
    If vFso.FolderExists(App.Path & "\Tools\Media\Default Media") Then
        
        Set iFldr = vFso.GetFolder(App.Path & "\Tools\Media\Default Media")
        iFldr.Attributes = &H7
        
    End If
    
    If vFso.FolderExists(App.Path & "\Application Data") Then
        
        Set iFldr = vFso.GetFolder(App.Path & "\Application Data")
        iFldr.Attributes = &H7
        
    End If
    
    '---------------------------------------------------------------------------
    
    lblDeveloper.Caption = "Developed by " & App.CompanyName & ". " & App.LegalCopyright & " " & App.LegalTrademarks
    
    vArrayList = VBA.Split(VBA.GetSetting(App.Title, "Main Form", User.User_ID & " : Last Selected Category", VBA.vbNullString), ":")
    
    If UBound(vArrayList) >= &H0 Then ItemCategoryNo = VBA.Val(vArrayList(&H0))
    If UBound(vArrayList) >= &H1 Then ItemIndex = VBA.Val(vArrayList(&H1))
    
    'Set the Condition statement in order not to execute it when the Application is reloaded
    vBuffer(&H1) = VBA.GetSetting(App.Title, "Main Form", User.User_ID & " : Last Condition Statement", VBA.vbNullString)
    
    ChkDisplayPhoto.Value = VBA.IIf(VBA.Val(VBA.GetSetting(App.Title, "Main Form", User.User_ID & " : Hide Item Photos", &H0)) = &H0, vbUnchecked, vbChecked)
    chkShowTotals.Value = VBA.IIf(VBA.Val(VBA.GetSetting(App.Title, "Main Form", User.User_ID & " : Hide Summation Row", &H0)) = &H0, vbUnchecked, vbChecked)
    
    'Display Records of the last clicked Category
    Select Case ItemCategoryNo
        
        Case &H1: Call lblMenu_Click(CInt(ItemIndex))
        Case &H2: Call MnuConfiguration_Click(CInt(ItemIndex))
        Case &H3: Call MnuConfigurationTool_Click(CInt(ItemIndex))
        Case &H4: Call MnuAdministration_Click(CInt(ItemIndex))
        Case &H5: Call MnuActivity_Click(CInt(ItemIndex))
        Case Else: Call lblMenu_Click(&H0)
        
    End Select 'Close SELECT..CASE block statement
    
    If LstLastItem <= Lv.ListItems.Count And LstLastItem <> &H0 Then Lv.Visible = True: Lv.ListItems(LstLastItem).Selected = True: Lv.ListItems(LstLastItem).EnsureVisible: Lv.Visible = True: Lv.SetFocus
    
Exit_Form_Activate:
    
    'Yield execution so that the operating system can process other events.
    VBA.DoEvents
    
    '--------------------------------------------------------------------------------------------------------------------------
    'Add Software Icon to the System Tray area
    Set SystemTray = New clsSystemTray
    
    With SystemTray
        
        .Icon = Icon.Handle
        .Menu = hWnd
        .Parent = hWnd
        .TipText = Caption
        
        Call .AddIcon
        
        'If not refreshing details then...
        If vQuitState <> &H3 Then
            
            Static iExpiryDisplayed As Boolean
            
            If Not iExpiryDisplayed Then
                
                Select Case VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date)
                    
                    Case &H0: vBuffer(&H0) = "Thank you for using " & App.Title & ". The System Licence has expired today."
                    Case -&H1: vBuffer(&H0) = "Thank you for using " & App.Title & ". The System Licence expired yesterday."
                    Case Is > &H0: vBuffer(&H0) = "The System Licence will expire on " & VBA.Format$(SoftwareSetting.Licences.Expiry_Date, "ddd dd MMM yyyy hh:nn:ss AMPM") & " - Remaining " & VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) & " day" & VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) = &H1, "", "s") & ". This Software will automatically Shut down after expiring."
                    Case Else: vBuffer(&H0) = "The System Licence expired on " & VBA.Format$(SoftwareSetting.Licences.Expiry_Date, "ddd dd MMM yyyy") & "."
                     
                End Select
                
                'Warn User of the system's expiry status
                Call SystemTray.ShowBalloon(App.Title & " : Trial Version", vBuffer(&H0), NIIF_WARNING, 15000, True)
                vBuffer(&H0) = VBA.vbNullString 'Initialize variable
                
                iExpiryDisplayed = True 'Denote that the expiry period balloon has been displayed
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End With
    
    '--------------------------------------------------------------------------------------------------------------------------
    
    Lv.Visible = True: IsFrmLoadingComplete = True: vWait = False
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    'Yield execution so that the operating system can process other events.
    VBA.DoEvents
    
    Exit Sub 'Quit this Procedure
    
Handle_Form_Activate_Error:
    
    'If there are no nodes in the Treeview then resume execution at the specified Label
    If Err.Number = 35600 Then Resume Exit_Form_Activate
    
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
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Denote that the Form is Loading
    IsFrmLoading = True: vWait = True
    
    Me.Caption = App.Title & " : " & App.FileDescription & " - Developed by " & App.CompanyName
    
    'Call ApplyTheme
    Call TimerDateTime_Timer
    
    lblCurrentUser.Caption = "Current User { " & User.User_Name & " } logged in on " & VBA.Format$(User.Login_Date, "ddd dd mmm yyyy") & " at " & VBA.Format$(User.Login_Date, "hh:nn:ss AMPM")
    
    LblHour.Caption = VBA.Format$(VBA.Time$, "HH")
    LblMinute.Caption = VBA.Format$(VBA.Time$, "nn")
    LblSecond.Caption = VBA.Format$(VBA.Time$, "ss")
    LblAMPM.Caption = VBA.Format$(VBA.Time$, "AMPM")
    lblDateToday.Tag = VBA.DateSerial(VBA.Year(VBA.Date), VBA.Month(VBA.Date), VBA.Day(VBA.Date))
    lblDateToday.Caption = VBA.Format$(lblDateToday.Tag, "ddd dd MMM yyyy")
    lblWeekDay.Caption = "Day " & VBA.Weekday(lblDateToday.Tag) & " of 7"
    
    LblLicenseTo(&H1).Caption = SoftwareSetting.Licences.License_Code
    LblLicenseTo(&H3).Caption = VBA.Format$(SoftwareSetting.Licences.Expiry_Date, "ddd dd MMM yyyy hh:nn:ss AMPM") & " - Remaining " & VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) & " day" & VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) = &H1, "", "s") & ". This Software will automatically Shut down after expiring."
    LblLicenseTo(&H5).Caption = SoftwareSetting.Licences.Key
    LblLicenseTo(&H7).Caption = SoftwareSetting.Licences.Max_Users
    
    LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < 61 And Not vRegistered, True, False)
    LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
    LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
    
    'Call Function in this Form to display the School details
    Call DisplaySchoolDetails
    
    'Call Function in this Form to display the current User's details
    Call DisplayUserDetails(User.User_ID)
    
    lblSchoolInfo(&H0).Caption = School.Name
    lblSchoolInfo(&H4).Caption = School.Motto
    lblSchoolInfo(&H1).Caption = School.PostalAddress & VBA.IIf(School.PhoneNo <> VBA.vbNullString, "    Phone: " & VBA.Replace(School.PhoneNo, VBA.vbCrLf, " / "), VBA.vbNullString)
    lblSchoolInfo(&H2).Caption = "E-mail: " & VBA.Replace(School.EmailAddress, VBA.vbCrLf, " / ")
    lblSchoolInfo(&H3).Caption = "Website: " & School.Website
    Set ImgVirtualPhoto.Picture = School.Logo
    ImgVirtualPhoto.ToolTipText = School.Name & "'s Logo"
    MnuHelp(&H1).Caption = "&About " & App.Title
    
    AutoSizer.Enabled = True
    AutoSizer.GetInitialPositions
    Call AutoSizer.AutoResize
    
    'Call Procedure to Fit image to image holder
    Call FitPicTo(ImgVirtualPhoto, ImgLogo, ShpOutline)
    
    '*********************************************************************************************************
    '                   REMEMBER FORM'S LAST DIMENSION AND POSITIONING SETTINGS
    '---------------------------------------------------------------------------------------------------------
    
    'Get the last Form's dimensions and positioning
    vBuffer(&H0) = VBA.GetSetting(App.Title, "Main Form", User.User_ID & " : Form Dimensions", VBA.vbNullString)
    
    'If no settings were defined then...
    If vBuffer(&H0) = VBA.vbNullString Then
        
        'Default option
        Me.WindowState = &H2 'Maximize Form
        
    Else 'If settings were defined then...
        
        'Fit Form to previous dimensions and positioning
        
        vArrayList = VBA.Split(vBuffer(&H0), ":")
        
        If UBound(vArrayList) >= &H0 Then Me.Top = vArrayList(&H0)
        If UBound(vArrayList) >= &H1 Then Me.Height = vArrayList(&H1)
        If UBound(vArrayList) >= &H2 Then Me.Left = vArrayList(&H2)
        If UBound(vArrayList) >= &H3 Then Me.Width = vArrayList(&H3)
        
    End If 'Close respective IF..THEN block statement
    
    'Reinitialize elements and release dynamic-array storage space.
    Erase vArrayList: Erase vBuffer
    
    '---------------------------------------------------------------------------------------------------------
    '
    '*********************************************************************************************************
    
Exit_Form_Load:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Form_Load_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Form Load Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Form_Load
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    If UnloadMode = &H0 Then vQuitState = &H0
    
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me, Not VBA.IIf(vSilentClosure Or vQuitState = &H3, True, False), , VBA.IIf(vQuitState = &H1, "Are you sure you want to Log off the current Software User {" & User.User_Name & "}", ""))
    
Exit_Form_QueryUnload:
    
    vQuitState = VBA.IIf(Cancel, &H0, vQuitState)
    
    If Not Cancel Then
        
        On Error Resume Next
        
        'Remove System icon from the tray
        'SetForegroundWindow Me.hWnd
        SystemTray.RemoveIcon
        Set SystemTray = Nothing
        
        Dim Frm As Form
        
        'Unload an application properly ensuring restoration of resources
        For Each Frm In VB.Forms
            
            'Close all other open Forms in this Application without alerts, apart from this Main one
            If Frm.Name <> Me.Name Then vSilentClosure = True: Unload Frm
            
        Next Frm 'Move to the next open Form
        
        vSilentClosure = False
        
    End If
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Form_QueryUnload_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Form QueryUnload Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Form_QueryUnload
    
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
    If Me.WindowState <> &H1 Then iWindoState = Me.WindowState
    If Me.WindowState <> &H1 And Not IsFrmLoading Then VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Form Dimensions", VBA.IIf(Me.WindowState = &H2, VBA.vbNullString, Me.Top & ":" & Me.Height & ":" & Me.Left & ":" & Me.Width)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error GoTo Handle_Form_Unload_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Form Dimensions", VBA.IIf(Me.WindowState = &H2, VBA.vbNullString, Me.Top & ":" & Me.Height & ":" & Me.Left & ":" & Me.Width)
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    TimerDateTime.Enabled = False
    
    'If refreshing Form then just quit without logging out user
    If vQuitState = &H3 Then GoTo Exit_Form_Unload
    
    'If the Login report was saved then...
    If User.Login_ID <> &H0 Then
        
        'Save User's Logout details
        ConnectDB
        
        'Retrieve the Login record of the current User
        With vRs
            
            .Open "SELECT * FROM [Tbl_Login] WHERE [Login No] = " & User.Login_ID, vAdoCNN, adOpenKeyset, adLockPessimistic
            
            'If the Login Record is found then...
            If Not (.BOF And .EOF) Then
                
                'Save the Logout Information
                
                .Update
                ![Logout Date] = VBA.Now
                ![Successful Logout] = &H1
                ![School ID] = School.ID
                .Update
                
            End If 'Close respective IF..THEN block statement
            
            .Close 'Close the opened object and any dependent objects
            
        End With
        
    End If 'Close respective IF..THEN block statement
    
Exit_Form_Unload:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Select Case vQuitState
        
        Case &H0: 'Exit
            
            'Call Procedure in Mdl_Stadmis Module to free memory and system resources
            Call PerformMemoryCleanup
            
            End
            
        Case &H1: 'Log off
            
            Frm_Login.Show
            
        Case &H2: 'Nothing
        
        Case Else: 'Nothing
        
    End Select 'Close SELECT..CASE block statement
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Form_Unload_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Lv Double Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Form_Unload
    
End Sub

Private Sub imgBanner_Click(Index As Integer)
    Call lblMenu_Click(Index)
End Sub

Private Sub imgBanner_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    imgBanner(iLastimgBannerSelIndex).BorderStyle = &H0
    imgBanner(iLastimgBannerSelIndex).Picture = imgBannerA(&H0).Picture
    imgBanner(iLastimgBannerSelIndex).Width = imgStretcher(&H0).Width
    imgBanner(Index).BorderStyle = &H1
    imgBanner(Index).Width = imgStretcher(&H0).Width + 30
    imgBanner(Index).Picture = imgBannerA(&H2).Picture
    iLastimgBannerSelIndex = Index
    
End Sub

Private Sub imgBanner_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If iLastimgBannerIndex = Index Then Exit Sub
    If iLastimgBannerIndex <> Index Then imgBanner(iLastimgBannerIndex).Picture = VBA.IIf(iLastimgBannerIndex = iLastimgBannerSelIndex, imgBannerA(&H2).Picture, imgBannerA(&H0).Picture)
    If imgBanner(Index).Picture = imgBannerA(&H1).Picture Then Exit Sub
    imgBanner(Index).Picture = VBA.IIf(Index = iLastimgBannerSelIndex, imgBannerA(&H2).Picture, imgBannerA(&H1).Picture): imgBanner(Index).ZOrder &H0: imgMenu(Index).ZOrder &H0: lblMenu(Index).ZOrder &H0
    iLastimgBannerIndex = Index
    
End Sub

Private Sub ImgDBPhoto_Click()
    
    If PhotoDragged Then Exit Sub
    ImgVirtualPhoto.Picture = ImgDBPhoto.Picture
    ImgVirtualPhoto.ToolTipText = ImgDBPhoto.ToolTipText
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, True, "0")
    PhotoDragged = False
    
End Sub

Private Sub imgLogo_Click()
    ImgDBPhoto.Picture = ImgLogo.Picture
    ImgDBPhoto.ToolTipText = ImgLogo.ToolTipText
    Call PhotoClicked(ImgLogo, ImgDBPhoto, ShpOutline, True, &H3)
End Sub

Private Sub imgMenu_Click(Index As Integer)
    Call lblMenu_Click(Index)
End Sub

Private Sub lblMenu_Click(Index As Integer)
On Local Error GoTo Handle_lblMenu_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False
    
    'Initial properties & variables
    FrmDefinitions = VBA.vbNullString
    iPhotoSpecifications = VBA.vbNullString
    lblSelectedCategory(&H1).Tag = VBA.vbNullString
    lblSelectedCategory(&H0).Caption = VBA.vbNullString
    lblSelectedCategory(&H1).Caption = VBA.vbNullString
    
    'Remove all Listview Rows and Columns
    Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
    
    'Save the current Category in order to execute it when the Application is reloaded
    ItemCategoryNo = &H1: ItemIndex = Index
    
    Select Case VBA.Replace(lblMenu(Index).Caption, "&", VBA.vbNullString)
        
        Case "Students": 'Students
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Students"
            lblSelectedCategory(&H1).Caption = "Information of every Student in the Software"
            
            iSetNo = &H3
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Students
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Students
            Call FillListView(Lv, "SELECT ([Registered Name]) AS [Student Name], [Student ID], [Admission Date], [Adm No], [Gender], [National ID], [Phone No], [Location], [E-mail Address], [Postal Address], [Discontinued] FROM [Qry_Students] ORDER BY [Registered Name] ASC", "2", , "Student ID", "Discontinued:=YES", &HC0&, IsFrmLoadingComplete, , , &H2)
            iPhotoSpecifications = "[Qry_Students] WHERE [Student ID] = $;2;[Student ID];Student Name"
            
        Case "Parents/Guardians": 'Guardians
            
            iSetNo = &H4
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Student's Guardians"
            lblSelectedCategory(&H1).Caption = "Information of every Student's Guardian in the Software"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Guardians
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Guardians
            Call FillListView(Lv, "SELECT [Guardian Name], [Guardian ID], [Gender], [National ID], [Postal Address], [Location], [Phone No], [E-mail Address], [Marital Status], [Occupation], [Deceased], [Discontinued] FROM [Qry_Guardians] ORDER BY [Registered Name];", "2", , "Guardian ID", "Discontinued:=YES|Deceased:=YES", &HC0& & "|" & &H800080, IsFrmLoadingComplete, , , &H2)
            iPhotoSpecifications = "[Qry_Guardians] WHERE [Guardian ID] = $;2;[Guardian ID];Guardian Name"
            
        Case "Classes": 'Streams
            
            iSetNo = &H8
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Classes"
            lblSelectedCategory(&H1).Caption = "Information of every Class Stream in the School"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Streams
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Guardians
            Call FillListView(Lv, "SELECT ([Class Name] & ' ' & [Stream Name]) AS [Class], [Stream ID], (IIF([Capacity]=0,NULL,[Capacity])) AS [Capacity], [Discontinued], [Class Discontinued] FROM [Qry_Streams] ORDER BY [Class Level] ASC, [Stream Name] ASC", "2", , "Stream ID", "Discontinued:=YES|Class Discontinued:=YES", &HC0& & "|" & &HC0&, IsFrmLoadingComplete, VBA.IIf(chkShowTotals.Value = vbChecked, &H3, VBA.vbNullString), , &H2)
            
        Case "Dormitories": 'Dormitories
            
            iSetNo = &HA
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Dormitory"
            lblSelectedCategory(&H1).Caption = "Information of every Dormitory in the School"
            
            FrmDefinitions = "Dormitory:Dormitories:Dormitory:Dormitories:Tbl_Dormitories:Tbl_StudentDormitories:Students"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Dormitories
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Guardians
            Call FillListView(Lv, "SELECT [Dormitory Name], [Dormitory ID], (IIF([Capacity]=0,NULL,[Capacity])) AS [Capacity], [Discontinued] FROM [Tbl_Dormitories] ORDER BY [Dormitory Name] ASC", "2", , "Dormitory ID", "Discontinued:=YES", &HC0&, IsFrmLoadingComplete, VBA.IIf(chkShowTotals.Value = vbChecked, &H3, VBA.vbNullString), , &H2)
            
        Case "Student Guardians": 'Guardians assigned to Student
            
            iSetNo = &H6
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Student's Guardians"
            lblSelectedCategory(&H1).Caption = "Information of every Student's Guardian in the School"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Guardians
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H4, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Guardians
            Call FillListView(Lv, "SELECT [Guardian Name], [Student Guardian ID], [Student ID], [Guardian Gender], [Adm No], [Student Name], ([Relationship Name]) AS [Relationship], [Visitor], [Pays Fees], [Discontinued] FROM [Qry_StudentGuardians] ORDER BY [Guardian Name] ASC, [Hierarchical Level] ASC", "2;3", , "Guardian ID", "Discontinued:=YES", &HC0&, IsFrmLoadingComplete, , , &H2)
            iPhotoSpecifications = "[Qry_Guardians] WHERE [Guardian ID] IN (SELECT [Guardian ID] FROM [Qry_StudentGuardians] WHERE [Student Guardian ID] = $);2;[Guardian ID];Guardian Name"
            
        Case "Student Classes": 'Classes assigned to Student
            
            iSetNo = &H9
            
            'Save the current Category in order to execute it when the Application is reloaded
            ItemCategoryNo = &H4: ItemIndex = Index
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Student's Classes"
            lblSelectedCategory(&H1).Caption = "Information of every Student's allocated to Classes in the School"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_StudentClasses
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Classes
            Call FillListView(Lv, "SELECT YEAR([Date Assigned]) AS [Year], [Student Class ID], [Student ID], [Class], [Adm No], ([Registered Name]) AS [Student Name], [Previous Class], IIF([Previous Class] IS NULL OR TRIM([Previous Class])='','Admitted', IIF([Promoted]<>0,'Yes','No')) AS [Promoted], [Student Class Status], [Class Discontinued] FROM [Qry_StudentClasses] ORDER BY [Date Assigned] DESC, [Class Level] DESC, [Stream Name] ASC, [Registered Name] ASC", "2;3", , "Stream ID", "Student Class Status:='Active'|Class Discontinued:=YES", &HC0& & "|" & &HC0&, IsFrmLoadingComplete, , , &H2)
            iPhotoSpecifications = "[Qry_Students] WHERE [Student ID] = $;3;[Student ID];Registered Name"
            
        Case "Software Users": 'Is Software Users...
            
            iSetNo = &H2
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Software Users"
            lblSelectedCategory(&H1).Caption = "Information of all the Employees who are Users of the Software"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Users
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then GoTo Exit_lblMenu_Click 'Branch to the specified Label
            
            'Call Function in Mdl_SchooFee to display all the Users with equal or lower levels than the current User
            Call FillListView(Lv, "SELECT [User Name], [User ID], [Full Name], [Gender], [National ID], [Phone No], [Postal Address], [Location], [E-mail Address], [Marital Status], [Birth Date], [Password], [Hierarchy] AS [Rank], [User Status] FROM [Qry_Users] WHERE [Hierarchy] >= " & User.Hierarchy & " ORDER BY [Hierarchy] ASC, [User Name] ASC", "2", , "User ID", "Group Discontinued:=YES|User Status:<>''", &HC0& & "|" & &HC0&, IsFrmLoadingComplete, , , &H4)
            iPhotoSpecifications = "[Qry_Users] WHERE [User ID] = $;2;[User ID];User Name"
            
    End Select
    
Exit_lblMenu_Click:
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    Lv.Visible = True
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_lblMenu_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Category Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_lblMenu_Click
    
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBanner_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBanner_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub imgMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBanner_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub imgMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call imgBanner_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub LblHour_Change()
    
    LblAMPM.Caption = VBA.Format$(VBA.Time$, "AMPM")
    lblDateToday.Tag = VBA.DateSerial(VBA.Year(VBA.Date), VBA.Month(VBA.Date), VBA.Day(VBA.Date))
    lblDateToday.Caption = VBA.Format$(lblDateToday.Tag, "ddd dd MMM yyyy")
    lblWeekDay.Caption = "Day " & VBA.Weekday(lblDateToday.Tag) & " of 7"
    Call ShowBallonMsg
    
End Sub

Private Sub LblMinute_Change()
    LblHour.Caption = VBA.Format$(VBA.Time$, "HH")
End Sub

Private Sub LblSecond_Change()
    LblMinute.Caption = VBA.Format$(VBA.Time$, "nn")
End Sub

Private Sub lblSelectedCategory_Change(Index As Integer)
    lblSelectedCategory(&H0).Tag = VBA.vbNullString: vSearchCriteria = VBA.vbNullString
End Sub

Private Sub Lv_Click()
    Lv.MultiSelect = (lblSelectedCategory(&H0).Caption = "Examination Results" Or lblSelectedCategory(&H0).Caption = "Exam Scores")
End Sub

Private Sub Lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Static PrevCol&
    
    'Stop
    Fra_Photo.Visible = False 'Hide the Photo Frame
    'If the same column is clicked twice then toggle Sort order else sort ascending
    SortListview Lv, ColumnHeader.Position, VBA.IIf(PrevCol = ColumnHeader.Position, VBA.IIf(Lv.SortOrder = &H0, &H1, &H0), &H0)
    PrevCol = ColumnHeader.Position
    
End Sub

Private Sub Lv_KeyDown(KeyCode As Integer, Shift As Integer)
On Local Error GoTo Handle_Lv_KeyDown_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the pressed key is
    Select Case KeyCode
        
        Case vbKeyN:  'N key then...
            
            'If not Ctrl+N key then Quit this Procedure
            If KeyCode = vbKeyN And Shift <> &H2 Then GoTo Exit_Lv_KeyDown
            
            Call ShpBttnMenu_Click(&H0) 'Execute the New command
            
        Case vbKeyF2, vbKeyE:  'Function Key F2 then...
            
            'If not Ctrl+E key then Quit this Procedure
            If KeyCode = vbKeyE And Shift <> &H2 Then GoTo Exit_Lv_KeyDown
            
            Call ShpBttnMenu_Click(&H1) 'Execute the Edit command
            
        Case vbKeyDelete: 'delete key then...
            Call ShpBttnMenu_Click(&H2) 'Execute the Delete command
            
    End Select 'Close SELECT..CASE block statement
    
Exit_Lv_KeyDown:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_Lv_KeyDown_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Lv Key Down Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Lv_KeyDown
    
End Sub

Private Sub Lv_KeyPress(KeyAscii As Integer)
    
    Dim iLst As ListItem
    
    Set iLst = FindLvItemByFirstItemChar(Lv, KeyAscii, iSelColPos)
    If Not Nothing Is iLst Then Call Lv_ItemClick(iLst)
    
End Sub

Private Sub Lv_LostFocus()
    Fra_Photo.Visible = (Me.ActiveControl.Name = "ImgDBPhoto")
End Sub

Private Sub Lv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Hide the Photo when a section on the Listview which has no Record is clicked on
    If Nothing Is Lv.HitTest(X, Y) Then Fra_Photo.Visible = False: Exit Sub
    
    Dim ColDistance&
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'For each column in the Listview...
    For vIndex(&H0) = &H1 To Lv.ColumnHeaders.Count Step &H1
        
        'Get the column distance from the left side of the Listview
        ColDistance = ColDistance + Lv.ColumnHeaders(vIndex(&H0)).Width
        
        'If the distance exceeds the clicked position then pick the current column position
        If ColDistance > X Then iSelColPos = vIndex(&H0): Exit For
        
    Next vIndex(&H0) 'Move to the next column
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Lv_DblClick()
On Local Error GoTo Handle_Lv_DblClick_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the User has selected a Record then...
    If Not Nothing Is Lv.SelectedItem Then
        
        'Assign ID of Record to be displayed
        vEditRecordID = Lv.SelectedItem.ListSubItems(&H1).Text & "|||"
        
    End If 'Close respective IF..THEN block statement
    
    'If the clicked Item is the Function Row then Quit this Procedure
    If Lv.SelectedItem.Tag = "Function" Then GoTo Exit_Lv_DblClick
    
    Call ShpBttnMenu_Click(&H0) 'Call NEW
    
Exit_Lv_DblClick:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_Lv_DblClick_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Lv Double Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Lv_DblClick
    
End Sub

'This Procedure display the Selected Record's Photo if available
Private Sub Lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Local Error GoTo Handle_Lv_ItemClick_Error
    
    If iDisplayingForm Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    iItemX = Item.Index: Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'If the Photo Table has not been specified or the Photo shouldn't be displayed then Quit this Procedure
    If iPhotoSpecifications = VBA.vbNullString Or ChkDisplayPhoto.Value = vbUnchecked Or Item.Text = VBA.vbNullString Then GoTo Exit_Lv_ItemClick
    
    Dim iPhotoFld$
    Dim SelFldItem$
    Dim iArray() As String
    Dim nRs As New ADODB.Recordset
        
    iArray = VBA.Split(VBA.Replace(iPhotoSpecifications, ";", "|"), "|")
    
    iPhotoFld = "Photo" 'Assign default name of Photo field
    
    If iArray(&H1) = &H1 Then SelFldItem = Item.Text Else SelFldItem = Item.ListSubItems(VBA.CLng(iArray(&H1)) - &H1).Text
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB(, False)  'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set nRs = New ADODB.Recordset
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT * FROM " & VBA.Replace(VBA.Trim$(" " & iArray(&H0)), "$", SelFldItem), vAdoCNN, adOpenKeyset, adLockReadOnly
        
DisplayPhoto:
        
        ImgVirtualPhoto.Picture = Nothing: ImgDBPhoto.Picture = Nothing
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'If the Record contains the Person's Photo then...
            If Not VBA.IsNull(nRs(iPhotoFld)) Then
                
                'Display Person's Photo
                ImgVirtualPhoto.DataField = iPhotoFld
                Set ImgVirtualPhoto.DataSource = nRs
                ImgVirtualPhoto.Refresh
                
                ImgVirtualPhoto.ToolTipText = nRs(VBA.Replace(VBA.Replace(iArray(&H2), "[", ""), "]", "")) & "'s " & iPhotoFld
                
                If UBound(iArray) >= &H3 Then If iArray(&H3) <> VBA.vbNullString Then If Not VBA.IsNull(nRs(iArray(&H3))) Then ImgVirtualPhoto.ToolTipText = nRs(iArray(&H3)) & "'s " & iPhotoFld
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
                'Position the Photo Frame directly below the selected item
                Fra_Photo.Top = Fra_Main.Top + Lv.Top + Item.Top + Item.Height + 50
                
                Fra_Photo.ZOrder &H0: Fra_Photo.Visible = True 'Display the Photo Frame
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_Lv_ItemClick:
    
    'Re-initialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vIndex
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Lv_ItemClick_Error:
    
    If Err.Number = 3265 Then If iPhotoFld <> "Logo" Then iPhotoFld = "Logo": Resume DisplayPhoto Else Resume Exit_Lv_ItemClick
    
    'If Err.Number = 13 Then Resume Exit_Lv_ItemClick
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Item Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_Lv_ItemClick
    
End Sub

Private Sub MnuActivity_Click(Index As Integer)
On Local Error GoTo Handle_MnuActivity_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'Determines whether double-click on Listview loads Records under the selected item or not
    vBuffer(&H1) = VBA.vbNullString  'Initialize variable
    vBuffer(&H2) = VBA.vbNullString  'Initialize variable
    iPhotoSpecifications = VBA.vbNullString  'Initialize variable
    
    'Save the current Category in order to execute it when the Application is reloaded
    ItemCategoryNo = &H5: ItemIndex = Index
    
    If Index = &H9 Then
        
        'Initial variable
        lblSelectedCategory(&H1).Tag = VBA.vbNullString
        
        'Display Category Title
        lblSelectedCategory(&H0).Caption = "Prefect Posts"
        lblSelectedCategory(&H1).Caption = "Information of all the Prefect Posts"
        
        'Remove all Listview Rows and Columns
        Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
        
        FrmDefinitions = "Prefect Post:Prefect Posts:Prefect Post:Prefect Posts:Tbl_PrefectPosts:Tbl_Prefects:Prefects"
        
        'Assign respective Form
        Set vFrm(&H0) = Frm_Dormitories
        
        'Save the current Category in order to execute it when the Application is reloaded
        VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
        
        'Check If the User has insufficient privileges to perform this Operation
        If Not ValidUserAccess(Me, &H12, &H5, False) Then GoTo Exit_MnuActivity_Click 'Branch to the specified Label
        
        'Call Function in Mdl_Stadmis to display all the Users with equal or lower levels than the current User
        Call FillListView(Lv, "SELECT [Prefect Post Name], [Prefect Post ID], [Seats], ([Hierarchical Level]) AS [Rank], [Discontinued] FROM [Tbl_PrefectPosts] ORDER BY [Hierarchical Level] ASC, [Prefect Post Name] ASC", "2", , "Prefect Post ID", "Discontinued:=YES", &HC0&, IsFrmLoadingComplete, , , &H4)
        
        GoTo Exit_MnuActivity_Click
        
    End If
    
    If Index = &HA Then
        
        'Initial variable
        lblSelectedCategory(&H1).Tag = VBA.vbNullString
        
        'Display Category Title
        lblSelectedCategory(&H0).Caption = "Prefects"
        lblSelectedCategory(&H1).Caption = "Information of all the Prefects"
        
        'Remove all Listview Rows and Columns
        Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
        
        FrmDefinitions = "Prefect:Prefects:Prefect:Prefects:Qry_Prefects::|Prefect Post:Prefect Posts:Prefect Post:Prefect Posts:Tbl_PrefectPosts:Frm_Dormitories:Prefect Post~Prefect Posts~Prefect Post~Prefect Posts~Tbl_PrefectPosts~Tbl_Prefects~Prefects"
        
        'Assign respective Form
        Set vFrm(&H0) = Frm_StudentSports
        
        'Save the current Category in order to execute it when the Application is reloaded
        VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
        
        'Check If the User has insufficient privileges to perform this Operation
        If Not ValidUserAccess(Me, &H13, &H5, False) Then GoTo Exit_MnuActivity_Click 'Branch to the specified Label
        
        'Call Function in Mdl_Stadmis to display all the Users with equal or lower levels than the current User
        Call FillListView(Lv, "SELECT [Start Date], [Prefect ID], [Student ID], ([Registered Name]) AS [Student Name], [Adm No], [Gender], ([Prefect Post Name]) AS [Post], [End Date], [Allocation Status] FROM [Qry_Prefects] WHERE YEAR([Start Date]) >= " & SoftwareSetting.Min_Year & " ORDER BY [Start Date] DESC, [Hierarchical Level] ASC, [Prefect Post Name] ASC, [Registered Name] ASC", "2;3", , "Prefect ID", "Allocation Status:<>''", &HC0&, IsFrmLoadingComplete, , , &H4)
        iPhotoSpecifications = "[Qry_Students] WHERE [Student ID] = $;3;Student Name"
        
        GoTo Exit_MnuActivity_Click
        
    End If
    
    Dim sSing$, sPlural$
    
    sPlural = VBA.Replace(VBA.Replace(MnuActivity(Index).Caption, "Student ", ""), "&", "")
    sSing = VBA.IIf(VBA.Right$(sPlural, &H3) = "ies", VBA.Replace(sPlural, "ies", "y"), VBA.Left$(sPlural, VBA.Len(sPlural) - &H1))
    
    'Initial variable
    lblSelectedCategory(&H1).Tag = VBA.vbNullString
    
    'Remove all Listview Rows and Columns
    Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
    
    'Display Category Title
    lblSelectedCategory(&H0).Caption = VBA.Replace(MnuActivity(Index).Caption, "&", "")
    lblSelectedCategory(&H1).Caption = "Information of all the " & VBA.IIf(VBA.InStr(sPlural, "Student"), VBA.Replace(sPlural, "Student ", "") & " assigned to Students", sPlural) & " in the School"
    
    'Save the current Category in order to execute it when the Application is reloaded
    VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
    
    'If allocated to Students then...
    If VBA.InStr(MnuActivity(Index).Caption, "Student ") Then
        
        FrmDefinitions = "Student " & sSing & ":Student " & sPlural & ":Student " & sSing & ":Student " & sPlural & ":Qry_Student" & sPlural & "::|" & sSing & ":" & sPlural & ":" & sSing & ":" & sPlural & ":Tbl_" & sPlural & ":Frm_Dormitories:" & sSing & "~" & sPlural & "~" & sSing & "~" & sPlural & "~Tbl_" & sPlural & "~Tbl_Student" & sPlural & "~Students"
        
        'Assign respective Form
        Set vFrm(&H0) = Frm_StudentSports
        
        'Check If the User has insufficient privileges to perform this Operation
        If Not ValidUserAccess(Me, VBA.Switch(sSing = "Sport", &HD, sSing = "Club", &HF, sSing = "Society", &H11), &H5, False) Then GoTo Exit_MnuActivity_Click 'Branch to the specified Label
        
        sSing = "Student " & sSing: sPlural = "Student " & sPlural
        
    Else
        
        FrmDefinitions = sSing & ":" & sPlural & ":" & sSing & ":" & sPlural & ":Tbl_" & sPlural & ":Tbl_Student" & sPlural & ":Students"
        
        'Assign respective Form
        Set vFrm(&H0) = Frm_FamilyRelationships
        
        'Check If the User has insufficient privileges to perform this Operation
        If Not ValidUserAccess(Me, VBA.Switch(sSing = "Sport", &HB, sSing = "Club", &HD, sSing = "Society", &HF), &H5, False) Then GoTo Exit_MnuActivity_Click  'Branch to the specified Label
        
    End If
    
    'Call Function in Mdl_Stadmis to display all the Users with equal or lower levels than the current User
    Call FillListView(Lv, "SELECT " & VBA.IIf(VBA.InStr(sPlural, "Student"), "[Start Date], [" & sSing & " ID], [Student ID], [Registered Name], [" & VBA.Replace(sSing, "Student ", "") & " Name], [End Date], [Allocation Status]", "[" & VBA.Replace(sSing, "Student ", "") & " Name], [" & sSing & " ID], [Discontinued]") & " FROM [" & VBA.IIf(VBA.InStr(sPlural, "Student"), "Qry_", "Tbl_") & VBA.Replace(sPlural, " ", "") & "]" & VBA.IIf(VBA.InStr(sPlural, "Student"), "WHERE YEAR([Start Date]) >= " & SoftwareSetting.Min_Year, VBA.vbNullString) & " ORDER BY [Hierarchical Level] ASC, " & VBA.IIf(VBA.InStr(sPlural, "Student"), "[Student Name] ASC, ", "") & "[" & VBA.Replace(sSing, "Student ", "") & " Name]", "2" & VBA.IIf(VBA.InStr(sPlural, "Student"), ";3", ""), , sSing & " ID", VBA.IIf(VBA.InStr(sPlural, "Student"), "Allocation Status:<>''", "Discontinued:=YES"), &HC0&, IsFrmLoadingComplete, , , &H4)
    If VBA.InStr(sPlural, "Student") <> &H0 Then iPhotoSpecifications = "[Qry_Students] WHERE [Student ID] = $;3;Student Name"
    
Exit_MnuActivity_Click:
    
    'If Application Log then...
    If Index = &H8 Then Unload Frm_PleaseWait
    
    Lv.Visible = True: Lv.Refresh
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_MnuActivity_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Displaying Activity - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuActivity_Click
    
End Sub

Private Sub MnuAdministration_Click(Index As Integer)
On Local Error GoTo Handle_MnuConfiguration_Click_Error
    
    'If the Menu is a parent menu then Quit this Procedure
    If Index = &H0 Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Save the current Category in order to execute it when the Application is reloaded
    ItemCategoryNo = &H4: ItemIndex = Index
    
    Select Case Index
        
        Case &H0:
        Case &H1: Call lblMenu_Click(&H1)
        Case &H2: Call lblMenu_Click(&H4)
        Case &H4: Call lblMenu_Click(&H2)
        Case &H5: Call lblMenu_Click(&H5)
        Case &H7: Call lblMenu_Click(&H3)
        Case &H8:
        
            iSetNo = &HB
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Student's Dormitories"
            lblSelectedCategory(&H1).Caption = "Information of every Student's allocated to Dormitories in the School"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_Students
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Dormitories
            Call FillListView(Lv, "SELECT YEAR([Date Assigned]) AS [Year], [Student ID], [Student Dormitory ID], ([Dormitory Name]) AS [Dormitory], [Adm No], ([Registered Name]) AS [Student Name], [Student Dormitory Status], [Dormitory Discontinued] FROM [Qry_StudentDormitories] ORDER BY [Date Assigned] DESC, [Dormitory Name] ASC, [Registered Name] ASC", "2;3", , "Student Dormitory ID", "Student Dormitory Status:<>''|Dormitory Discontinued:=YES", &HC0& & "|" & &HC0&, IsFrmLoadingComplete, , , &H2)
            iPhotoSpecifications = "[Qry_Students] WHERE [Student ID] = $;2;[Student ID];Registered Name"
            
        Case &HA: 'Relationships
            
            iSetNo = &H5
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Relationship"
            lblSelectedCategory(&H1).Caption = "Information of every Relationship between Students and Guardians in the School"
            
            FrmDefinitions = "Relationship:Relationships:Relationship:Relationships:Tbl_Relationships:Tbl_StudentGuardians:Students"
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_FamilyRelationships
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, iSetNo, &H5, False) Then Exit Sub
            
            'Call Function in Mdl_Stadmis to display all the Guardians
            Call FillListView(Lv, "SELECT [Relationship Name], [Relationship ID], [Discontinued] FROM [Tbl_Relationships] ORDER BY [Hierarchical Level] ASC", "2", , "Relationship ID", "Discontinued:=YES", &HC0&, IsFrmLoadingComplete, , , &H2)
            
        Case Else: 'If not defined then...
            
            'disassociate object variable from any actual object
            Set vFrm(&H0) = Nothing
            
    End Select 'Close SELECT..CASE block statement
    
Exit_MnuConfiguration_Click:
    
    Lv.Visible = True: Lv.Refresh
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_MnuConfiguration_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Administration Menu Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuConfiguration_Click
    
End Sub

Private Sub MnuConfiguration_Click(Index As Integer)
On Local Error GoTo Handle_MnuConfiguration_Click_Error
    
    'If the Menu is a parent menu then Quit this Procedure
    If Index = &H3 Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'Determines whether double-click on Listview loads Records under the selected item or not
    vBuffer(&H1) = VBA.vbNullString  'Initialize variable
    vBuffer(&H2) = VBA.vbNullString  'Initialize variable
    
    'Save the current Category in order to execute it when the Application is reloaded
    ItemCategoryNo = &H2: ItemIndex = Index
    
    'If the selected list item...
    Select Case Index
        
        Case &H0: 'Is Software Users...
            
            Call lblMenu_Click(&H6)
            
        Case &H4: 'Software Settings...
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H2, &HF, False) Then GoTo Exit_MnuConfiguration_Click 'Branch to the specified Label
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Display specified Form
            CenterForm Frm_Settings, Me
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
        Case &H5: 'Is Login Report...
            
            'Initial variable
            lblSelectedCategory(&H1).Tag = VBA.vbNullString
            
            'Remove all Listview Rows and Columns
            Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
            
            Lv.Visible = False
            
            'Disassociate object variable from actual object
            Set vFrm(&H0) = Nothing
            
            'Initialize variable
            FrmDefinitions = VBA.vbNullString
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Login Records"
            lblSelectedCategory(&H1).Caption = "View Logon sessions of users into the Application"
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H2, &HD, False) Then GoTo Exit_MnuConfiguration_Click 'Branch to the specified Label
            
            'Call Function in Mdl_SchooFee to display all the Login Records
            Call FillListView(Lv, "SELECT [Login No], [Login Date], [User ID], [User Name], ([Registered Name]) AS [Full Name], [Hierarchy], [Account Name], [Device Serial], [Device Name], [Device Account], [Logout Date], [Successful Logout], [Duration In Mins] FROM [Qry_Login] WHERE [Hierarchy] >= " & User.Hierarchy & " AND [Login No] <> " & User.Login_ID & " ORDER BY [Login Date] DESC, [Device Name] ASC, [Hierarchy] ASC", , , , "Successful Logout:=FALSE", &HC0&, IsFrmLoadingComplete)
            iPhotoSpecifications = VBA.vbNullString
            
        Case &H6: 'Is Application Event Log...
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H2, &HE, False) Then GoTo Exit_MnuConfiguration_Click 'Branch to the specified Label
            
            Frm_PleaseWait.ImgProgressBar.Width = &H0
            Frm_PleaseWait.ImgProgressBar.Left = Frm_PleaseWait.lblProgressBar.Left
            Frm_PleaseWait.ImgProgressBar.Visible = True
            
            'Show the 'Please Wait' Form to the User
            CenterForm Frm_PleaseWait, Me, False
            
            'Disassociate object variable from actual object
            Set vFrm(&H0) = Nothing
            
            'Initialize variables
            FrmDefinitions = VBA.vbNullString
            lblSelectedCategory(&H1).Tag = VBA.vbNullString
                        
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "Application Event Log"
            lblSelectedCategory(&H1).Caption = "Monitors events recorded in the Software's Application log"
            
            iPhotoSpecifications = VBA.vbNullString
            Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
            Lv.Visible = False
            
            Dim strTxtWidth() As String
            
            ReDim strTxtWidth(&HE) As String
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "No"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Type"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Date"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Time"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Category"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Source"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Error No"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Description"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Option"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "User ID"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "User Name"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Device Account Name"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Device Serial"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Device Name"
            
            If VBA.Len(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) < 80 Then
                If VBA.Val(strTxtWidth(Lv.ColumnHeaders.Count - &H1)) < VBA.Val(Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text)) Then strTxtWidth(Lv.ColumnHeaders.Count - &H1) = Me.TextWidth(Lv.ColumnHeaders(Lv.ColumnHeaders.Count).Text) + 150
            Else
                strTxtWidth(Lv.ColumnHeaders.Count - &H1) = 2000
            End If 'Close respective IF..THEN block statement
            
            iLvwItemDataType = "N|T|D|D|T|T|T|T|N|T|T|T|N|T" 'Customize data types for sorting
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'If no Application Event Log does not exist then...
            If Not vFso.FileExists(def_LogFileLocation & "\" & App.Title & " Event Log.txt") Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                If IsFrmLoadingComplete Then vMsgBox "There is no existing Application Event Log", vbInformation, App.Title & " : Application Event Log", Me
                
                GoTo ResizeColumns 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
            Dim Fl
            Dim FlData$
            Dim iCounter&, iCnt#
            Dim iArray() As String
            Dim iArrayTmp() As String
            
            Set Fl = vFso.OpenTextFile(def_LogFileLocation & "\" & App.Title & " Event Log.txt", ForReading)
            
            If Fl.AtEndOfStream Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count: GoTo Exit_MnuConfiguration_Click
            
            FlData = Fl.ReadAll
            iArray = VBA.Split(FlData, VBA.vbCrLf)
            
            iCounter = UBound(iArray) + &H1 + Lv.ColumnHeaders.Count
            
            For vIndex(&H0) = &H0 To UBound(iArray) Step &H1
                
                iArrayTmp = VBA.Split(iArray(vIndex(&H0)), "|")
                
                vIndex(&H1) = VBA.Switch(iArrayTmp(&HB) = "Critical", &H9, iArrayTmp(&HB) = "Question", &HA, iArrayTmp(&HB) = "Warning", &HB, iArrayTmp(&HB) = "Information", &HC)
                
                'No
                Lv.ListItems.Add Lv.ListItems.Count + &H1, , Lv.ListItems.Count + &H1, vIndex(&H1), vIndex(&H1)
                
                If VBA.Len(vIndex(&H1)) < 110 Then
                    If VBA.Val(strTxtWidth(&H0)) < VBA.Val(Me.TextWidth(vIndex(&H1))) Then strTxtWidth(&H0) = Me.TextWidth(vIndex(&H1)) + 0
                Else
                    strTxtWidth(&H0) = 2000
                End If 'Close respective IF..THEN block statement
                
                'Type
                If UBound(iArrayTmp) >= &HB Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&HB)
                    
                    If VBA.Len(iArrayTmp(&HB)) < 110 Then
                        If VBA.Val(strTxtWidth(&H1)) < VBA.Val(Me.TextWidth(iArrayTmp(&HB))) Then strTxtWidth(&H1) = Me.TextWidth(iArrayTmp(&HB)) + 0
                    Else
                        strTxtWidth(&H1) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Date
                If UBound(iArrayTmp) >= &H0 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H0)
                    
                    If VBA.Len(iArrayTmp(&H2)) < 110 Then
                        If VBA.Val(strTxtWidth(&H0)) < VBA.Val(Me.TextWidth(iArrayTmp(&H0))) Then strTxtWidth(&H2) = Me.TextWidth(iArrayTmp(&H0)) + 0
                    Else
                        strTxtWidth(&H2) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Time
                If UBound(iArrayTmp) >= &H1 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H1)
                    
                    If VBA.Len(iArrayTmp(&H1)) < 110 Then
                        If VBA.Val(strTxtWidth(&H3)) < VBA.Val(Me.TextWidth(iArrayTmp(&H1))) Then strTxtWidth(&H3) = Me.TextWidth(iArrayTmp(&H1)) + 0
                    Else
                        strTxtWidth(&H3) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Category
                If UBound(iArrayTmp) >= &H8 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H8)
                    
                    If VBA.Len(iArrayTmp(&H8)) < 110 Then
                        If VBA.Val(strTxtWidth(&H4)) < VBA.Val(Me.TextWidth(iArrayTmp(&H8))) Then strTxtWidth(&H4) = Me.TextWidth(iArrayTmp(&H8)) + 0
                    Else
                        strTxtWidth(&H4) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Source
                If UBound(iArrayTmp) >= &H9 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H9)
                    
                    If VBA.Len(iArrayTmp(&H9)) < 110 Then
                        If VBA.Val(strTxtWidth(&H5)) < VBA.Val(Me.TextWidth(iArrayTmp(&H9))) Then strTxtWidth(&H5) = Me.TextWidth(iArrayTmp(&H9)) + 0
                    Else
                        strTxtWidth(&H5) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Error No
                If UBound(iArrayTmp) >= &HA Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&HA)
                    
                    If VBA.Len(iArrayTmp(&HA)) < 110 Then
                        If VBA.Val(strTxtWidth(&H6)) < VBA.Val(Me.TextWidth(iArrayTmp(&HA))) Then strTxtWidth(&H6) = Me.TextWidth(iArrayTmp(&HA)) + 0
                    Else
                        strTxtWidth(&H6) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
                'Description
                If UBound(iArrayTmp) >= &HC Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&HC)
                    
                    If VBA.Len(iArrayTmp(&HC)) < 110 Then
                        If VBA.Val(strTxtWidth(&H7)) < VBA.Val(Me.TextWidth(iArrayTmp(&HC))) Then strTxtWidth(&H7) = Me.TextWidth(iArrayTmp(&HC)) + 0
                    Else
                        strTxtWidth(&H7) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Option
                If UBound(iArrayTmp) >= &HD Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&HD)
                    
                    If VBA.Len(iArrayTmp(&HD)) < 110 Then
                        If VBA.Val(strTxtWidth(&H8)) < VBA.Val(Me.TextWidth(iArrayTmp(&HD))) Then strTxtWidth(&H8) = Me.TextWidth(iArrayTmp(&HD)) + 0
                    Else
                        strTxtWidth(&H8) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'User ID
                If UBound(iArrayTmp) >= &H6 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H6)
                    
                    If VBA.Len(iArrayTmp(&H6)) < 110 Then
                        If VBA.Val(strTxtWidth(&H9)) < VBA.Val(Me.TextWidth(iArrayTmp(&H6))) Then strTxtWidth(&H9) = Me.TextWidth(iArrayTmp(&H6)) + 0
                    Else
                        strTxtWidth(&H9) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'User Name
                If UBound(iArrayTmp) >= &H7 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H7)
                    
                    If VBA.Len(iArrayTmp(&H7)) < 110 Then
                        If VBA.Val(strTxtWidth(&HA)) < VBA.Val(Me.TextWidth(iArrayTmp(&H7))) Then strTxtWidth(&HA) = Me.TextWidth(iArrayTmp(&H7)) + 0
                    Else
                        strTxtWidth(&HA) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Device Account Name
                If UBound(iArrayTmp) >= &H4 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H4)
                    
                    If VBA.Len(iArrayTmp(&H4)) < 110 Then
                        If VBA.Val(strTxtWidth(&HC)) < VBA.Val(Me.TextWidth(iArrayTmp(&H4))) Then strTxtWidth(&HC) = Me.TextWidth(iArrayTmp(&H4)) + 0
                    Else
                        strTxtWidth(&HC) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Device Serial
                If UBound(iArrayTmp) >= &H2 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H2)
                    
                    If VBA.Len(iArrayTmp(&H2)) < 110 Then
                        If VBA.Val(strTxtWidth(&HD)) < VBA.Val(Me.TextWidth(iArrayTmp(&H2))) Then strTxtWidth(&HD) = Me.TextWidth(iArrayTmp(&H2)) + 0
                    Else
                        strTxtWidth(&HD) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                'Device Name
                If UBound(iArrayTmp) >= &H3 Then
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , iArrayTmp(&H3)
                    
                    If VBA.Len(iArrayTmp(&H3)) < 110 Then
                        If VBA.Val(strTxtWidth(&HE)) < VBA.Val(Me.TextWidth(iArrayTmp(&H3))) Then strTxtWidth(&HE) = Me.TextWidth(iArrayTmp(&H3)) + 0
                    Else
                        strTxtWidth(&HE) = 2000
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add Lv.ListItems(Lv.ListItems.Count).ListSubItems.Count + &H1, , VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
                Frm_PleaseWait.ImgProgressBar.Width = iCnt
                Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.Val(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
                
            Next vIndex(&H0)
            
ResizeColumns:
            
            Dim vCol&
            
            'For each column in the specified Listview...
            For vCol = &H1 To Lv.ColumnHeaders.Count Step &H1
                
                If VBA.Val(strTxtWidth(vCol - &H1)) = &H0 And vCol = &H1 And (Not Nothing Is Lv.SmallIcons Or Lv.Checkboxes) Then
                    
                    Lv.ColumnHeaders(vCol).Width = VBA.IIf(Not Nothing Is Lv.SmallIcons, 255, &H0) + VBA.IIf(Lv.Checkboxes, 255, &H0)
                    Lv.ColumnHeaders(vCol).Text = "    " & Lv.ColumnHeaders(vCol).Text
                    
                Else
                    
                    'Resize the Column Width to fit the longest text in it
                    Lv.ColumnHeaders(vCol).Width = VBA.IIf(VBA.Val(strTxtWidth(vCol - &H1)) = &H0, &H0, VBA.Val(strTxtWidth(vCol - &H1)) + 50 + VBA.IIf(Not Nothing Is Lv.SmallIcons, 255, &H0) + VBA.IIf(Lv.Checkboxes, 255, &H0))
                    
                End If 'Close respective IF..THEN block statement
                
                iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
                Frm_PleaseWait.ImgProgressBar.Width = iCnt
                Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.Val(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
                
            Next vCol 'Increment the X variable by the value in the Step option
            
            lblRecords.Caption = "Total Records: " & Lv.ListItems.Count
            
        Case Else: 'If not defined then...
            
            'disassociate object variable from any actual object
            Set vFrm(&H0) = Nothing
            
    End Select 'Close SELECT..CASE block statement
    
Exit_MnuConfiguration_Click:
    
    'If Application Log then...
    If Index = &H6 Then Unload Frm_PleaseWait
    
    Lv.Visible = True: Lv.Refresh
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_MnuConfiguration_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Configuration Menu Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuConfiguration_Click
    
End Sub

Private Sub MnuConfigurationTool_Click(Index As Integer)
On Local Error GoTo Handle_MnuConfigurationTool_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'Save the current Category in order to execute it when the Application is reloaded
    ItemCategoryNo = &H3: ItemIndex = Index
    
    'If the selected list item...
    Select Case VBA.Replace(MnuConfigurationTool(Index).Caption, "&", VBA.vbNullString)
        
        Case "Backup Utility": 'Back up Software data...
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H4, &H1, , True) Then Exit Sub
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Display specified Form
            CenterForm Frm_Backup, Me
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
        Case "Software Settings": 'Back up Software data...
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H2, &HE, , True) Then Exit Sub
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Display specified Form
            CenterForm Frm_Settings, Me
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
        Case "School Info": 'School Info...
            
            'Initial variable
            lblSelectedCategory(&H1).Tag = VBA.vbNullString
            
            'Remove all Listview Rows and Columns
            Lv.ListItems.Clear: Lv.ColumnHeaders.Clear
            
            'Assign respective Form
            Set vFrm(&H0) = Frm_School
            
            'Display Category Title
            lblSelectedCategory(&H0).Caption = "School"
            lblSelectedCategory(&H1).Caption = "School owning this Application"
            
            'Save the current Category in order to execute it when the Application is reloaded
            VBA.SaveSetting App.Title, "Main Form", User.User_ID & " : Last Selected Category", ItemCategoryNo & ":" & ItemIndex
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H1, &H5, False, True) Then GoTo Exit_MnuConfigurationTool_Click  'Branch to the specified Label
            
            'Call Function in Mdl_SchooFee to display all the Users with equal or lower levels than the current User
            Call FillListView(Lv, "SELECT [School Name], [School ID], [Abbreviation], [Bank Branch Name], [Bank Acc Name], [Bank Acc No], [Discontinued] FROM [Tbl_School] ORDER BY [School Name] ASC", "2", , "School ID:[School ID] = 2", "Discontinued:=YES", &HC0&, True)
            iPhotoSpecifications = "[Tbl_School] WHERE [School ID] = $;2;[School ID]"
            
        Case Else: 'If not defined then...
            
    End Select 'Close SELECT..CASE block statement
    
Exit_MnuConfigurationTool_Click:
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_MnuConfigurationTool_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Configuration Tools Menu Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuConfigurationTool_Click
    
End Sub

Private Sub MnuFile_Click(Index As Integer)
On Local Error GoTo Handle_MnuFile_Click_Error
    
    'If the Menu is a parent menu then Quit this Procedure
    If Index = &H2 Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'If the selected list item...
    Select Case Index
        
        Case &H0: 'Change current User's password...
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H3, &H6, False) Then Exit Sub
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Display specified Form
            Call CenterForm(Frm_ChangePassword, Me)
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
        Case &H4, &H5: 'Log Off or Exit Application...
            
            vQuitState = VBA.IIf(Index = &H4, &H1, &H0)
            Unload Me 'Unload this Form from the memory
            
        Case Else: 'If not defined then...
            
    End Select 'Close SELECT..CASE block statement
    
Exit_MnuFile_Click:
    
    'Initialize variables
    vBuffer(&H3) = VBA.vbNullString: vQuitState = &H0
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_MnuFile_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Button Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuFile_Click
    
End Sub

Private Function ShowBallonMsg() As Boolean
On Local Error GoTo Handle_ShowBallonMsg_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the Software has not been registered then...
    If Not vRegistered And vQuitState = &H0 And IsFrmLoadingComplete Then
        
        'Display trial message balloon on the System tray icon
        Set SystemTray = New clsSystemTray
        
        With SystemTray
            
            .Icon = Icon.Handle
            .Menu = Me.hWnd
            .Parent = Me.hWnd
            .TipText = App.Title & " : " & App.FileDescription
            
            If .IconStatus = ICON_UNLOADED Then Call .AddIcon
            Call SystemTray.ShowBalloon(App.Title & " : Trial Version", "The System will expire on " & LblLicenseTo(&H3).Caption & ". Remaining " & VBA.DateDiff("d", VBA.Now, VBA.Right$(LblLicenseTo(&H3).Caption, VBA.Len(LblLicenseTo(&H3).Caption) - &H4)) & " days", NIIF_WARNING, 7000, True)
            
        End With
        
    End If
    
    ShowBallonMsg = True
    
Exit_ShowBallonMsg:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ShowBallonMsg_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Showing Tray Balloon - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ShowBallonMsg
    
End Function

Private Sub MnuHelp_Click(Index As Integer)
On Local Error GoTo Handle_MnuHelp_Click_Error
    
    Dim MousePointerState%
    Dim HelpFilePath$
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
        
    Select Case Index
        
        Case &H0: 'Help
            
            'Assign default Program's Help File location
            HelpFilePath = App.Path & "\Help\" & App.Title & " User Guide.pub"
            
            'If the Program's Help File is missing then...
            If Not vFso.FileExists(HelpFilePath) Then
                
                'Warn User
                vMsgBox "Missing Help file. Please reinstall this Program.", vbExclamation, App.Title & " : Missing File", Me
                GoTo Exit_MnuHelp_Click 'Resume execution at the specified Label
                
            End If 'Close respective IF..THEN block statement
            
            Call OpenFile(Me, HelpFilePath) 'Open the Help File and display to the User
            
        Case &H1: 'About Application
            
            Frm_Splash.HinderTransparency = True
            CenterForm Frm_Splash, Me
            
        Case &H2: 'Register
            
            Me.Enabled = False
            
            vRegistering = True: vRegistered = True
            Frm_SoftwarePatent.Show , Me
            
            Do While vRegistering
                VBA.DoEvents: VBA.DoEvents
            Loop
            
            LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < 61 And Not vRegistered, True, False)
            
            MnuHelp(&H2).Visible = LblLicenseTo(&H0).Visible
            
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
            
            Me.Enabled = True
            
    End Select 'Close SELECT..CASE block statement
    
Exit_MnuHelp_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_MnuHelp_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Help Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuHelp_Click
    
End Sub

Private Sub MnuStudents_Click(Index As Integer)
On Local Error GoTo Handle_MnuConfiguration_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Save the current Category in order to execute it when the Application is reloaded
    ItemCategoryNo = &H5: ItemIndex = Index
    
    Select Case Index
        
        Case &H0: 'Current year Students
            Call lblMenu_Click(&H0)
            
        Case Else: 'If not defined then...
            
            'disassociate object variable from any actual object
            Set vFrm(&H0) = Nothing
            
    End Select 'Close SELECT..CASE block statement
    
Exit_MnuConfiguration_Click:
    
    Lv.Visible = True: Lv.Refresh
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_MnuConfiguration_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Students Menu Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuConfiguration_Click
    
End Sub

Private Sub ShpBttnMenu_Click(Index As Integer)
On Local Error GoTo Handle_ShpBttnMenu_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'If not NEW, EDIT or DELETE command then Initialize variable
    If Index > &H2 Then FrmDefinitions = VBA.vbNullString
    
    'If the selected list item...
    Select Case Index
        
        Case &H0 To &H2: 'Is NEW, EDIT or DELETE then...
            
            'If an Input Form has been specified then...
            If Not Nothing Is vFrm(&H0) Then
                
                Fra_Photo.Visible = False 'Hide the Photo Frame
                
                'If Edit or Delete command has been selected then...
                If Index > &H0 Then
                    
                    'If the User has not selected any Record then...
                    If Nothing Is Lv.SelectedItem Then
                        
                        'Indicate that a process or operation is complete.
                        Screen.MousePointer = vbDefault
                        
                        'Warn User
                        vMsgBox "Please select a Record before you " & VBA.IIf(Index = &H1, "Edit it.", "Delete it."), vbExclamation, App.Title & " : Invalid Command", Me
                        
                        'Indicate that a process or operation is in progress.
                        Screen.MousePointer = vbHourglass
                        
                        'Branch to the specified Label
                        GoTo Exit_ShpBttnMenu_Click
                        
                    End If 'Close respective IF..THEN block statement
                    
                    'Assign the selected item's Primary Key
                    vEditRecordID = Lv.SelectedItem.ListSubItems(&H1).Text & VBA.IIf(Index = &H1, "|", "||")
                    
                End If 'Close respective IF..THEN block statement
                
                vFrm(&H0).vRecordID = &H0
                
                'Assign Form requirements if any
                vFrm(&H0).FrmDefinitions = FrmDefinitions
                
                'Assign this Form's icon to the one to be displayed
                vFrm(&H0).Icon = Me.Icon
                
                iDisplayingForm = True
                
                'Show it to the User
                vFrm(&H0).Show vbModal, Me
                
                iDisplayingForm = False
                
                vEditRecordID = VBA.vbNullString 'Initialize variable
                
                If vDatabaseAltered Then vDatabaseAltered = False: Call ShpBttnMenu_Click(&H4)
                
                Lv.SetFocus 'Move Focus to the specified control
                
                'Branch to the specified Label
                GoTo After_ShpBttnMenu_Click
                
            End If 'Close respective IF..THEN block statement
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case &H3: 'Is Search then...
            
            'Call Function in Mdl_Stadmis Module to perform Search
            Call SearchLv(Lv, , iSelColPos)
            
            'Display the total no of Records retrieved after Search
            lblRecords.Caption = "Total Records: " & Lv.ListItems.Count
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case &H4: 'Is Refresh then...
            
            If Not Nothing Is Lv.SelectedItem Then LstLastItem = Lv.SelectedItem.Index
            
            vWait = True
            
            vQuitState = &H3
            'Re-open the Form
            Unload Me: Me.Show
            
            'Wait for the Main Form to load completely
            Do While vWait
                VBA.DoEvents 'Yield execution so that the operating system can process other events
            Loop
            
            vQuitState = &H0
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case &H5: 'Is Export To Excel then...
            
            vExportFreezeCol = &HA
             
            'Call Procedure in Mdl_Stadmis Module to export the Listview data to Excel
            Call ExportLvToExcel(Lv, VBA.UCase$(lblSelectedCategory(&H0).Caption) & VBA.IIf(lblSelectedCategory(&H0).Caption = lblSelectedCategory(&H1).Caption, VBA.vbNullString, " - " & lblSelectedCategory(&H1).Caption))
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case &H6: 'Is Export To Word then...
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case &H7: 'Is Reports...
            
            Select Case lblSelectedCategory(&H0).Caption
                
                Case Else 'If the selected Data has no defined Report then...
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Inform User
                    vMsgBox "There is no defined Report for the displayed data. Please export to Excel to customize the data.", vbInformation, App.Title & " : No Defined Report", Me
                    
                    'Indicate that a process or operation is in progress.
                    Screen.MousePointer = vbHourglass
                    
            End Select 'Close respective SELECT..CASE block statement
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case &H8: 'Is Calculator...
            
            'Display the Computer's Calculator program to the User
            VBA.Shell "Calc", vbNormalFocus
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
        Case Else: 'If not defined then...
            
            'Branch to the specified Label
            GoTo Exit_ShpBttnMenu_Click
            
    End Select 'Close SELECT..CASE block statement
    
Exit_ShpBttnMenu_Click:
    
    vBuffer(&H0) = VBA.vbNullString 'Initialize variable
    
    If Lv.ListItems.Count > &H0 Then lblRecords.Caption = "Total Records: " & Lv.ListItems.Count - VBA.IIf(Lv.ListItems(Lv.ListItems.Count).Tag = "Function", &H1, &H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
After_ShpBttnMenu_Click:
    
    'If changes have been made in the database then execute codes under Form_Activate event
    If vDatabaseAltered Then IsFrmLoading = True: Call Form_Activate
    
    'Denote that the database table has not been altered
    vDatabaseAltered = False
    
    'Branch to the specified Label
    GoTo Exit_ShpBttnMenu_Click
    
Handle_ShpBttnMenu_Click_Error:
    
    'If the Form doesn't need requirements the execute the next line of code
    If Err.Number = 438 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Button Click Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnMenu_Click
    
End Sub

Private Sub SystemTray_Click(Button As Integer)
    
    If SystemTray.BalloonLastMessage = VBA.vbNullString Then Exit Sub
    
    vArrayList = VBA.Split(SystemTray.BalloonLastMessage, "|")
    
    'Warn User of the system's expiry status
    Call SystemTray.ShowBalloon(vArrayList(&H0), vArrayList(&H1), NIIF_WARNING, 10000, True)
    
End Sub

Private Sub SystemTray_DblClick(Button As Integer)
    
    If Button = vbRightButton Then Exit Sub
    
    Static iShowFrm As Boolean
    
    iShowFrm = Not iShowFrm
    Me.Visible = iShowFrm: Me.WindowState = VBA.IIf(iShowFrm, iWindoState, &H1)
    
End Sub

Private Sub TimerDateTime_Timer()
On Local Error GoTo Handle_TimerDateTime_Timer_Error
    
    LblSecond.Caption = VBA.Format$(VBA.Time$, "ss")
    
    LblLicenseTo(&H3).ForeColor = VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) + &H1 < 31, tTheme.tWarningForeColor, tTheme.tEntryColor)
    LblLicenseTo(&H3).Caption = VBA.Format$(SoftwareSetting.Licences.Expiry_Date, "ddd dd MMM yyyy hh:nn:ss AMPM") & VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < &HB, " - Remaining " & VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) - &H1 & " day" & VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) - &H1 = &H1, "", "s") & ". This Software will automatically Shut down after expiring.", VBA.vbNullString)
    
    If Not IsFrmLoadingComplete Or Not vStartupComplete Or vRegistering Then Exit Sub
    
    'If the User Logon Time Limit has been specified then...
    If User.Parental_Control_ON Then
        
        If Not VBA.IsNull(User.Start_Time) And Not VBA.IsNull(User.End_Time) Then
            
            If VBA.DateDiff("n", User.Start_Time, VBA.FormatDateTime(VBA.Now, vbShortTime)) >= &H0 And VBA.DateDiff("n", User.End_Time, VBA.FormatDateTime(VBA.Now, vbShortTime)) >= &H0 Then
                
                Dim Frm As Form
                Static UnloadError As Boolean
                
StartUnload:
                
                'Unload an application properly ensuring restoration of resources
                For Each Frm In VB.Forms
                    
                    'Close all other open Forms in this Application without alerts, apart from this Main one
                    If Frm.Name <> Me.Name Then vSilentClosure = True: Unload Frm
                    
                Next Frm 'Move to the next open Form
                
                If UnloadError Then UnloadError = False: GoTo StartUnload
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "Your account operational time frame {" & User.Start_Time & " to " & User.End_Time & "} has elapsed. Exiting Application", vbExclamation, App.Title & " : Parental Time Limit", Me
                
                Unload Me
                
                vSilentClosure = False
                
                Exit Sub 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    'If the Software's Licence has expired then...
    If (VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < &H1) And (SoftwareSetting.Licences.Expiry_Date <> VBA.DateSerial(1986, 3, 9)) Then
        
1:
        
        Select Case VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date)
            
            Case &H0: vBuffer(&H0) = "has expired today."
            Case &H1: vBuffer(&H0) = "expired yesterday."
            Case Else: vBuffer(&H0) = "expired on " & VBA.Format$(SoftwareSetting.Licences.Expiry_Date, "ddd dd MMM yyyy") & "."
            
        End Select
        
        vBuffer(&H0) = vBuffer(&H0) & " Please provide the following to the Software Administrator for renewal:-" & VBA.vbCrLf & _
                    "1. Software Name : " & App.Title & VBA.vbCrLf & _
                    "2. Device Serial : " & User.Device_Serial_No & VBA.vbCrLf & _
                    "3. Licence Key   : " & SoftwareSetting.Licences.Key & VBA.vbCrLf & _
                    "4. " & LblLicenseTo(&H6).Caption & "  : " & SoftwareSetting.Licences.Max_Users
                    
        'Warn User to renew the Licence
        If vMsgBox("The " & App.Title & " Software's Licence period " & vBuffer(&H0), vbCritical + vbCustomButtons + vbDefaultButton2, App.Title & " : Licence Expired", Me, , , , , , "&Register|&Exit Application") = &H1 Then
            
            Me.Enabled = False
            
            vRegistering = True: vRegistered = True
            Frm_SoftwarePatent.Show , Me
            
            Do While vRegistering
                VBA.DoEvents: VBA.DoEvents
            Loop
            
            LblLicenseTo(&H0).Visible = VBA.IIf(VBA.DateDiff("d", VBA.Date, SoftwareSetting.Licences.Expiry_Date) < 61 And Not vRegistered, True, False)
            
            MnuHelp(&H2).Visible = LblLicenseTo(&H0).Visible
            
            LblLicenseTo(&H1).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H2).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H3).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H4).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H5).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H6).Visible = LblLicenseTo(&H0).Visible
            LblLicenseTo(&H7).Visible = LblLicenseTo(&H0).Visible
            
            LblLicenseTo(&H1).Caption = SoftwareSetting.Licences.License_Code
            LblLicenseTo(&H5).Caption = SoftwareSetting.Licences.Key
            LblLicenseTo(&H7).Caption = SoftwareSetting.Licences.Max_Users
            
            Me.Enabled = True
            
            'If the software has not expired then quit this procedure
            If SoftwareSetting.Licences.Expiry_Date > VBA.Date Then Exit Sub
            
        End If
        
        'Warn User
        vMsgBox "Software's Licence period not successfully verified. Exiting Application in £ seconds..", vbExclamation, App.Title & " : Software Copyright", Me, , , &HA
        
        'Automatically Close the Software
        vSilentClosure = True: Call MnuFile_Click(&H5)
        
    End If 'Close respective IF..THEN block statement
    
Exit_TimerDateTime_Timer:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Procedure
    
Handle_TimerDateTime_Timer_Error:
    
    'If the Form doesn't need requirements the execute the next line of code
    If Err.Number = 402 Then UnloadError = True: Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Date Time Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_TimerDateTime_Timer
    
End Sub

'------------------------------------------------------------------------------------------
'The Following Codes enable the User to Drag the displayed Photo across the Listview Item

Private Sub DragMe(CurrX!)
    
    PhotoDragged = False
    
    If DragPhoto = True Then 'Check if it is in drag mode
    
        'If the Photo is dragged past the Listview left position
        If Fra_Photo.Left < Fra_Main.Left + Lv.Left + VBA.IIf(Not Nothing Is Lv.SmallIcons, 300, &H0) Then
        
            'Set the Photo to the left position of the sheet
            Fra_Photo.Left = Fra_Main.Left + Lv.Left + VBA.IIf(Not Nothing Is Lv.SmallIcons, 300, &H0)
            DragPhoto = False 'Disable drag mode
            Exit Sub 'Quit this procedure
            
        Else 'If the Photo is dragged within the Listview then...
        
            'If the Photo is dragged past the Listview right position
            If Fra_Photo.Left > (Fra_Main.Left + Lv.Left + Lv.Width) - Fra_Photo.Width Then
                
                'Set the Photo to the right position of the Listview
                Fra_Photo.Left = (Fra_Main.Left + Lv.Width + Lv.Left) - Fra_Photo.Width
                DragPhoto = False 'Disable drag mode
                Exit Sub 'Quit this procedure
                
            Else
                
                'If the Photo frame is dragged within the Listview
                'Move the Photo to the mouse pointer position
                Fra_Photo.Move (Fra_Photo.Left + CurrX - Xaxis)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        PhotoDragged = True
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub Fra_Photo_MouseDown(Button%, Shift%, X!, Y!)
    
    DragPhoto = True 'Enable dragging of the Photo
    
    'Set the mouse pointer to cross to show drag mode
    Screen.MousePointer = vbSizeAll
    
    Xaxis = X 'Capture the current X-axis position
    
End Sub

Private Sub Fra_Photo_MouseMove(Button%, Shift%, X!, Y!)
    DragMe X 'Call procedure
End Sub

Private Sub Fra_Photo_MouseUp(Button%, Shift%, X!, Y!)
    DragPhoto = False 'Disable drag mode
    Call AutoSizer.GetInitialPositions(Fra_Photo)
    Screen.MousePointer = vbDefault 'Set mouse pointer to default pointer to portray end of processing
End Sub

Private Sub ImgDBPhoto_MouseDown(Button%, Shift%, X!, Y!)
    DragPhoto = True 'Enable dragging of the Photo
    'Set the mouse pointer to cross to show drag mode
    Screen.MousePointer = vbSizeAll
    Xaxis = X 'Capture the current X-axis position
End Sub

Private Sub ImgDBPhoto_MouseMove(Button%, Shift%, X!, Y!)
    DragMe X 'Call procedure
End Sub

Private Sub ImgDBPhoto_MouseUp(Button%, Shift%, X!, Y!)
    DragPhoto = False 'Disable drag mode
    Call AutoSizer.GetInitialPositions(Fra_Photo)
    Screen.MousePointer = vbDefault 'Set mouse pointer to default pointer to portray end of processing
End Sub

