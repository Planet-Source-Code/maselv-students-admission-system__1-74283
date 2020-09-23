VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_StudentSports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Student Sports"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7095
   Icon            =   "Frm_StudentSports.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdAddName 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   360
      Picture         =   "Frm_StudentSports.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add Sport"
      Top             =   3240
      Width           =   345
   End
   Begin VB.CommandButton CmdAddStudent 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   2040
      Picture         =   "Frm_StudentSports.frx":0AF4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Add a Student"
      Top             =   960
      Width           =   345
   End
   Begin VB.TextBox txtPhoneNo 
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
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtPostalAddress 
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
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2160
      Width           =   4695
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
      Left            =   6600
      Picture         =   "Frm_StudentSports.frx":0E7E
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Move Last"
      Top             =   4680
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
      Left            =   6240
      Picture         =   "Frm_StudentSports.frx":11C0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Move Next"
      Top             =   4680
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
      Left            =   5880
      Picture         =   "Frm_StudentSports.frx":1502
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Move Previous"
      Top             =   4680
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
      Left            =   5520
      Picture         =   "Frm_StudentSports.frx":1844
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move First"
      Top             =   4680
      Width           =   375
   End
   Begin VB.Frame Fra_Photo 
      Height          =   1815
      Left            =   360
      TabIndex        =   19
      Tag             =   "XY"
      Top             =   660
      Width           =   1575
      Begin VB.Image ImgVirtualPhoto 
         Height          =   135
         Left            =   240
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image ImgDBPhoto 
         Height          =   1455
         Left            =   120
         Stretch         =   -1  'True
         Tag             =   "HW"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Fra_StudentSport 
      Height          =   4095
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   6855
      Begin VB.TextBox txtAdmNo 
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Frame Fra_Details 
         Height          =   1575
         Left            =   120
         TabIndex        =   32
         Top             =   2400
         Width           =   6615
         Begin VB.ComboBox cboAllocationStatus 
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
            ItemData        =   "Frm_StudentSports.frx":1B86
            Left            =   4200
            List            =   "Frm_StudentSports.frx":1B93
            Style           =   2  'Dropdown List
            TabIndex        =   13
            ToolTipText     =   "Whether the Student graduated, was suspended, transferred, dropped out or was discontinued"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtName 
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
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   480
            Width           =   5295
         End
         Begin VB.CommandButton cmdClearName 
            Height          =   315
            Left            =   480
            Picture         =   "Frm_StudentSports.frx":1BB2
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Clear selected Sport"
            Top             =   480
            Width           =   345
         End
         Begin VB.CommandButton cmdSelectName 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Left            =   840
            Picture         =   "Frm_StudentSports.frx":1F3C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Select a Sport"
            Top             =   480
            Width           =   345
         End
         Begin MSComCtl2.DTPicker dtStartDate 
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   59506691
            CurrentDate     =   31480
         End
         Begin MSComCtl2.DTPicker dtEndDate 
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   59506691
            CurrentDate     =   31480
         End
         Begin VB.Label lblAllocationStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allocation Status:"
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
            Left            =   4200
            TabIndex        =   12
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End Date:"
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
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sport:"
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
            TabIndex        =   4
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lblStartDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date:"
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
            TabIndex        =   8
            Top             =   840
            Width           =   810
         End
      End
      Begin VB.CommandButton CmdClearStudent 
         Height          =   315
         Left            =   2280
         Picture         =   "Frm_StudentSports.frx":24C6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Clear selected Student"
         Top             =   600
         Width           =   345
      End
      Begin VB.CommandButton CmdSelectStudent 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   2640
         Picture         =   "Frm_StudentSports.frx":2850
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Select a Student"
         Top             =   600
         Width           =   345
      End
      Begin VB.TextBox txtStudentName 
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
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtGender 
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
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtNationalID 
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
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1200
         Width           =   150
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00DFDFDF&
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
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1200
         Width           =   150
      End
      Begin VB.Shape ShpEmployee 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Height          =   2055
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label lblStudentName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Student Details:"
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
         TabIndex        =   0
         Top             =   360
         Width           =   1155
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
         Left            =   3120
         TabIndex        =   31
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lnlPhoneNo 
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
         TabIndex        =   30
         Top             =   960
         Width           =   750
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
         Left            =   1920
         TabIndex        =   29
         Top             =   1560
         Width           =   1125
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
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   -720
      Top             =   2760
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
            Picture         =   "Frm_StudentSports.frx":2DDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   4680
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
      Picture         =   "Frm_StudentSports.frx":3554
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
      TabIndex        =   20
      Tag             =   "Y"
      Top             =   4770
      Width           =   1395
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_StudentSports.frx":38EE
      Stretch         =   -1  'True
      Tag             =   "WY"
      Top             =   4560
      Width           =   7095
   End
   Begin VB.Image ImgHeader 
      Height          =   375
      Left            =   0
      Picture         =   "Frm_StudentSports.frx":40E4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7095
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
      Caption         =   "&Delete"
   End
   Begin VB.Menu MnuSearch 
      Caption         =   "Search"
   End
End
Attribute VB_Name = "Frm_StudentSports"
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

Public IsNewRecord As Boolean
Public FrmDefinitions$, SetPrivileges$

Private myRecIndex&, mySetNo&
Private myTable$, myTablePryKey$
Private myTableFixedFldName() As String
Private myTableParentFldName() As String
Private myRecDisplayON, IsLoading As Boolean

Public Function ClearEntries(Optional cSection& = &H0) As Boolean
On Local Error GoTo Handle_ClearEntries_Error
    
    Dim MousePointerState%
    Dim myRecDisplayState As Boolean
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayState = myRecDisplayON
    
    'Clear Sport details
    If (VBA.InStr(cSection, &H0) <> &H0) Or (VBA.InStr(cSection, &H2) <> &H0) Then
        Call cmdClearName_Click
    End If 'Close respective IF..THEN block statement
    
    'Clear other details
    If (VBA.InStr(cSection, &H0) <> &H0) Or (VBA.InStr(cSection, &H5) <> &H0) Then
        
        ShpBttnBriefNotes.TagExtra = VBA.vbNullString:
        myRecDisplayON = True: cboAllocationStatus.ListIndex = &H0: myRecDisplayON = myRecDisplayState
        
        VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
        
    End If 'Close respective IF..THEN block statement
    
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
    
    Fra_StudentSport.Enabled = Not State
    Fra_Details.Enabled = Not State
    
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
    
    VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM " & myTableFixedFldName(&H4) & " WHERE [" & myTableFixedFldName(&H0) & " ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'Assign the Record's primary key value
            Fra_StudentSport.Tag = vRs(myTableFixedFldName(&H0) & " ID")
            
            If Not VBA.IsNull(![Student ID]) Then
                
                'Display Student Details
                txtStudentName.Tag = ![Student ID]
                Call DisplayStudentDetails(![Student ID], False)
                
            End If 'Close respective IF..THEN block statement
            
            If Not VBA.IsNull(vRs(myTableParentFldName(&H0) & " ID")) Then txtName.Tag = vRs(myTableParentFldName(&H0) & " ID")
            
            If myTableParentFldName(&H0) = "Stream" Then txtName.Text = ![Class Name]
            If Not VBA.IsNull(vRs(myTableParentFldName(&H0) & " Name")) Then txtName.Text = VBA.Trim$(txtName.Text & " " & vRs(myTableParentFldName(&H0) & " Name"))
            
            If myTableParentFldName(&H0) = "Stream" Then dtStartDate.Value = ![Date Assigned] Else dtStartDate.Value = ![Start Date]
            If myTableParentFldName(&H0) <> "Stream" Then If Not VBA.IsNull(![End Date]) Then dtEndDate.Value = ![End Date] Else dtEndDate.Value = Null
            
            If Not VBA.IsNull(![Brief Notes]) Then ShpBttnBriefNotes.TagExtra = ![Brief Notes]
            If myTableParentFldName(&H0) <> "Stream" Then
                If Not VBA.IsNull(![Allocation Status]) Then If VBA.Trim$(![Allocation Status]) <> VBA.vbNullString Then cboAllocationStatus.Text = VBA.Trim$(![Allocation Status])
            Else
                cboAllocationStatus.ListIndex = VBA.IIf(![Discontinued], &H1, &H0)
            End If
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_DisplayRecord:
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources
    PerformMemoryCleanup
    
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

Private Function DisplayStudentDetails(StudentID&, Optional iPerformCleanUp As Boolean = True) As Boolean
On Local Error GoTo Handle_DisplayStudentDetails_Error
    
    Dim MousePointerState%
    Dim nRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim RecID&
    
    RecID = txtStudentName.Tag
    
    'Call event to clear Student's details if any
    Call CmdClearStudent_Click
    
    txtStudentName.Tag = RecID
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB(, False)   'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set nRs = New ADODB.Recordset
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT * FROM [Qry_Students] WHERE [Student ID] = " & StudentID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            If Not VBA.IsNull(![Student Name]) Then txtStudentName.Text = ![Student Name]
            If Not VBA.IsNull(![Admission Date]) Then txtAdmNo.Tag = ![Admission Date]
            If Not VBA.IsNull(![Adm No]) Then txtAdmNo.Text = ![Adm No]
            If Not VBA.IsNull(![Gender]) Then txtGender.Text = ![Gender]
            If Not VBA.IsNull(![National ID]) Then txtNationalID.Text = ![National ID]
            If Not VBA.IsNull(![Phone No]) Then txtPhoneNo.Text = ![Phone No]
            If Not VBA.IsNull(![Postal Address]) Then txtPostalAddress.Text = ![Postal Address]
            If Not VBA.IsNull(![Location]) Then txtLocation.Text = ![Location]
            
            'If the Record contains the Student's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display Student's Photo
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = nRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = txtStudentName.Text & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_DisplayStudentDetails:
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources, when needed
    If iPerformCleanUp Then PerformMemoryCleanup
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_DisplayStudentDetails_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_DisplayStudentDetails
    
End Function

Private Sub cboAllocationStatus_Click()
On Local Error GoTo Handle_cboAllocationStatus_Click_Error
    
    If myRecDisplayON Then Exit Sub
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, mySetNo, &H4, False, , True) Then myRecDisplayON = True: cboAllocationStatus.ListIndex = VBA.Val(cboAllocationStatus.Tag): myRecDisplayON = False: Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If cboAllocationStatus.ListIndex > &H0 Then If vMsgBox("Ticking this option will disable the " & myTableFixedFldName(&H2) & " allocation Record and will not be available in other Modules. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then myRecDisplayON = True: cboAllocationStatus.ListIndex = VBA.Val(cboAllocationStatus.Tag): myRecDisplayON = False: GoTo Exit_cboAllocationStatus_Click
    
    cboAllocationStatus.Tag = cboAllocationStatus.ListIndex
    
Exit_cboAllocationStatus_Click:
    
    myRecDisplayON = False 'Denote that Record display process is complete
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cboAllocationStatus_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cboAllocationStatus_Click
    
End Sub

Private Sub CmdAddName_Click()
On Local Error GoTo Handle_cmdAddParent3_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If txtName.Tag <> VBA.vbNullString Then vEditRecordID = txtName.Tag & "||||"
    Call LoadUnloadedFormViaStringName(Me, Frm_FamilyRelationships.Name, VBA.Join(myTableParentFldName, ":"))
    If vDatabaseAltered Then vDatabaseAltered = False: Call cmdClearName_Click
    
Exit_cmdAddParent3_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdAddParent3_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Loading Form- " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_cmdAddParent3_Click
    
End Sub

Private Sub CmdAddStudent_Click()
    If txtStudentName.Tag <> VBA.vbNullString Then vEditRecordID = txtStudentName.Tag & "||||"
    Frm_Students.Show vbModal, Me
    If vDatabaseAltered Then vDatabaseAltered = False: Call DisplayStudentDetails(VBA.Val(txtStudentName.Tag))
End Sub

Private Sub cmdClearName_Click()
    txtName.Tag = VBA.vbNullString: txtName.Text = VBA.vbNullString
End Sub

Private Sub CmdClearStudent_Click()
    
    txtStudentName.Text = VBA.vbNullString
    txtAdmNo.Tag = VBA.vbNullString
    txtAdmNo.Text = VBA.vbNullString
    txtGender.Text = VBA.vbNullString
    txtNationalID.Text = VBA.vbNullString
    txtPhoneNo.Text = VBA.vbNullString
    txtPostalAddress.Text = VBA.vbNullString
    txtLocation.Text = VBA.vbNullString
    
    'Clear Picture
    ImgVirtualPhoto.Picture = Nothing
    ImgVirtualPhoto.ToolTipText = VBA.vbNullString
    ImgDBPhoto.Picture = Nothing
    ImgDBPhoto.ToolTipText = VBA.vbNullString
    
End Sub

Private Sub CmdMoveRec_Click(Index As Integer)
On Local Error GoTo Handle_CmdMoveRec_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Navigate through Records in the specified table
    myRecIndex = NavigateToRec(Me, "SELECT * FROM " & myTableFixedFldName(&H4) & " WHERE [School Type] = " & School.Type & " ORDER BY [Adm No] ASC, [Student Name] ASC, " & VBA.IIf(myTableParentFldName(&H0) = "Stream", "[Date Assigned]", "[Start Date]") & " DESC, " & VBA.IIf(myTableParentFldName(&H0) = "Stream", "[Class Level] DESC, ", "") & "[" & myTableParentFldName(&H0) & " Name] ASC", myTableFixedFldName(&H0) & " ID", Index, myRecIndex)
    
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

Private Sub cmdSelectName_Click()
On Local Error GoTo Handle_cmdSelectName_Click_Error
    
    'If the User has not selected a Student Record then...
    If txtStudentName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a Student.", vbExclamation, App.Title & " : Student not Selected", Me
        Call CmdSelectStudent_Click 'Move the focus to the specified control
        If txtStudentName.Tag = VBA.vbNullString Then Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If myTableFixedFldName(&H0) = "Stream" Then
        
        'Enable User to select a Record
        vBuffer(&H0) = PickDetails(Me, "SELECT [Qry_Streams].[Stream ID], [Qry_Streams].[Class] FROM [Qry_Streams] WHERE ((([Qry_Streams].[Stream ID]) Not In (SELECT [Tbl_StudentStreams].[Stream ID] FROM [Tbl_StudentStreams] WHERE [Tbl_StudentStreams].[Student ID] = 2))) ORDER BY Qry_Streams.Class;", txtName.Tag, "1;2", "1", , , myTableParentFldName(&H3), , , , 3400)
        
    Else
        
        'Enable User to select a Record
        vBuffer(&H0) = PickDetails(Me, "SELECT DISTINCT [" & myTableParentFldName(&H4) & "].[" & myTableParentFldName(&H0) & " ID], [" & myTableParentFldName(&H4) & "].[" & myTableParentFldName(&H0) & " Name], [" & myTableParentFldName(&H4) & "].[Hierarchical Level] FROM [" & myTableParentFldName(&H4) & "] LEFT JOIN [" & myTable & "] ON [" & myTableParentFldName(&H4) & "].[" & myTableParentFldName(&H0) & " ID] = [" & myTable & "].[" & myTableParentFldName(&H0) & " ID] WHERE ((([" & myTableParentFldName(&H4) & "].[" & myTableParentFldName(&H0) & " ID]) NOT IN (SELECT [" & myTable & "].[" & myTableParentFldName(&H0) & " ID] FROM [" & myTable & "] WHERE [" & myTable & "].[Student ID] = " & VBA.Val(txtStudentName.Tag) & "))) ORDER BY [" & myTableParentFldName(&H4) & "].[Hierarchical Level] ASC, [" & myTableParentFldName(&H4) & "].[" & myTableParentFldName(&H0) & " Name];", txtName.Tag, "1;2", "1", , , myTableParentFldName(&H3), , , , 3400)
        
    End If 'Close respective IF..THEN block statement
    
    'If a Record has not been selected then Branch to the specified Label
    If vBuffer(&H0) = VBA.vbNullString Then GoTo Exit_cmdSelectName_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtName.Tag = vArrayList(&H0)  'Assign Record ID
    txtName.Text = vArrayList(&H1) 'Assign Record Name
    
    If Fra_StudentSport.Enabled Then dtStartDate.SetFocus 'Move the focus to the specified control
    
Exit_cmdSelectName_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_cmdSelectName_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Select Category Error " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_cmdSelectName_Click
    
End Sub

Private Sub CmdSelectStudent_Click()
On Local Error GoTo Handle_CmdSelectStudent_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Student ID], ([Registered Name]) AS [Student Name], [Adm No], [Gender], [National ID] FROM [Qry_Students] ORDER BY [Registered Name] ASC,[Adm No] ASC;", txtStudentName.Tag, , "1", , , "Students", , "[Tbl_Students]; WHERE [Student ID] = $;1;Student Name", , 6000)
    
    'If a Record has not been selected then Branch to the specified Label
    If vBuffer(&H0) = VBA.vbNullString Then GoTo Exit_CmdSelectStudent_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    Call cmdClearName_Click
    
    txtStudentName.Tag = vArrayList(&H0)
    
    'Call event to display the selected Student's details
    Call DisplayStudentDetails(VBA.CLng(vArrayList(&H0)))
    
    If Fra_StudentSport.Enabled Then cmdSelectName.SetFocus 'Move the focus to the specified control
    
Exit_CmdSelectStudent_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_CmdSelectStudent_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_CmdSelectStudent_Click
    
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
    
    'Set Defaults
    Select Case myTableFixedFldName(&H0)
        
        Case "Prefect": mySetNo = &H21
        Case "Student Sport": mySetNo = &H1B
        Case "Student Club": mySetNo = &H1D
        Case "Student Society": mySetNo = &H1F
        Case "Student Subject": mySetNo = &H26
        Case Else: mySetNo = &H1C
        
    End Select
    
    dtStartDate.Value = VBA.Date: dtEndDate.Value = Null
    
    'Yield execution so that the operating system can process other events
    VBA.DoEvents
    
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
            .Filter = "[" & myTableFixedFldName(&H0) & " ID] = " & vArrayList(&H0)
            
            'If the record Exists then Call Procedure in this Form to display it
            If Not (.BOF And .EOF) Then DisplayRecord VBA.CLng(VBA.Val(vRs(myTableFixedFldName(&H0) & " ID")))
            
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
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources
    PerformMemoryCleanup
    
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
    
    If FrmDefinitions = VBA.vbNullString Then
        
        FrmDefinitions = "Student Club:Student Clubs:Student Club:Student Clubs:Qry_StudentClubs::|" & _
                        "Club:Clubs:Club:Clubs:Tbl_Clubs:Frm_Dormitories:Club~Clubs~Club~Clubs~Tbl_Clubs~~"
                        
'        FrmDefinitions = "Student Stream:Student Streams:Student Stream:Student Streams:Qry_StudentStreams::|" & _
'                        "Stream:Streams:Stream:Streams:Tbl_Streams:Frm_Streams:Stream~Streams~Stream~Streams~Tbl_Streams~~"
                        
                        '"Student Sport:Student Sports:Student Sport:Student Sports:Qry_StudentSports::|" & _
                        '"Sport:Sports:Sport:Sports:Tbl_Sports:Frm_Dormitories:Sport~Sports~Sport~Sports~Tbl_Sports~~"
                        
    End If 'Close respective IF..THEN block statement
    
    vArrayList = VBA.Split(FrmDefinitions, "|")
    
    myTableFixedFldName = VBA.Split(vArrayList(&H0), ":")
    myTableParentFldName = VBA.Split(vArrayList(&H1), ":")
    
    myTable = "Tbl_" & VBA.Replace(myTableFixedFldName(&H1), " ", VBA.vbNullString)
    
    Me.Caption = App.Title & " : " & myTableFixedFldName(&H3)
    
    lblName.Caption = myTableParentFldName(&H2) & ":"
    CmdAddName.ToolTipText = "Add " & myTableParentFldName(&H2)
    cmdClearName.ToolTipText = "Clear selected " & myTableParentFldName(&H2)
    cmdSelectName.ToolTipText = "Select a " & myTableParentFldName(&H2)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = Not CloseFrm(Me) 'Call Procedure in Mdl_Stadmis Module to confirm Form closure
End Sub

Private Sub Fra_Photo_Click()
    Call ImgDBPhoto_Click
End Sub

Private Sub ImgDBPhoto_Click()
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, True)
End Sub

Private Sub MnuDelete_Click()
On Local Error GoTo Handle_MnuDelete_Click_Error
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, mySetNo, &H3, False) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the User has not selected any existing Record then...
    If Fra_StudentSport.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a " & myTableFixedFldName(&H2) & " Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    If vMsgBox("Are you sure you want to DELETE the displayed Record?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM " & myTable & " WHERE [" & myTableFixedFldName(&H0) & " ID] = " & (Fra_StudentSport.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
    
    'Denote that the database table has been altered
    vDatabaseAltered = True
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    'Get the total number of Records already saved
    lblRecords.Tag = VBA.Val(lblRecords.Tag) - &H1
    lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "The displayed " & myTableFixedFldName(&H2) & " Record has successfully been deleted.", vbInformation, App.Title & " : Delete", Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
Exit_MnuDelete_Click:
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources
    PerformMemoryCleanup
    
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
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, mySetNo, &H2, False) Then Exit Sub
    
    'If the User has not selected any existing Record then...
    If Fra_StudentSport.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a " & myTableFixedFldName(&H2) & " Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        Exit Sub 'Quit this saving procedure
        
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
    
    CmdSelectStudent.SetFocus 'Move focus to Name textbox
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuNew_Click()
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, mySetNo, &H1, False) Then Exit Sub
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    Call LockEntries(False) 'Call Procedure in this Form to UnLock Input Controls
    
    IsNewRecord = True 'Denote that the displayed Record does not exist in the database
    
    CmdSelectStudent.SetFocus 'Move focus to Name textbox
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuSave_Click()
On Local Error GoTo Handle_MnuSave_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the User has not selected a Student Record then...
    If txtStudentName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a Student.", vbExclamation, App.Title & " : Student not Selected", Me
        CmdSelectStudent.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not selected a Sport then...
    If txtName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a " & myTableParentFldName(&H2) & ".", vbExclamation, App.Title & " : " & myTableParentFldName(&H2) & " not Selected", Me
        cmdSelectName.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the selected Start Date is earlier than when the Student was admitted then...
    If dtStartDate.Value < txtAdmNo.Tag Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a Start Date later than the date the Student was Admitted {" & VBA.Format$(txtAdmNo.Tag, "ddd dd MMM yyyy") & "}", vbExclamation, App.Title & " : " & myTableParentFldName(&H2) & " not Selected", Me
        dtStartDate.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the End Date has been specified then...
    If Not VBA.IsNull(dtEndDate.Value) Then
        
        'If the selected End Date is earlier than the Start Date then...
        If dtStartDate.Value > dtEndDate.Value Then
            
            'Warn User to select an existing Record
            vMsgBox "Please select an End Date later than the specified Start Date.", vbExclamation, App.Title & " : " & myTableParentFldName(&H2) & " not Selected", Me
            dtEndDate.SetFocus 'Move the focus to the specified control
            Exit Sub 'Quit this saving procedure
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    'Start the saving process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[" & myTableFixedFldName(&H0) & " ID] = " & (Fra_StudentSport.Tag): .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        ![Student ID] = txtStudentName.Tag
        vRs(myTableParentFldName(&H0) & " ID") = txtName.Tag
        If dtEndDate.Visible Then ![Start Date] = VBA.FormatDateTime(dtStartDate.Value, vbShortDate) Else ![Date Assigned] = VBA.FormatDateTime(dtStartDate.Value, vbShortDate)
        If dtEndDate.Visible Then If VBA.IsNull(dtEndDate.Value) Then ![End Date] = Null Else ![End Date] = VBA.FormatDateTime(dtEndDate.Value, vbShortDate)
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![Allocation Status] = cboAllocationStatus.Text
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        Fra_StudentSport.Tag = vRs(myTableFixedFldName(&H0) & " ID")
        
        .Close 'Close the opened object and any dependent objects
        
        'Denote that the database table has been altered
        vDatabaseAltered = True
        
        'Get the total number of Records already saved
        lblRecords.Tag = VBA.Val(lblRecords.Tag) + VBA.IIf(IsNewRecord, &H1, &H0)
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
        LockEntries True 'Call Procedure in this Form to Lock Input Controls
        
    End With 'Close the WITH block statements
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "The Record has successfully been " & VBA.IIf(IsNewRecord, "Saved", "Modified"), vbInformation, App.Title & " : Saving Report", Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    IsNewRecord = True 'Cancel Edit Mode
    
Exit_MnuSave_Click:
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources
    PerformMemoryCleanup
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_MnuSave_Click_Error:
    
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [" & myTableFixedFldName(&H0) & " ID], [Student Name], [Adm No], [Gender], " & VBA.IIf(myTableParentFldName(&H0) = "Stream", "[Class Name], ", "") & "[" & myTableParentFldName(&H0) & " Name] FROM " & myTableFixedFldName(&H4) & " ORDER BY [Student Name] ASC, [Adm No] ASC, " & VBA.IIf(myTableParentFldName(&H0) = "Stream", "[Date Assigned]", "[Start Date]") & " DESC, " & VBA.IIf(myTableParentFldName(&H0) = "Stream", "[Class Level] DESC, ", "") & "[Hierarchical Level] ASC, [" & myTableParentFldName(&H0) & " Name] ASC", (Fra_StudentSport.Tag), , "1", , , myTableParentFldName(&H3), , "[Qry_Students]; WHERE [Student ID] IN (SELECT [Student ID] FROM " & myTable & " WHERE [" & myTableFixedFldName(&H0) & " ID] = $)", , 8500)
    
    'If a Record has not been selected then Branch to the specified Label
    If vBuffer(&H0) = VBA.vbNullString Then GoTo Exit_MnuSearch_Click
    
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

Private Sub ShpBttnBriefNotes_Click()
    'Call Function in Mdl_Stadmis to display Notes Input Form
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_StudentSport.Enabled, "Student Sports")
End Sub
