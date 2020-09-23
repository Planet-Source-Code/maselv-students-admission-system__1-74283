VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Frm_StudentClasses 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Student Classes"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8910
   Icon            =   "Frm_StudentClasses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAddDormitory 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   3480
      Picture         =   "Frm_StudentClasses.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Add Dormitory"
      Top             =   3960
      Width           =   345
   End
   Begin VB.Frame Fra_StudentClass 
      BackColor       =   &H00CFE1E2&
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   8655
      Begin VB.CommandButton cmdClearClass 
         Height          =   315
         Left            =   960
         Picture         =   "Frm_StudentClasses.frx":0AF4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Clear selected Class"
         Top             =   360
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectClass 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   1320
         Picture         =   "Frm_StudentClasses.frx":0E7E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Select a Class"
         Top             =   360
         Width           =   345
      End
      Begin VB.TextBox txtClassName 
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
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Top             =   360
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtYear 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   58785795
         UpDown          =   -1  'True
         CurrentDate     =   31480
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         TabIndex        =   0
         Top             =   120
         Width           =   390
      End
      Begin VB.Label lblClassName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Name:"
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
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   885
      End
   End
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   3840
      TabIndex        =   27
      Top             =   4740
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
      Picture         =   "Frm_StudentClasses.frx":1408
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
   Begin VB.CommandButton cmdAddStudent 
      BackColor       =   &H00D8E9EC&
      Height          =   315
      Left            =   5160
      Picture         =   "Frm_StudentClasses.frx":17A2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Add a Student"
      Top             =   2160
      Width           =   360
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00CFE1E2&
      Height          =   1815
      Left            =   3480
      TabIndex        =   19
      Tag             =   "XY"
      Top             =   1920
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
         Width           =   1350
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
      Index           =   3
      Left            =   8400
      Picture         =   "Frm_StudentClasses.frx":1B2C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Move Last"
      Top             =   4740
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
      Left            =   8040
      Picture         =   "Frm_StudentClasses.frx":1E6E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move Next"
      Top             =   4740
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
      Left            =   7680
      Picture         =   "Frm_StudentClasses.frx":21B0
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Move Previous"
      Top             =   4740
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
      Left            =   7320
      Picture         =   "Frm_StudentClasses.frx":24F2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Move First"
      Top             =   4740
      Width           =   375
   End
   Begin VB.Frame Fra_StudentClasses 
      BackColor       =   &H00CFE1E2&
      Height          =   3375
      Left            =   3240
      TabIndex        =   17
      Top             =   1080
      Width           =   5535
      Begin VB.CommandButton cmdClearDormitory 
         Height          =   315
         Left            =   600
         Picture         =   "Frm_StudentClasses.frx":2834
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Clear selected Dormitory"
         Top             =   2880
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectDormitory 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   960
         Picture         =   "Frm_StudentClasses.frx":2BBE
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Select a Dormitory"
         Top             =   2880
         Width           =   345
      End
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
         Left            =   1320
         TabIndex        =   36
         Top             =   2880
         Width           =   2175
      End
      Begin VB.ComboBox cboStudentClassStatus 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "Frm_StudentClasses.frx":3148
         Left            =   3600
         List            =   "Frm_StudentClasses.frx":3158
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtNewYear 
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
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   31
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNewClass 
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
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   480
         Width           =   3615
      End
      Begin VB.CommandButton cmdSelectNewClass 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   1320
         Picture         =   "Frm_StudentClasses.frx":3185
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Select a Class"
         Top             =   480
         Width           =   345
      End
      Begin VB.CommandButton cmdClearNewClass 
         Height          =   315
         Left            =   960
         Picture         =   "Frm_StudentClasses.frx":370F
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Clear selected Class"
         Top             =   480
         Width           =   345
      End
      Begin VB.CheckBox chkHideAlreadyEntered 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Hide those already entered"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3120
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtAdmNo 
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdClearStudent 
         Height          =   315
         Left            =   2280
         Picture         =   "Frm_StudentClasses.frx":3A99
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Clear selected Student"
         Top             =   1080
         Width           =   345
      End
      Begin VB.CommandButton cmdSelectStudent 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   2640
         Picture         =   "Frm_StudentClasses.frx":3E23
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Select a Student"
         Top             =   1080
         Width           =   345
      End
      Begin VB.TextBox txtStudentName 
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
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtGender 
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
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1455
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2280
         Width           =   3375
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
         Left            =   240
         TabIndex        =   39
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class Allocation Status:"
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
         TabIndex        =   34
         Top             =   2640
         Width           =   1680
      End
      Begin VB.Label lblNewYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   390
      End
      Begin VB.Label lblNewClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Class Name:"
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
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lblStudentName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Details:"
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
         TabIndex        =   8
         Top             =   840
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
         Left            =   3840
         TabIndex        =   26
         Top             =   1440
         Width           =   585
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
         Left            =   1920
         TabIndex        =   25
         Top             =   2040
         Width           =   660
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
         TabIndex        =   24
         Top             =   1440
         Width           =   615
      End
   End
   Begin MSChart20Lib.MSChart MSChart 
      Height          =   3615
      Left            =   -120
      OleObjectBlob   =   "Frm_StudentClasses.frx":43AD
      TabIndex        =   35
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Image ImgHeader 
      Height          =   375
      Left            =   0
      Picture         =   "Frm_StudentClasses.frx":5F1D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9135
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
      TabIndex        =   18
      Top             =   4830
      Width           =   1395
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   -240
      Picture         =   "Frm_StudentClasses.frx":6713
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   9135
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
Attribute VB_Name = "Frm_StudentClasses"
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

Public vRecordID&
Public IsNewRecord As Boolean
Public FrmDefinitions$, SetPrivileges$

Private myRecIndex&, iSetNo&
Private myRecDisplayON, IsLoading As Boolean

Public iSelColPos&
Public iTargetFields$, iLvwItemDataType$, iPhotoSpecifications$, iSearchIDs$

Private Xaxis&
Private iShiftKey%
Private DragPhoto, iSelectionComplete As Boolean

Public Function ClearEntries() As Boolean
On Local Error GoTo Handle_ClearEntries_Error
    
    Dim MousePointerState%
    Dim myRecDisplayState As Boolean
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayState = myRecDisplayON
    
    Call CmdClearStudent_Click
    dtYear.Tag = VBA.vbNullString
    cboStudentClassStatus.ListIndex = &H0
    cmdClearDormitory.Tag = VBA.vbNullString
    cmdSelectDormitory.Tag = VBA.vbNullString
    txtStudentName.LinkItem = VBA.vbNullString
    ShpBttnBriefNotes.TagExtra = VBA.vbNullString:
    
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
    
    Fra_StudentClasses.Enabled = Not State
    
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
        
        .Open "SELECT * FROM [Qry_StudentClasses] WHERE [Student Class ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            'Assign the Record's primary key value
            Fra_StudentClasses.Tag = ![Student Class ID]
            
            If Not VBA.IsNull(![Student ID]) Then
                
                'Display Student Details
                txtStudentName.Tag = ![Student ID]
                txtStudentName.LinkItem = ![Student ID]
                Call DisplayStudentDetails(![Student ID], False)
                
            End If 'Close respective IF..THEN block statement
            
            If Not VBA.IsNull(![Date Assigned]) Then
                
                dtYear.Value = VBA.DateSerial(VBA.Year(![Date Assigned]), &H1, &H1): dtYear.Tag = dtYear.Value
                If Not VBA.IsNull(![Previous Class]) Then If VBA.Trim$(![Previous Class]) <> VBA.vbNullString Then dtYear.Value = VBA.DateSerial(VBA.Year(![Date Assigned]) - &H1, &H1, &H1): dtYear.Tag = dtYear.Value
                txtNewYear.Text = VBA.Year(![Date Assigned])
                
            End If 'Close respective IF..THEN block statement
            
            If Not VBA.IsNull(![Class]) Then txtNewClass.Text = ![Class]
            If Not VBA.IsNull(![Stream ID]) Then txtNewClass.Tag = ![Stream ID]: cmdSelectNewClass.Tag = ![Stream ID]
            If Not VBA.IsNull(![Class Level]) Then cmdClearNewClass.Tag = ![Class Level]
            
            txtClassName.LinkItem = VBA.vbNullString
            txtClassName.Text = txtNewClass.Text: If Not VBA.IsNull(![Previous Class]) Then If VBA.Trim$(![Previous Class]) <> VBA.vbNullString Then txtClassName.LinkItem = ![Previous Class]: txtClassName.Text = ![Previous Class]
            txtClassName.Tag = txtNewClass.Tag: If Not VBA.IsNull(![Previous Stream ID]) Then txtClassName.Tag = ![Previous Stream ID]
            cmdSelectClass.Tag = txtClassName.Tag: If Not VBA.IsNull(![Previous Class Capacity]) Then If VBA.Trim$(![Previous Class Capacity]) <> VBA.vbNullString Then cmdSelectClass.Tag = ![Previous Class Capacity]
            cmdClearClass.Tag = cmdClearNewClass.Tag: If Not VBA.IsNull(![Previous Class Level]) Then cmdClearClass.Tag = ![Previous Class Level]
            
            Call RefreshClassDetails(VBA.Val(txtNewClass.Tag))
            
            If Not VBA.IsNull(![Student Class Status]) Then If VBA.Trim$(![Student Class Status]) <> VBA.vbNullString Then cboStudentClassStatus.Text = ![Student Class Status]
            
            .Close 'Close the opened object and any dependent objects
            
            .Open "SELECT [Student ID] FROM [Tbl_StudentDormitories] WHERE [Student ID] = " & VBA.Val(txtStudentName.Tag) & " AND YEAR([Date Assigned]) = " & VBA.Year(dtYear.Value), vAdoCNN, adOpenKeyset, adLockReadOnly
            
            'If there are records in the table then...
            If Not (.BOF And .EOF) Then
                
                If Not VBA.IsNull(![Student Dormitory ID]) Then cmdClearDormitory.Tag = ![Student Dormitory ID]
                If Not VBA.IsNull(![Dormitory ID]) Then txtDormitoryName.Tag = ![Dormitory ID]: cmdSelectDormitory.Tag = ![Dormitory ID]
                If Not VBA.IsNull(![Dormitory Name]) Then txtDormitoryName.Text = ![Dormitory Name]
                
            End If 'Close respective IF..THEN block statement
            
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
    
    ConnectDB , False 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT * FROM [Qry_Students] WHERE [Student ID] = " & StudentID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            CmdSelectStudent.Tag = ![Student ID]
            If Not VBA.IsNull(![Student Name]) Then txtStudentName.Text = ![Student Name]
            If Not VBA.IsNull(![Admission Date]) Then txtAdmNo.Tag = ![Admission Date]
            If Not VBA.IsNull(![Adm No]) Then txtAdmNo.Text = ![Adm No]
            If Not VBA.IsNull(![Gender]) Then txtGender.Text = ![Gender]
            
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
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Displaying Student Details - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_DisplayStudentDetails
    
End Function

Private Function RefreshClassDetails(StreamID&, Optional iPerformCleanUp As Boolean = True) As Boolean
On Local Error GoTo Handle_RefreshClassDetails_Error
    
    'If the User has not selected a Class then...
    If VBA.LenB(VBA.Trim$(txtClassName.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a Class.", vbExclamation, App.Title & " : Record not Selected"
        Call cmdSelectClass_Click 'Move the focus to the specified control
        
        'If the User has not yet selected a Class then...
        If txtClassName.Tag = VBA.vbNullString Then Exit Function 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    Dim nRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
        
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With MSChart 'Execute a series of statements on vRs recordset
        
        .ColumnCount = 1
        
        If VBA.Val(cmdSelectClass.Tag) > &H0 Then
            
            .ColumnCount = &H2: .Column = &H1
            .ColumnLabel = "Class Capacity"
            .data = VBA.Val(cmdSelectClass.Tag)
            
        Else
            
        End If
        
        Set vRs = New ADODB.Recordset
        vRs.Open "SELECT * FROM [Qry_StudentClasses] WHERE YEAR([Date Assigned]) = " & VBA.Val(txtNewYear.Text) & " AND [Stream ID] = " & StreamID&, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        .Column = .ColumnCount
        .ColumnLabel = "Total Students"
        .RowLabel = VBA.vbNullString ' txtClassName.Text
        .data = vRs.RecordCount
        
        vRs.Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_RefreshClassDetails:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_RefreshClassDetails_Error:
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Refreshing Class Details - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_RefreshClassDetails
    
End Function

Private Sub cboStudentClassStatus_Click()
    
    If myRecDisplayON Then Exit Sub
    
    'If the User has not selected a New Class then...
    If VBA.LenB(VBA.Trim$(txtNewClass.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a New Class.", vbExclamation, App.Title & " : Record not Selected"
        Call cmdSelectNewClass_Click 'Move the focus to the specified control
        
        'If the User has not yet selected a Class then...
        If txtNewClass.Tag = VBA.vbNullString Then Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub cmdAddDormitory_Click()
    
    If VBA.LenB(VBA.Trim$(txtDormitoryName.Tag)) <> &H0 Then vEditRecordID = txtDormitoryName.Tag & "||||"
    FrmDefinitions = "Dormitory:Dormitories:Dormitory:Dormitories:Tbl_Dormitories:Tbl_StudentDormitories:Students"
    Frm_Dormitories.Show vbModal, Me
    If vDatabaseAltered And Fra_StudentClasses.Enabled Then vDatabaseAltered = False: Call cmdClearDormitory_Click
    
End Sub

Private Sub CmdAddStudent_Click()
    
    If VBA.LenB(VBA.Trim$(txtStudentName.Tag)) <> &H0 Then vEditRecordID = txtStudentName.Tag & "|||"
    Call LoadUnloadedFormViaStringName(Me, Frm_Students.Name)
    If vDatabaseAltered Then vDatabaseAltered = False: Call CmdClearStudent_Click
    
End Sub

Private Sub cmdClearClass_Click()
    txtClassName.Tag = VBA.vbNullString: txtClassName.Text = VBA.vbNullString: cmdClearClass.Tag = VBA.vbNullString
End Sub

Private Sub cmdClearDormitory_Click()
    
    'If the User is not allowed to alter the Student's allocated Dormitory then quit this procedure
    If cmdSelectDormitory.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &HB, &H2) Then Exit Sub
    
    txtDormitoryName.Tag = VBA.vbNullString: txtDormitoryName.Text = VBA.vbNullString
    
End Sub

Private Sub cmdClearNewClass_Click()
    txtNewClass.Tag = VBA.vbNullString: txtNewClass.Text = VBA.vbNullString: cmdClearNewClass.Tag = VBA.vbNullString
End Sub

Private Sub CmdClearStudent_Click()
    
    txtStudentName.Text = VBA.vbNullString
    txtAdmNo.Tag = VBA.vbNullString
    txtAdmNo.Text = VBA.vbNullString
    txtGender.Text = VBA.vbNullString
    txtLocation.Text = VBA.vbNullString
    CmdSelectStudent.Tag = VBA.vbNullString
    
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
    myRecIndex = NavigateToRec(Me, "SELECT [Student Class ID] FROM [Qry_StudentClasses] WHERE NOT ([Previous Class] IS NULL OR TRIM([Previous Class])='') ORDER BY [Date Assigned] DESC, [Class Level] DESC, [Registered Name] ASC;", "Student Class ID", Index, myRecIndex)
    
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [Stream ID], ([Class Name] & ' ' & [Stream Name]) AS [Class], [Capacity], [Class Level] FROM [Qry_Streams] WHERE [Discontinued] = FALSE  ORDER BY [Class Level] ASC, [Stream Name] ASC", cmdSelectClass.Tag, "1;2;3;4", "1;4", , , "Classes", , , , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectClass_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtClassName.Tag = vArrayList(&H0) 'Assign Stream ID
    txtClassName.Text = vArrayList(&H1) 'Assign Class Name
    cmdSelectClass.Tag = vArrayList(&H2) 'Assign Class Capacity
    cmdClearClass.Tag = vArrayList(&H3) 'Assign Class Level
    
    cmdSelectNewClass.SetFocus
    
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

Private Sub cmdSelectNewClass_Click()
On Local Error GoTo Handle_cmdSelectClass_Click_Error
    
    'If the User is not allowed to alter the Student's allocated class then quit this procedure
    If cmdSelectClass.Tag <> VBA.vbNullString Then If Not ValidUserAccess(Me, &H9, &H2) Then Exit Sub
    
    'If the User has not selected a Class then...
    If VBA.LenB(VBA.Trim$(txtClassName.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a Class.", vbExclamation, App.Title & " : Record not Selected", Me
        Call cmdSelectClass_Click 'Move the focus to the specified control
        
        'If the User has not yet selected a Class then...
        If txtClassName.Tag = VBA.vbNullString Then Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the new year exceeds the current system year then...
    If VBA.Val(txtNewYear.Text) > VBA.Year(VBA.Date) Then
        
        'Warn User to select an existing Record
        vMsgBox "The Software cannot perform early registration of Students. Please wait for " & txtNewYear.Text & " or check your Computer's date.", vbExclamation, App.Title & " : Record not Selected", Me
        dtYear.SetFocus
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Stream ID], ([Class Name] & ' ' & [Stream Name]) AS [Class], [Capacity], [Class Level] FROM [Qry_Streams] WHERE [Discontinued] = FALSE  ORDER BY [Class Level] ASC, [Stream Name] ASC", txtNewClass.Tag, "1;2;3;4", "1;4", , , "Classes", , , , 4500)
    
    'If a Record has not been selected then Branch to the specified Label
    If VBA.LenB(vBuffer(&H0)) = &H0 Then GoTo Exit_cmdSelectClass_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    'Display the selected Record
    txtNewClass.Tag = vArrayList(&H0) 'Assign Stream ID
    txtNewClass.Text = vArrayList(&H1) 'Assign Class Name
    cmdSelectNewClass.Tag = vArrayList(&H2) 'Assign Class Capacity
    cmdClearNewClass.Tag = vArrayList(&H3) 'Assign Class Level
    
    Call RefreshClassDetails(VBA.Val(txtNewClass.Tag))
    
    CmdSelectStudent.SetFocus
    
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

Private Sub CmdSelectStudent_Click()
On Local Error GoTo Handle_CmdSelectStudent_Click_Error
    
    'If the User has not selected a New Class then...
    If VBA.LenB(VBA.Trim$(txtNewClass.Tag)) = &H0 Then
        
        'Warn User to select an existing Record
        vMsgBox "Please select a New Class.", vbExclamation, App.Title & " : Record not Selected", Me
        Call cmdSelectNewClass_Click 'Move the focus to the specified control
        
        'If the User has not yet selected a Class then...
        If txtNewClass.Tag = VBA.vbNullString Then Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Enable User to select a Record
    vBuffer(&H0) = PickDetails(Me, "SELECT [Student ID], [Adm No], [Student Name], [Gender], [Location] FROM [Qry_Students] WHERE [Student ID] IN (SELECT [Student ID] FROM [Tbl_StudentClasses] WHERE [Stream ID] = " & VBA.Val(txtClassName.Tag) & " AND YEAR([Date Assigned]) = " & VBA.Year(dtYear.Value) & ")" & VBA.IIf(chkHideAlreadyEntered.Value = vbChecked, " AND [Student ID] NOT IN (SELECT [Student ID] FROM [Tbl_StudentClasses] WHERE YEAR([Date Assigned]) = " & VBA.Val(txtNewYear.Text), VBA.vbNullString) & ") ORDER BY [Student Name] ASC;", txtStudentName.Tag, , "1", , , "Students", , "[Qry_Students]; WHERE [Student ID] = $", , 8600)
    
    'If a Record has not been selected then Branch to the specified Label
    If vBuffer(&H0) = VBA.vbNullString Then GoTo Exit_CmdSelectStudent_Click
    
    vArrayListTmp = VBA.Split(vMultiSelectedData, "|")
    vArrayList = VBA.Split(vArrayListTmp(&H0), "~")
    
    txtStudentName.Tag = vArrayList(&H0)
    
    'Call event to display the selected Student's details
    Call DisplayStudentDetails(VBA.CLng(vArrayList(&H0)))
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT [Student ID] FROM [Tbl_StudentDormitories] WHERE [Student ID] = " & VBA.Val(txtStudentName.Tag) & " AND YEAR([Date Assigned]) = " & VBA.Year(dtYear.Value), vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            If Not VBA.IsNull(![Student Dormitory ID]) Then cmdClearDormitory.Tag = ![Student Dormitory ID]
            If Not VBA.IsNull(![Dormitory ID]) Then txtDormitoryName.Tag = ![Dormitory ID]: cmdSelectDormitory.Tag = ![Dormitory ID]
            If Not VBA.IsNull(![Dormitory Name]) Then txtDormitoryName.Text = ![Dormitory Name]
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
    If Fra_StudentClasses.Enabled Then cmdSelectClass.SetFocus 'Move the focus to the specified control
    
Exit_CmdSelectStudent_Click:
    
    'Reinitialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vBuffer: Erase vArrayList: vMultiSelectedData = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_CmdSelectStudent_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Student - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_CmdSelectStudent_Click
    
End Sub

Private Sub dtYear_Change()
    Call CmdClearStudent_Click: Fra_StudentClass.Tag = VBA.vbNullString: txtNewYear.Text = VBA.Year(dtYear.Value) + &H1
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
    iSetNo = &H9
    dtYear.Value = VBA.Date: dtYear.Tag = dtYear.Value: txtNewYear.Text = VBA.Year(dtYear.Value) + &H1
    myRecDisplayON = True: cboStudentClassStatus.ListIndex = &H0: myRecDisplayON = False
    
    'Yield execution so that the operating system can process other events
    VBA.DoEvents
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to clear all entries in Input Boxes
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records in the specified database table
        .Open "SELECT * FROM [Tbl_StudentClasses]", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'Get the total number of Records already saved
        lblRecords.Tag = .RecordCount
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        'If a Record is to be displayed then...
        If vEditRecordID <> VBA.vbNullString Then
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            'Retrieve the Record with the specified ID
            .Filter = "[Student Class ID] = " & vArrayList(&H0)
            
            'If the record Exists then Call Procedure in this Form to display it
            If Not (.BOF And .EOF) Then DisplayRecord VBA.CLng(VBA.Val(![Student Class ID]))
            
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
    
    Exit Sub 'Quit this Procedure
    
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
    Me.Caption = App.Title & " : Student Fee Allocation"
    
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
On Error GoTo Handle_MnuDelete_Click_Error
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, iSetNo, &H3, False) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the User has not selected any existing Record then...
    If Fra_StudentClasses.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Student Class Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    If vMsgBox("Are you sure you want to DELETE the displayed Record?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    Dim iDependency As Boolean
    Dim iDependencyArray() As String
    
    iDependencyArray = VBA.Split("Tbl_StudentFees:Student Fees", "|")
    
    'For each entry...
    For vIndex(&H0) = &H0 To UBound(iDependencyArray) Step &H1
        
        'Check if there are Records in other tables depending on the displayed Record
        iDependency = CheckForRecordDependants(Me, iDependencyArray(vIndex(&H0)) & " WHERE [Student Class ID] = " & (Fra_StudentClasses.Tag))
        
        'If the Record is depended upon then Quit this FOR..LOOP block statement
        If iDependency Then Exit For
        
    Next vIndex(&H0) 'Move to the next entry
    
    'If the Records exist then...
    If iDependency Then
        
        'Warn User
        vMsgBox "The displayed Student Classes Record has other Records depending on it. Delete operation aborted", vbExclamation, App.Title & " : Operation Aborted", Me
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM [Tbl_StudentClasses] WHERE [Student Class ID] = " & (Fra_StudentClasses.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
    
    'Denote that the database table has been altered
    vDatabaseAltered = True
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    'Get the total number of Records already saved
    lblRecords.Tag = VBA.Val(lblRecords.Tag) - &H1
    lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "The displayed Student Classes Record has successfully been deleted.", vbInformation, App.Title & " : Delete", Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
Exit_MnuDelete_Click:
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources
    PerformMemoryCleanup
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
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
    If Not ValidUserAccess(Me, iSetNo, &H2, False) Then Exit Sub
    
    'If the User has not selected any existing Record then...
    If Fra_StudentClasses.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Student Class Record.", vbExclamation, App.Title & " : No Record Displayed", Me
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
    If Not ValidUserAccess(Me, iSetNo, &H1, False) Then Exit Sub
    
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
On Error GoTo Handle_MnuSave_Click_Error
    
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
    
    'Start the saving process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Tbl_StudentClasses]", vAdoCNN, adOpenKeyset, adLockPessimistic
        
        .Filter = "[Student ID] = " & VBA.Val(txtStudentName.Tag) & " AND [Date Assigned] >= #" & VBA.DateSerial(VBA.Val(txtNewYear.Text), &H1, &H1) & "# AND [Date Assigned] <= #" & VBA.DateSerial(VBA.Val(txtNewYear.Text), &HC, 31) & "#"
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Or Not (.BOF And .EOF) Then .Update Else .AddNew
        
        'Assign entries to their respective database fields
        ![Date Assigned] = VBA.DateSerial(VBA.Val(txtNewYear.Text), &H1, &H1)
        ![Student ID] = VBA.Val(txtStudentName.Tag)
        ![Stream ID] = txtNewClass.Tag
        ![Previous Stream ID] = txtClassName.Tag
        
'        'Get the Student's total School Fees for the previous Years
'        vRsTmp.Open "SELECT SUM([Amount]) AS [Total Fees] FROM [Qry_StudentFees] WHERE [Student ID] = " & VBA.Val(txtStudentName.Tag) & " AND [Fee Structure Year] < " & VBA.Year(dtYear.Value), vAdoCNN, adOpenKeyset, adLockReadOnly
'        If Not VBA.IsNull(vRsTmp![Total Fees]) Then vIndex(&H0) = VBA.FormatNumber(vRsTmp![Total Fees], &H2)
'        vRsTmp.Close 'Close the opened object and any dependent objects
'
'        'Get the total School Fees paid for the previous Year
'        vRsTmp.Open "SELECT SUM([Amount]) AS [Total Fees Paid] FROM [Qry_Payments] WHERE [Student ID] = " & VBA.Val(txtStudentName.Tag) & " AND [Fee Structure Year] < " & VBA.Year(dtYear.Value), vAdoCNN, adOpenKeyset, adLockReadOnly
'        If Not VBA.IsNull(vRsTmp![Total Fees Paid]) Then vIndex(&H0) = vIndex(&H0) - VBA.FormatNumber(vRsTmp![Total Fees Paid], &H2)
'        vRsTmp.Close 'Close the opened object and any dependent objects
        
        ![Balance Carried Forward] = vIndex(&H0)
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![Date Last Modified] = VBA.Now
        ![Student Class Status] = cboStudentClassStatus.Text
        ![User ID] = User.User_ID
'        ![School ID] = VBA.IIf(School.ID = &H0, &H1, School.ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        Fra_StudentClasses.Tag = ![Student Class ID]
        
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
    
    Exit Sub 'Quit this Procedure
    
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [Student Class ID], YEAR([Date Assigned]) AS [Year], ([Registered Name]) AS [Student Name], [Class] FROM [Qry_StudentClasses] WHERE NOT ([Previous Class] IS NULL OR TRIM([Previous Class])='') AND [Date Assigned] >= " & SoftwareSetting.Min_Year & " ORDER BY [Date Assigned] DESC, [Class Level] DESC, [Stream Name] ASC, [Student Name] ASC;", (Fra_StudentClasses.Tag), , "1", , , "Student Class Allocations", &H3, "[Qry_StudentClasses]; WHERE [Student Class ID] = $;3", , 6000)
    
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
    
    Exit Sub 'Quit this Procedure
    
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
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_StudentClasses.Enabled, "Student Sports")
End Sub
