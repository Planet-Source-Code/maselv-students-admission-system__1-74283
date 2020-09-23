VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_Dormitories 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Dormitories"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5175
   Icon            =   "Frm_Dormitories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3600
      Picture         =   "Frm_Dormitories.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Move First"
      Top             =   1320
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
      Left            =   3960
      Picture         =   "Frm_Dormitories.frx":06CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Move Previous"
      Top             =   1320
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
      Left            =   4320
      Picture         =   "Frm_Dormitories.frx":0A0E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Move Next"
      Top             =   1320
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
      Left            =   4680
      Picture         =   "Frm_Dormitories.frx":0D50
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Move Last"
      Top             =   1320
      Width           =   375
   End
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
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
      Picture         =   "Frm_Dormitories.frx":1092
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
   Begin VB.Frame Fra_Dormitories 
      BackColor       =   &H00CFE1E2&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   4935
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
         Left            =   3480
         TabIndex        =   4
         Tag             =   "Y"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtCapacity 
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
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   465
         Width           =   420
      End
      Begin VB.TextBox txtDormitoryName 
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
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.UpDown udCapacity 
         Height          =   315
         Left            =   3060
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   465
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         BuddyControl    =   "txtCapacity"
         BuddyDispid     =   196613
         OrigLeft        =   3000
         OrigTop         =   465
         OrigRight       =   3255
         OrigBottom      =   780
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCapacity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity:"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblStreamName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dormitory Name:"
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
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
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
      TabIndex        =   12
      Top             =   1410
      Width           =   1395
   End
   Begin VB.Image ImgHeader 
      Height          =   360
      Left            =   0
      Picture         =   "Frm_Dormitories.frx":142C
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   5295
   End
   Begin VB.Image ImgFooter 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Dormitories.frx":1CCC
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   1320
      Width           =   5295
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
Attribute VB_Name = "Frm_Dormitories"
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

'To simplify administration of user accounts which have similar resource needs, the administrator can
'categorize the user accounts into Relationships, which makes granting access rights and resource permissions
'easier. Instead of performing many individual actions to grant certain rights or permissions, the
'administrator can perform a single action that gives a Relationship that right or permission to all the present
'and future members of that Relationship

Option Explicit

Public IsNewRecord As Boolean
Public FrmDefinitions$, SetPrivileges$

Private myRecIndex&, iSetNo&
Private myTable$, myTablePryKey$
'Private myTableFixedFldName() As String
Private myRecordIDs() As String
Private myRecDisplayON, IsLoading, iUseSchoolType As Boolean

Public Function ClearEntries() As Boolean
On Local Error GoTo Handle_ClearEntries_Error
    
    Dim MousePointerState%
    Dim myRecDisplayState As Boolean
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayState = myRecDisplayON
    
    txtDormitoryName.Tag = VBA.vbNullString
    txtDormitoryName.Text = VBA.vbNullString
    SetPrivileges = def_Privileges
    ShpBttnBriefNotes.TagExtra = VBA.vbNullString
    
    VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
    
    myRecDisplayON = True: chkDiscontinued.Value = vbUnchecked: myRecDisplayON = myRecDisplayState
    
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
    
    Fra_Dormitories.Enabled = Not State
    
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
    Static LastStructureID&
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayON = True 'Denote that Record display process is in progress
    
    VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Tbl_Dormitories] WHERE [Dormitory ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            txtDormitoryName.Tag = ![Dormitory ID] 'Assign the Record's primary key value
            If Not VBA.IsNull(![Dormitory Name]) Then txtDormitoryName.Text = ![Dormitory Name]
            If Not VBA.IsNull(![Capacity]) Then txtCapacity.Text = ![Capacity]
            If Not VBA.IsNull(![Brief Notes]) Then ShpBttnBriefNotes.TagExtra = ![Brief Notes]
            chkDiscontinued.Value = VBA.IIf(![Discontinued], vbChecked, vbUnchecked)
            
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
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, iSetNo, &H4) Then myRecDisplayON = True: chkDiscontinued.Value = VBA.IIf(chkDiscontinued.Value = vbUnchecked, vbChecked, vbUnchecked): myRecDisplayON = False: Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If chkDiscontinued.Value = vbChecked Then If vMsgBox("Ticking this option will disable the Dormitory and will not be available in other Modules. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then myRecDisplayON = True: chkDiscontinued.Value = vbUnchecked: myRecDisplayON = False
    
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
    myRecIndex = NavigateToRec(Me, "SELECT * FROM [Tbl_Dormitories] ORDER BY [Dormitory Name] ASC", "Dormitory ID", Index, myRecIndex)
    
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
    
    Select Case myTable
        
        Case "Tbl_Dormitories": iSetNo = &HB
        
    End Select
    
    'Yield execution so that the operating system can process other events
    VBA.DoEvents
    
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
        If vEditRecordID <> VBA.vbNullString Then
            
            'Return a zero-based, one-dimensional array containing a specified number of substrings
            vArrayList = VBA.Split(vEditRecordID, "|")
            
            'Retrieve the Record with the specified ID
            .Filter = "[Dormitory ID] = " & vArrayList(&H0)
            
            Dim nHasRecords As Boolean
            
            nHasRecords = Not (.BOF And .EOF)
            
            'If the record Exists then Call Procedure in this Form to display it
            If nHasRecords Then DisplayRecord VBA.CLng(VBA.Val(![Dormitory ID]))
            
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
    
    Me.Caption = App.Title & " : Dormitories"
    myTable = "Tbl_Dormitories"
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub MnuDelete_Click()
On Local Error GoTo Handle_MnuDelete_Click_Error
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, iSetNo, &H3) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'If the User has not selected any existing Record then...
    If txtDormitoryName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        GoTo Exit_MnuDelete_Click 'Branch to the specified Label
        
    End If 'Close respective IF..THEN block statement
    
    'Confirm deletion
    If vMsgBox("Are you sure you want to DELETE the displayed Dormitory?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    Dim iDependencyTitle As String
    Dim iDependency As Boolean
    Dim iDependencyArray() As String
    Dim iDependencyArrayTmp() As String
    
    iDependencyArray = VBA.Split("[Tbl_StudentDormitories]", "|")
    
    'For each entry...
    For vIndex(&H0) = &H0 To UBound(iDependencyArray) Step &H1
        
        iDependencyArrayTmp = VBA.Split(iDependencyArray(vIndex(&H0)), ";")
        
        'Check if there are Records in other tables depending on the displayed Record
        iDependency = CheckForRecordDependants(Me, iDependencyArrayTmp(&H0) & " WHERE [Dormitory ID] = " & VBA.Val(txtDormitoryName.Tag))
        
        'If the Record is depended upon then Quit this FOR..LOOP block statement
        If iDependency Then Exit For
        
    Next vIndex(&H0) 'Move to the next entry
    
    'If the Records exist then...
    If iDependency Then
        
        'Warn User
        vMsgBox "The displayed Dormitory has other Records depending on it {" & iDependencyArrayTmp(&H1) & "}. Delete operation aborted", vbExclamation, App.Title & " : Operation Aborted", Me
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM " & myTable & " WHERE [Dormitory ID] = " & VBA.Val(txtDormitoryName.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
    
    'Denote that the database table has been altered
    vDatabaseAltered = True
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    'Get the total number of Records already saved
    lblRecords.Tag = VBA.Val(lblRecords.Tag) - &H1
    lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "The Dormitory has successfully been deleted.", vbInformation, App.Title & " : Delete", Me
    
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
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, iSetNo, &H2) Then Exit Sub
    
    'If the User has not selected any existing Record then...
    If txtDormitoryName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        If Fra_Dormitories.Enabled Then txtDormitoryName.SetFocus 'Move the focus to the specified control
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
    
    txtDormitoryName.SetFocus 'Move focus to Name textbox
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
End Sub

Private Sub MnuNew_Click()
    
    'Check If the User has insufficient privileges to perform this Operation
    If Not ValidUserAccess(Me, iSetNo, &H1) Then Exit Sub
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    Call LockEntries(False) 'Call Procedure in this Form to UnLock Input Controls
    
    IsNewRecord = True 'Denote that the displayed Record does not exist in the database
    
    txtDormitoryName.SetFocus 'Move focus to Name textbox
    
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
    
    'If the User has not entered the name then...
    If txtDormitoryName.Text = VBA.vbNullString Then
        
        'Request User to make the entry
        vMsgBox "Please enter the Dormitory name before proceeding with saving", vbExclamation, App.Title & " : Blank Entry", Me
        txtDormitoryName.SetFocus 'Move the focus to the specified control
        GoTo Exit_MnuSave_Click 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'Start the saving process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Format entry appropriately
    txtDormitoryName.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtDormitoryName.Text, "  ", " ")))
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        If VBA.Val(txtCapacity.Text) <> &H0 Then
            
            .Open "SELECT * FROM [Tbl_StudentDormitories] WHERE [Dormitory ID] = " & VBA.Val(txtDormitoryName.Tag), vAdoCNN, adOpenKeyset, adLockPessimistic
            
            If .RecordCount > VBA.Val(txtCapacity.Text) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                vMsgBox "The entered Dormitory Capacity is lower than the total number of Students already assigned to it {" & .RecordCount & "}. Please set a capacity of equal or higher figure, or set to zero for unlimited capacity.", vbExclamation, App.Title & " : Invalid Capacity Limit", Me
                txtCapacity.SetFocus 'Move the focus to the specified control
                txtCapacity.SelStart = &H0: txtCapacity.SelLength = VBA.Len(txtCapacity.Text) 'Highlight contents
                
                .Close 'Close the opened object and any dependent objects
                GoTo Exit_MnuSave_Click 'Quit this saving procedure
                
            End If 'Close respective IF..THEN block statement
            
            .Close 'Close the opened object and any dependent objects
            
        End If 'Close respective IF..THEN block statement
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'Check if the entered Dormitory name already exists in the database
        .Filter = "[Dormitory ID] <> " & VBA.Val(txtDormitoryName.Tag) & " AND [Dormitory Name] = '" & VBA.Replace(txtDormitoryName.Text, "'", "''") & "'"
        
        'If the name already exists then...
        If Not (.BOF And .EOF) Then
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            vMsgBox "The entered Dormitory name had already been saved for the selected Class. Please enter a different name.", vbExclamation, App.Title & " : Duplicate Entry", Me
            txtDormitoryName.SetFocus 'Move the focus to the specified control
            GoTo Exit_MnuSave_Click 'Quit this saving procedure
            
        End If 'Close respective IF..THEN block statement
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[Dormitory ID] = " & txtDormitoryName.Tag: .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        ![Dormitory Name] = txtDormitoryName.Text
        ![Capacity] = VBA.Val(txtCapacity.Text)
        ![Discontinued] = chkDiscontinued.Value
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        txtDormitoryName.Tag = ![Dormitory ID]
        
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [Dormitory ID], ([Dormitory Name]) AS [Dormitory] FROM [Tbl_Dormitories] ORDER BY [Dormitory Name] ASC", txtDormitoryName.Tag, , "1", , , "Dormitories")
    
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
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_Dormitories.Enabled, "Dormitory")
End Sub

Private Sub txtCapacity_KeyPress(KeyAscii As Integer)
    'Discard non-numeric entries
    KeyAscii = VBA.IIf((((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126))) And KeyAscii <> vbKeyReturn, KeyAscii = Empty, KeyAscii)
End Sub
