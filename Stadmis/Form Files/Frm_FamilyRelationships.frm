VERSION 5.00
Begin VB.Form Frm_FamilyRelationships 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Family Categories"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4935
   Icon            =   "Frm_FamilyRelationships.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.ShapeButton ShpBttnMove 
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   ""
      Picture         =   "Frm_FamilyRelationships.frx":09EA
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
   Begin Stadmis.ShapeButton ShpBttnMove 
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   6
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   ""
      Picture         =   "Frm_FamilyRelationships.frx":0F84
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
   Begin Stadmis.ShapeButton ShpBttnMove 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   ""
      Picture         =   "Frm_FamilyRelationships.frx":131E
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
   Begin Stadmis.ShapeButton ShpBttnBriefNotes 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
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
      Picture         =   "Frm_FamilyRelationships.frx":18B8
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
      Left            =   4440
      Picture         =   "Frm_FamilyRelationships.frx":1C52
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move Last"
      Top             =   2520
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
      Left            =   4080
      Picture         =   "Frm_FamilyRelationships.frx":1F94
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Move Next"
      Top             =   2520
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
      Left            =   3720
      Picture         =   "Frm_FamilyRelationships.frx":22D6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Move Previous"
      Top             =   2520
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
      Left            =   3360
      Picture         =   "Frm_FamilyRelationships.frx":2618
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Move First"
      Top             =   2520
      Width           =   375
   End
   Begin VB.ListBox LstItems 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "New Records appear in the list only after saving"
      Top             =   1080
      Width           =   4092
   End
   Begin VB.Frame Fra_FamilyRelationship 
      BackColor       =   &H00CFE1E2&
      Height          =   2172
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtName 
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
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   3135
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
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label lblHierarchicalLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hierarchical Level:"
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
         TabIndex        =   12
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Family Relationship Name:"
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
         Width           =   1875
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
      TabIndex        =   14
      Top             =   2610
      Width           =   1395
   End
   Begin VB.Image ImgHeader 
      Height          =   255
      Left            =   0
      Picture         =   "Frm_FamilyRelationships.frx":295A
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   5055
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_FamilyRelationships.frx":3150
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5055
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
Attribute VB_Name = "Frm_FamilyRelationships"
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
Private myTableFixedFldName() As String
Private myRecordIDs() As String
Private myRecDisplayON, IsLoading, iUseSchoolType, ShowHierarchySaveMsg As Boolean

Public Function ClearEntries() As Boolean
On Local Error GoTo Handle_ClearEntries_Error
    
    Dim MousePointerState%
    Dim myRecDisplayState As Boolean
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    myRecDisplayState = myRecDisplayON
    
    txtName.Tag = VBA.vbNullString
    txtName.Text = VBA.vbNullString
    SetPrivileges = def_Privileges
    lblHierarchicalLevel.Tag = VBA.vbNullString
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
    
    Fra_FamilyRelationship.Enabled = Not State
    
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
        
        .Open "SELECT * FROM " & myTableFixedFldName(&H4) & " WHERE [" & myTableFixedFldName(&H0) & " ID] = " & vRecordID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            'Display field values in their respective controls
            
            txtName.Tag = vRs(myTableFixedFldName(&H0) & " ID") 'Assign the Record's primary key value
            
            If Not VBA.IsNull(vRs(myTableFixedFldName(&H0) & " Name")) Then txtName.Text = vRs(myTableFixedFldName(&H0) & " Name")
            If Not VBA.IsNull(![Hierarchical Level]) Then lblHierarchicalLevel.Tag = ![Hierarchical Level]
            If Not VBA.IsNull(![Brief Notes]) Then ShpBttnBriefNotes.TagExtra = ![Brief Notes]
            chkDiscontinued.Value = VBA.IIf(![Discontinued], vbChecked, vbUnchecked)
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
        'Fill List items
        Call FillItemList
        
        'Highlight displayed item in the ListBox
        LstItems.ListIndex = SearchCboLst(LstItems, txtName.Text)
        
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

Private Function FillItemList() As Boolean
On Local Error GoTo Handle_FillItemList_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the existing Records in the database
        .Open "SELECT * FROM " & myTableFixedFldName(&H4) & VBA.IIf(iUseSchoolType, " WHERE [School ID] = " & School.ID & " OR [School ID] IS NULL OR [School ID] = 0", VBA.vbNullString) & " ORDER BY [Hierarchical Level] ASC, [" & myTableFixedFldName(&H0) & " Name] ASC", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'Clear the contents of the specified ListBox
        LstItems.Clear
        
        'Get the total number of Records already saved
        lblRecords.Tag = .RecordCount
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        If .RecordCount > &H0 Then ReDim myRecordIDs(.RecordCount - &H1) As String
        
        'Repeat to the last Record
        For vIndex(&H0) = &H0 To .RecordCount - &H1 Step &H1
            
            'Assign the field values
            If Not VBA.IsNull(vRs(myTableFixedFldName(&H0) & " ID")) Then myRecordIDs(vIndex(&H0)) = vRs(myTableFixedFldName(&H0) & " ID")
            If Not VBA.IsNull(vRs(myTableFixedFldName(&H0) & " Name")) Then LstItems.AddItem vRs(myTableFixedFldName(&H0) & " Name")
            
            .MoveNext 'Move to the next Record
            
        Next vIndex(&H0) 'Loop through all Records
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
    FillItemList = True 'Denote that the ListBox has successfully been filled with data
    
Exit_FillItemList:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_FillItemList_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Filling Item List - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_FillItemList
    
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
    
    If chkDiscontinued.Value = vbChecked Then If vMsgBox("Ticking this option will disable the " & myTableFixedFldName(&H2) & " and will not be available in other Modules. Proceed?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then myRecDisplayON = True: chkDiscontinued.Value = vbUnchecked: myRecDisplayON = False
    
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
    myRecIndex = NavigateToRec(Me, "SELECT * FROM " & myTableFixedFldName(&H4) & VBA.IIf(iUseSchoolType, " WHERE [School Type] = " & School.Type, VBA.vbNullString) & " ORDER BY [Hierarchical Level] ASC, [" & myTableFixedFldName(&H0) & " Name] ASC", myTableFixedFldName(&H0) & " ID", Index, myRecIndex)
    
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
    ShowHierarchySaveMsg = True
    
    Select Case myTableFixedFldName(&H0)
        
        Case "Class": iSetNo = &H7
        Case "Relationship": iSetNo = &H5
        
    End Select
    
    'Yield execution so that the operating system can process other events
    VBA.DoEvents
    
    LockEntries True 'Call Procedure in this Form to Lock Input Controls
    ClearEntries 'Call Procedure in this Form to clear all entries in Input Boxes
    
    Call FillItemList 'Call Procedure in this Form to display saved Records
    
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
            .Filter = "[" & myTableFixedFldName(&H0) & " ID] = " & vArrayList(&H0)
            
            Dim nHasRecords As Boolean
            
            nHasRecords = Not (.BOF And .EOF)
            
            'If the record Exists then Call Procedure in this Form to display it
            If nHasRecords Then DisplayRecord VBA.CLng(VBA.Val(vRs(myTableFixedFldName(&H0) & " ID")))
            
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
    
    If FrmDefinitions = VBA.vbNullString Then FrmDefinitions = "Relationship:Relationships:Relationship:Relationships:Tbl_Relationships:Tbl_StudentRelatives:Students"
    
    myTableFixedFldName = VBA.Split(FrmDefinitions, ":")
    
    myTable = "Tbl_" & VBA.Replace(myTableFixedFldName(&H1), " ", VBA.vbNullString)
    
    Me.Caption = App.Title & " : " & myTableFixedFldName(&H3)
    lblName.Caption = myTableFixedFldName(&H2) & ":"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub LstItems_Click()
On Local Error GoTo Handle_LstItems_Click_Error
    
    'If their are no list items to select then quit this Procedure
    If LstItems.ListCount = &H0 Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'If the selected list item...
    Select Case LstItems.ListIndex
        
        Case &H0: 'Is the first one in the List then...
            
            ShpBttnMove(&H0).Enabled = False 'Disable Move Up
            ShpBttnMove(&H2).Enabled = (LstItems.ListCount > &H1) 'Enable Move Down
            
        Case LstItems.ListCount - &H1: 'Is the Last one in the List then...
            
            ShpBttnMove(&H0).Enabled = True 'Enable Move Up
            ShpBttnMove(&H2).Enabled = False 'Disable Move Down
            
        Case Else
            
            ShpBttnMove(&H0).Enabled = True 'Enable Move Up
            ShpBttnMove(&H2).Enabled = True 'Enable Move Down
            
    End Select 'Close SELECT..CASE block statement
    
Exit_LstItems_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_LstItems_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Item - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_LstItems_Click
    
End Sub

Private Sub LstItems_DblClick()
On Local Error GoTo Handle_LstItems_DblClick_Error
    
    'If a record is being modified and the user chooses to proceed with modifying it then Quit this Procedure
    If ProceedEditting(Me) = True Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Display the Record of the double-clicked item
    Call DisplayRecord(VBA.CLng(myRecordIDs(LstItems.ListIndex)))
    
Exit_LstItems_DblClick:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_LstItems_DblClick_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Displaying Item - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_LstItems_DblClick
    
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
    If txtName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        GoTo Exit_MnuDelete_Click 'Branch to the specified Label
        
    End If 'Close respective IF..THEN block statement
    
    'Confirm deletion
    If vMsgBox("Are you sure you want to DELETE the displayed " & myTableFixedFldName(&H2) & "'s Record?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then
        
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
        iDependency = CheckForRecordDependants(Me, iDependencyArrayTmp(&H0) & " WHERE [" & myTableFixedFldName(&H0) & " ID] = " & txtName.Tag)
        
        'If the Record is depended upon then Quit this FOR..LOOP block statement
        If iDependency Then Exit For
        
    Next vIndex(&H0) 'Move to the next entry
    
    'If the Records exist then...
    If iDependency Then
        
        'Warn User
        vMsgBox "The displayed " & myTableFixedFldName(&H2) & " has other Records depending on it {" & iDependencyArrayTmp(&H1) & "}. Delete operation aborted", vbExclamation, App.Title & " : Operation Aborted", Me
        
        'Branch to the specified Label
        GoTo Exit_MnuDelete_Click
        
    End If 'Close respective IF..THEN block statement
    
    'Start the Delete process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    
    'Delete the Record with the specified primary key value
    vRs.Open "DELETE * FROM " & myTable & " WHERE [" & myTableFixedFldName(&H0) & " ID] = " & txtName.Tag, vAdoCNN, adOpenKeyset, adLockPessimistic
    
    'Denote that the database table has been altered
    vDatabaseAltered = True
    
    Call ClearEntries 'Call Procedure in this Form to erase all entries in the Input Controls
    
    Call FillItemList 'Call Procedure in this Form to display saved Records
    
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
    If txtName.Tag = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please display a Record.", vbExclamation, App.Title & " : No Record Displayed", Me
        If Fra_FamilyRelationship.Enabled Then txtName.SetFocus 'Move the focus to the specified control
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
    
    txtName.SetFocus 'Move focus to Name textbox
    
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
    
    txtName.SetFocus 'Move focus to Name textbox
    
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
    If txtName.Text = VBA.vbNullString Then
        
        'Request User to make the entry
        vMsgBox "Please enter the name before proceeding with saving", vbExclamation, App.Title & " : Blank Entry", Me
        txtName.SetFocus 'Move the focus to the specified control
        GoTo Exit_MnuSave_Click 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'Start the saving process
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Format entry appropriately
    txtName.Text = CapAllWords(VBA.Trim$(VBA.Replace(txtName.Text, "  ", " ")))
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve all the Records from the specified table
        .Open "SELECT * FROM " & myTable, vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'Check if the entered name already exists in the database
        .Filter = "[" & myTableFixedFldName(&H0) & " ID] <> " & VBA.Val(txtName.Tag) & " AND [" & myTableFixedFldName(&H0) & " Name] = '" & VBA.Replace(txtName.Text, "'", "''") & "'"
        
        'If the name already exists then...
        If Not (.BOF And .EOF) Then
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            vMsgBox "The entered name had already been saved. Please enter a different name.", vbExclamation, App.Title & " : Duplicate Entry", Me
            txtName.SetFocus 'Move the focus to the specified control
            GoTo Exit_MnuSave_Click 'Quit this saving procedure
            
        End If 'Close respective IF..THEN block statement
        
        'If the displayed Record exists in the database then update it else add it as a new Record
        If Not IsNewRecord Then .Filter = "[" & myTableFixedFldName(&H0) & " ID] = " & txtName.Tag: .Update Else .AddNew
        
        'Assign entries to their respective database fields
        
        vRs(myTableFixedFldName(&H0) & " Name") = txtName.Text
        If IsNewRecord Then ![Hierarchical Level] = LstItems.ListCount + &H1
        If iUseSchoolType Then ![School Type] = School.Type
        ![Discontinued] = chkDiscontinued.Value
        ![Brief Notes] = ShpBttnBriefNotes.TagExtra
        ![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        'Assign the current Primary Key value of the saved Record
        txtName.Tag = vRs(myTableFixedFldName(&H0) & " ID")
        
        .Close 'Close the opened object and any dependent objects
        
        'Denote that the database table has been altered
        vDatabaseAltered = True
        
        'Get the total number of Records already saved
        lblRecords.Tag = VBA.Val(lblRecords.Tag) + VBA.IIf(IsNewRecord, &H1, &H0)
        lblRecords.Caption = "Total Records: " & VBA.Val(lblRecords.Tag)
        
        VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
        LockEntries True 'Call Procedure in this Form to Lock Input Controls
        
        'If a new Record has been added then add the Record to the List
        If IsNewRecord Then LstItems.AddItem txtName.Text
        
    End With 'Close the WITH block statements
    
    'Call Procedure in this Form to save the Hierarchy as displayed
    If Not IsNewRecord Then myRecDisplayON = True: ShowHierarchySaveMsg = False: Call ShpBttnMove_Click(&H1): ShowHierarchySaveMsg = True: myRecDisplayON = False
    
    Call FillItemList 'Call Procedure in this Form to display saved Records
    
    'Select the entered Item
    If LstItems.ListCount > &H0 Then LstItems.Selected(SearchCboLst(LstItems, txtName.Text)) = True
    
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
    vBuffer(&H0) = PickDetails(Me, "SELECT [" & myTableFixedFldName(&H0) & " ID], [" & myTableFixedFldName(&H0) & " Name] FROM " & myTableFixedFldName(&H4) & VBA.IIf(iUseSchoolType, " WHERE [School Type] = " & School.Type, VBA.vbNullString) & " ORDER BY [Hierarchical Level] ASC, [" & myTableFixedFldName(&H0) & " Name] ASC", txtName.Tag, , "1", , , myTableFixedFldName(&H3))
    
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
    Call OpenNotes(ShpBttnBriefNotes, Not Fra_FamilyRelationship.Enabled, myTableFixedFldName(&H2))
End Sub

Private Sub ShpBttnMove_Click(Index As Integer)
On Local Error GoTo Handle_ShpBttnMove_Click_Error
    
    'If their are no list items to select then quit this Procedure
    If LstItems.ListCount = &H0 Then Exit Sub
    
    If Not ValidUserAccess(Me, iSetNo, &H6) Then Exit Sub
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim iOldListIndex&, iNewListIndex&
    
    iOldListIndex = LstItems.ListIndex
    
    'If the selected button is...
    Select Case Index
        
        Case &H0: 'Move Up then...
            
            iNewListIndex = LstItems.ListIndex - &H1
            LstItems.AddItem LstItems.List(iOldListIndex), iNewListIndex
            LstItems.RemoveItem iOldListIndex + &H1
            LstItems.ListIndex = iNewListIndex
            ShpBttnMove(&H1).Enabled = True 'Enable saving Hierarchy
            
        Case &H2: 'Move Down then...
            
            iNewListIndex = LstItems.ListIndex + &H1 'Get the new index value
            LstItems.AddItem LstItems.List(iOldListIndex), iNewListIndex + &H1
            LstItems.RemoveItem iOldListIndex
            LstItems.ListIndex = iNewListIndex
            ShpBttnMove(&H1).Enabled = True 'Enable saving Hierarchy
            
        Case Else 'Save Hierarchy then..
            
            If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
            Set vRs = New ADODB.Recordset
            With vRs 'Execute a series of statements on vRs recordset
                
                'For each item in the list...
                For vIndex(&H0) = &H1 To LstItems.ListCount Step &H1
                    
                    'Update the Hierarchy of the list items
                    .Open "UPDATE " & myTable & " SET [Hierarchical Level] = " & vIndex(&H0) & " WHERE [" & myTableFixedFldName(&H0) & " Name] = '" & VBA.Replace(LstItems.List(vIndex(&H0) - &H1), "'", "''") & "';", vAdoCNN, adOpenKeyset, adLockPessimistic
                    
                Next vIndex(&H0) 'Move to the next item in the list
                
            End With 'Close the WITH block statements
            
            ShpBttnMove(&H1).Enabled = False 'Disable saving Hierarchy
            
            'If saving then...
            If ShowHierarchySaveMsg Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox myTableFixedFldName(&H3) & " Hierarchy has successfully been saved", vbInformation, App.Title & " : Hierarchical Level", Me
                
            End If 'Close respective IF..THEN block statement
            
            'Denote that the database table has been altered
            vDatabaseAltered = True
            
    End Select 'Close SELECT..CASE block statement
    
Exit_ShpBttnMove_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_ShpBttnMove_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnMove_Click
    
End Sub

