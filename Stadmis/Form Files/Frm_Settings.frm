VERSION 5.00
Object = "{60D73138-4A06-4DA5-838D-E17FF732D00B}#1.0#0"; "prjXTab.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Software Settings"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "Frm_Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   120
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Wav Sounds (*.wav)|*.wav"
   End
   Begin VB.Frame Fra_Settings 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      Begin prjXTab.XTab XTab 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6800
         TabCount        =   2
         TabCaption(0)   =   "Software Sounds"
         TabContCtrlCnt(0)=   4
         Tab(0)ContCtrlCap(1)=   "Fra_Setting0"
         Tab(0)ContCtrlCap(2)=   "ShpBttnApply"
         Tab(0)ContCtrlCap(3)=   "ShpBttnLoadDefaults"
         Tab(0)ContCtrlCap(4)=   "chkEnableSoftwareSounds"
         TabCaption(1)   =   "Other Settings"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "Fra_Setting1"
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
         Begin VB.Frame Fra_Setting 
            BackColor       =   &H00FFFFFF&
            Height          =   2775
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   6255
            Begin VB.CommandButton cmdPlaySound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   3
               Left            =   5760
               Picture         =   "Frm_Settings.frx":038A
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Play selected sound"
               Top             =   2280
               Width           =   345
            End
            Begin VB.CommandButton cmdBrowseSound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   3
               Left            =   5400
               Picture         =   "Frm_Settings.frx":0914
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Browse for Sound"
               Top             =   2280
               Width           =   345
            End
            Begin VB.TextBox txtSound 
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
               Index           =   3
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   2280
               Width           =   5295
            End
            Begin VB.CommandButton cmdPlaySound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   2
               Left            =   5760
               Picture         =   "Frm_Settings.frx":0E9E
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Play selected sound"
               Top             =   1680
               Width           =   345
            End
            Begin VB.CommandButton cmdBrowseSound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   2
               Left            =   5400
               Picture         =   "Frm_Settings.frx":1428
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Browse for Sound"
               Top             =   1680
               Width           =   345
            End
            Begin VB.TextBox txtSound 
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
               Index           =   2
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   1680
               Width           =   5295
            End
            Begin VB.CommandButton cmdPlaySound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   1
               Left            =   5760
               Picture         =   "Frm_Settings.frx":19B2
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Play selected sound"
               Top             =   1080
               Width           =   345
            End
            Begin VB.CommandButton cmdBrowseSound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   1
               Left            =   5400
               Picture         =   "Frm_Settings.frx":1F3C
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Browse for Sound"
               Top             =   1080
               Width           =   345
            End
            Begin VB.TextBox txtSound 
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
               Index           =   1
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   1080
               Width           =   5295
            End
            Begin VB.CommandButton cmdPlaySound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   0
               Left            =   5760
               Picture         =   "Frm_Settings.frx":24C6
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Play selected sound"
               Top             =   480
               Width           =   345
            End
            Begin VB.CommandButton cmdBrowseSound 
               BackColor       =   &H00D8E9EC&
               Height          =   315
               Index           =   0
               Left            =   5400
               Picture         =   "Frm_Settings.frx":2A50
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Browse for Sound"
               Top             =   480
               Width           =   345
            End
            Begin VB.TextBox txtSound 
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
               Index           =   0
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   480
               Width           =   5295
            End
            Begin VB.Label lblSound 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Critical Sound:"
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
               Index           =   3
               Left            =   120
               TabIndex        =   19
               Top             =   2040
               Width           =   1035
            End
            Begin VB.Label lblSound 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Warning Sound:"
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
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   1440
               Width           =   1155
            End
            Begin VB.Label lblSound 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Question Sound:"
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
               Index           =   1
               Left            =   120
               TabIndex        =   11
               Top             =   840
               Width           =   1200
            End
            Begin VB.Label lblSound 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Information Sound:"
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
               Index           =   0
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   1395
            End
         End
         Begin Stadmis.ShapeButton ShpBttnApply 
            Height          =   375
            Left            =   5400
            TabIndex        =   23
            Tag             =   "AutoSizer:Y"
            Top             =   3360
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
            Caption         =   "      Apply"
            AccessKey       =   "A"
            Picture         =   "Frm_Settings.frx":2FDA
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
         Begin Stadmis.ShapeButton ShpBttnLoadDefaults 
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Tag             =   "AutoSizer:Y"
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
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
            Caption         =   "       Load Defaults"
            AccessKey       =   "D"
            Picture         =   "Frm_Settings.frx":3574
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
         Begin VB.CheckBox chkEnableSoftwareSounds 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Enable Software Sounds"
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
            TabIndex        =   8
            Tag             =   "Y"
            Top             =   360
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.Frame Fra_Setting 
            BackColor       =   &H00FFFFFF&
            Height          =   3375
            Index           =   1
            Left            =   -74880
            TabIndex        =   4
            Top             =   360
            Width           =   6255
         End
      End
   End
   Begin Stadmis.ShapeButton ShpBttnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Tag             =   "AutoSizer:XY"
      Top             =   4800
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
      Caption         =   "      Close"
      AccessKey       =   "C"
      Picture         =   "Frm_Settings.frx":390E
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
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software Settings"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2280
   End
   Begin VB.Image ImgHeader 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Settings.frx":3CA8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7095
   End
   Begin VB.Image ImgFooter 
      Height          =   615
      Left            =   0
      Picture         =   "Frm_Settings.frx":449E
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   7095
   End
End
Attribute VB_Name = "Frm_Settings"
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

Private strStatus$

Private Sub chkEnableSoftwareSounds_Click()
    Fra_Setting(&H0).Enabled = (chkEnableSoftwareSounds.Value = vbChecked)
End Sub

Private Sub cmdBrowseSound_Click(Index As Integer)
On Local Error GoTo Handle_cmdBrowseSound_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
BrowseSound:
    
    'Execute a series of statements on the specified object
    With Dlg
        
        .FLAGS = &H4 'Hide Read-Only checkbox
        
        'Generate error when the user chooses the Cancel button.
        .CancelError = True
        
        'Set Mouse pointer to indicate end of this process or operation to finish.
        Screen.MousePointer = vbDefault
        
        .ShowOpen 'Display the CommonDialog control's Open dialog box.
        
        'Set Mouse pointer to indicate beginning of process or operation
        Screen.MousePointer = vbHourglass
        
        'If a Sound has been Selected
        If VBA.LenB(VBA.Trim$(.FileName)) <> &H0 Then txtSound(Index).Text = VBA.Trim$(.FileName)
        
    End With 'End the WITH statement
    
Exit_cmdBrowseSound_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub procedure
    
Handle_cmdBrowseSound_Click_Error:
    
    'If it is a Cancel error then resume execution at the specified Label
    If Err.Number = 32755 Then Resume Exit_cmdBrowseSound_Click
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Selecting Sound - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_cmdBrowseSound_Click
    
End Sub

Private Sub cmdPlaySound_Click(Index As Integer)
On Local Error GoTo Handle_cmdPlaySound_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Call PlaySound(Me, txtSound(Index).Text)
    
Exit_cmdPlaySound_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub procedure
    
Handle_cmdPlaySound_Click_Error:
    
    'If it is a Cancel error then resume execution at the specified Label
    If Err.Number = 32755 Then Resume Exit_cmdPlaySound_Click
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Playing Sound - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_cmdPlaySound_Click
    
End Sub

Private Sub Form_Load()
On Local Error GoTo Handle_Form_Load_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Me.Caption = App.Title & " : Software Settings"
    
    strStatus = "Loading Form"
    
    Call ConnectDB
'    txtSound(&H4).Text = vAdoCNN.Properties("Data Source")
    
    Dim iFldr
    
    If vFso.FolderExists(App.Path & "\Tools\Media\Default Media") Then
        
        Set iFldr = vFso.GetFolder(App.Path & "\Tools\Media\Default Media")
        iFldr.Attributes = &H7
        
    End If
    
    If vFso.FolderExists(App.Path & "\Application Data") Then
        
        Set iFldr = vFso.GetFolder(App.Path & "\Application Data")
        iFldr.Attributes = &H7
        
    End If
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Tbl_Setup]", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            chkEnableSoftwareSounds.Value = VBA.IIf(![System Sound], vbChecked, vbUnchecked)
            
            If Not VBA.IsNull(![System Sounds]) Then
                
                chkEnableSoftwareSounds.Tag = ![System Sounds]
                
                vArrayList = VBA.Split(chkEnableSoftwareSounds.Tag, "|")
                
                For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
                    txtSound(vIndex(&H0)).Text = vArrayList(vIndex(&H0))
                Next vIndex(&H0)
                
            Else
                
                Call ShpBttnLoadDefaults_Click
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_Form_Load:
    
    strStatus = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub procedure
    
Handle_Form_Load_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Loading Form - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_Form_Load
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub ShpBttnApply_Click()
On Local Error GoTo Handle_ShpBttnApply_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Tbl_Setup]", vAdoCNN, adOpenKeyset, adLockPessimistic
        
        Dim IsNewRecord As Boolean
        
        IsNewRecord = (.BOF And .EOF)
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then .Update Else .AddNew
        
        chkEnableSoftwareSounds.Value = VBA.IIf(![System Sound], vbChecked, vbUnchecked)
        
        For vIndex(&H0) = &H0 To txtSound.UBound Step &H1
            vBuffer(&H0) = vBuffer(&H0) & "|" & txtSound(vIndex(&H0)).Text
        Next vIndex(&H0)
        
        If VBA.Left$(vBuffer(&H0), &H1) = "|" Then vBuffer(&H0) = VBA.Right$(vBuffer(&H0), VBA.Len(vBuffer(&H0)) - &H1)
        
        ![System Sounds] = vBuffer(&H0)
        
        'Save the Record
        .Update
        .UpdateBatch adAffectAllChapters
        
        chkEnableSoftwareSounds.Tag = ![System Sounds]
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
    If strStatus = VBA.vbNullString Then
        
        'Indicate that a process or operation is complete.
        Screen.MousePointer = vbDefault
        
        'Inform User
        vMsgBox "The Settings have successfully been " & VBA.IIf(IsNewRecord, "Saved", "Modified"), vbInformation, App.Title & " : Saving Report", Me
        
    End If 'Close respective IF..THEN block statement
    
Exit_ShpBttnApply_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub procedure
    
Handle_ShpBttnApply_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Applying Changes - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume 'Exit_ShpBttnApply_Click
    
End Sub

Private Sub ShpBttnClose_Click()
    Unload Me
End Sub

Private Sub ShpBttnLoadDefaults_Click()
    
    strStatus = VBA.IIf(strStatus <> VBA.vbNullString, strStatus, "Loading Defaults")
    
    'Confirm setting defaults
    If strStatus = "Loading Defaults" Then If vMsgBox("Are you sure you want to load default Sound settings for the Software?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then Exit Sub 'Quit this Sub procedure
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Copy default sounds to Software's default media location
    If vFso.FileExists(App.Path & "\Tools\Media\Default Media\vMsgBox - Notify.wav") Then vFso.CopyFile App.Path & "\Tools\Media\Default Media\vMsgBox - Notify.wav", App.Path & "\Tools\Media\vMsgBox - Notify.wav", True
    If vFso.FileExists(App.Path & "\Tools\Media\Default Media\vMsgBox - Question.wav") Then vFso.CopyFile App.Path & "\Tools\Media\Default Media\vMsgBox - Question.wav", App.Path & "\Tools\Media\vMsgBox - Question.wav", True
    If vFso.FileExists(App.Path & "\Tools\Media\Default Media\vMsgBox - Warning.wav") Then vFso.CopyFile App.Path & "\Tools\Media\Default Media\vMsgBox - Warning.wav", App.Path & "\Tools\Media\vMsgBox - Warning.wav", True
    If vFso.FileExists(App.Path & "\Tools\Media\Default Media\vMsgBox - Critical.wav") Then vFso.CopyFile App.Path & "\Tools\Media\Default Media\vMsgBox - Critical.wav", App.Path & "\Tools\Media\vMsgBox - Critical.wav", True
    If vFso.FileExists(App.Path & "\Tools\Media\Default Media\Progress - Error.wav") Then vFso.CopyFile App.Path & "\Tools\Media\Default Media\Progress - Error.wav", App.Path & "\Tools\Media\Progress - Error.wav", True
    If vFso.FileExists(App.Path & "\Tools\Media\Default Media\Progress - Complete.wav") Then vFso.CopyFile App.Path & "\Tools\Media\Default Media\Progress - Complete.wav", App.Path & "\Tools\Media\Progress - Complete.wav", True
    
    'Display the changes to the User
    If vFso.FileExists(App.Path & "\Tools\Media\vMsgBox - Notify.wav") Then txtSound(&H0).Text = App.Path & "\Tools\Media\vMsgBox - Notify.wav"
    If vFso.FileExists(App.Path & "\Tools\Media\vMsgBox - Question.wav") Then txtSound(&H1).Text = App.Path & "\Tools\Media\vMsgBox - Question.wav"
    If vFso.FileExists(App.Path & "\Tools\Media\vMsgBox - Warning.wav") Then txtSound(&H2).Text = App.Path & "\Tools\Media\vMsgBox - Warning.wav"
    If vFso.FileExists(App.Path & "\Tools\Media\vMsgBox - Critical.wav") Then txtSound(&H3).Text = App.Path & "\Tools\Media\vMsgBox - Critical.wav"
    If vFso.FileExists(App.Path & "\Tools\Media\Progress - Error.wav") Then txtSound(&H2).Text = App.Path & "\Tools\Media\Progress - Error.wav"
    If vFso.FileExists(App.Path & "\Tools\Media\Progress - Complete.wav") Then txtSound(&H3).Text = App.Path & "\Tools\Media\Progress - Complete.wav"
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    If strStatus = "Loading Defaults" Then vMsgBox "The default software sounds have been successfully loaded", , App.Title & " : Defaults", Me
    
    strStatus = VBA.IIf(strStatus = "Loading Defaults", VBA.vbNullString, strStatus)
    
End Sub
