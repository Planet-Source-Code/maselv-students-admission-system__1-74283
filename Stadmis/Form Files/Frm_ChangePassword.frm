VERSION 5.00
Begin VB.Form Frm_ChangePassword 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Change Password"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "Frm_ChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2040
      Width           =   3495
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
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
      ForeColor       =   &H00FF0000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00CFE1E2&
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1695
      Begin VB.Image ImgDBPhoto 
         Height          =   1455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image ImgVirtualPhoto 
         Height          =   135
         Left            =   240
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2550
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
      Caption         =   "      Cancel"
      AccessKey       =   "C"
      Picture         =   "Frm_ChangePassword.frx":038A
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
   Begin Stadmis.ShapeButton ShpBttnOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   2550
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
      Caption         =   "    OK"
      AccessKey       =   "O"
      Picture         =   "Frm_ChangePassword.frx":0724
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
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label LblTrials 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 Attempts"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   1170
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   -120
      Picture         =   "Frm_ChangePassword.frx":0CBE
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   5775
   End
   Begin VB.Image ImgHeader 
      Height          =   495
      Left            =   -120
      Picture         =   "Frm_ChangePassword.frx":14B4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Frm_ChangePassword"
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
'=  ALLOW USERS TO CHANGE THEIR ACCOUNT PASSWORDS

Option Explicit
Option Compare Binary

Public PwdChangeOK As Boolean
Public DontConfirmFromDatabase As Boolean

Private Trials%

Private Function DisplayUserDetails() As Boolean
On Local Error GoTo Handle_DisplayUserDetails_Error
    
    Dim MousePointerState%
    Dim nRs As New ADODB.Recordset
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Refresh connection without refreshing Recordsets
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB(, False)
    Set nRs = New ADODB.Recordset 'Create a new instance of the recordset object
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT * FROM [Qry_Users] WHERE [User ID] = " & User.User_ID, vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If the User's info exists then...
        If Not (.BOF And .EOF) Then
            
            lblUserName.Caption = User.User_Name & " - " & User.Full_Name
            
            'If the Record contains the User's Photo then...
            If Not VBA.IsNull(![Photo]) Then
                
                'Display User's Photo
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = "Photo"
                Set ImgVirtualPhoto.DataSource = nRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = User.User_Name & "'s Photo"
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
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
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Displaying User Details - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_DisplayUserDetails
    
End Function

Private Sub ShpBttnCancel_Click()
    Unload Me 'Unload this Form from the Memory
End Sub

Private Sub ShpBttnOK_Click()
On Local Error GoTo Handle_ShpBttnOK_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If TxtPassword(&H1).Text = VBA.vbNullString Then
        
        'If the database shouldn't be involved then...
        If Not DontConfirmFromDatabase Then
            
            'Check If the User has insufficient privileges to perform this Operation
            If Not ValidUserAccess(Me, &H2, &H7, , "You have no privilege of leaving your account without a password. Please contact System Administrator.") Then
                
                TxtPassword(&H1).SetFocus 'Move focus to the specified control
                GoTo Exit_ShpBttnOK_Click 'Quit this Procedure
                
            End If 'Close respective IF..THEN block statement
            
        Else
            'Do nothing
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    'If the entered password length is less than the required value then...
    If VBA.Len(TxtPassword(&H1).Text) < SoftwareSetting.Min_User_Password_Characters Then
        
        'Warn User
        vMsgBox "Please specify a Password of a minimum of " & SoftwareSetting.Min_User_Password_Characters & " characters.", vbInformation, App.Title & " : Invalid Password Entry", Me
        TxtPassword(&H2).Text = VBA.vbNullString  'Clear the confirmed one
        TxtPassword(&H1).SetFocus 'Move focus to the specified control
        GoTo Exit_ShpBttnOK_Click 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the entered Password does not match the confirmed one then...
    If TxtPassword(&H1).Text <> TxtPassword(&H2).Text Then
        
        vMsgBox "The password entry does not match the confirmed one.", vbExclamation, App.Title & " : Password Mismatch", Me
        TxtPassword(&H2).Text = VBA.vbNullString  'Clear the confirmed one
        TxtPassword(&H1).SetFocus 'Move focus to the specified control
        GoTo Exit_ShpBttnOK_Click 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'Hinder controls from responding to user-generated events
    ShpBttnOK.Enabled = False: ShpBttnCancel.Enabled = False
    
    If DontConfirmFromDatabase Then
        
        'If the Old Password is not correct then...
        If TxtPassword(&H0).Text <> TxtPassword(&H0).Tag Then
            
            vMsgBox "The entered Old Password is incorrect.", vbExclamation, App.Title & " : Invalid Password", Me
            TxtPassword(&H0).SetFocus 'Move focus to the specified control
            GoTo AccessDenied 'Branch unconditionally to a specified line
            
        End If 'Close respective IF..THEN block statement
        
        PwdChangeOK = True 'Denote that the Password change process is successful
        
    Else
        
        If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call procedure in Mdl_DataManipulators to create connection to Database
        Set vRs = New ADODB.Recordset
        With vRs 'Executes a series of statements for vRs Recordset
            
            'Retrieve Records of all Users with the Specified entries
            .Open "SELECT * FROM [Qry_Users] WHERE [User ID] = " & User.User_ID, vAdoCNN, adOpenKeyset, adLockReadOnly
            
            'If a User Record has been found then...
            If Not (.BOF And .EOF) Then
                
            Dim iPwd$
                
                If Not VBA.IsNull(![Password]) Then iPwd = ![Password]
                
                'If a password has been specified when the Account is not password protected then...
                If iPwd = VBA.vbNullString And TxtPassword(&H0).Text <> VBA.vbNullString Then
                    
                    GoTo AccessDenied 'Branch unconditionally to a specified line
                    
                'If a password has not been specified when the Account is not password protected then...
                ElseIf iPwd = VBA.vbNullString And TxtPassword(&H0).Text = VBA.vbNullString Then
                    
                    'Do nothing. Credentials are OK
                    
                Else 'Otherwise
                    
                    'If the Account is not password protected then if the entered password does not match the account password then Branch unconditionally to a specified line
                    If iPwd <> VBA.vbNullString Then If iPwd <> VBA.vbNullString Then If iPwd <> SmartEncrypt(TxtPassword(&H0).Text) Then GoTo AccessDenied
                    
                End If 'Close respective IF..THEN block statement
                
                iPwd = VBA.vbNullString
                
                .Close 'Close the opened object and any dependent objects
                
                'Change the User's Password
                'Create an update query that changes the Password field's value in the specified table on the Record with the specified User ID.
                .Open "SELECT * FROM [Tbl_Users] WHERE [User ID] = " & User.User_ID, vAdoCNN, adOpenKeyset, adLockPessimistic
                
                If Not (.BOF And .EOF) Then
                    
                    .Update
                    ![Password] = SmartEncrypt(TxtPassword(&H1).Text)
                    .Update
                    .UpdateBatch adAffectAllChapters
                    
                End If 'Close respective IF..THEN block statement
                
                If User.Hierarchy = &H0 Then VBA.SaveSetting App.Title, "Settings", "Admin Lock", ![Password]
                
                .Close 'Close the opened object and any dependent objects
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "The password has successfully been Changed.", vbInformation, App.Title & " : Password Change", Me
                
                PwdChangeOK = True 'Denote that the Password change process is successful
                
            End If 'Close respective IF..THEN block statement
            
            'Close the object and any dependent objects if opened
            If .State = adStateOpen Then .Close
            
        End With 'Close the WITH block statements
        
    End If 'Close respective IF..THEN block statement
    
    GoTo Exit_ShpBttnOK_Click
    
AccessDenied:
    
    'Denote that the User has not successfully logged in and increment the number of Password trials by 1
    Trials = Trials + &H1
    
    LblTrials.Caption = &H3 - Trials & " Attempts"
    
    Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
    
    'Inform User
    vMsgBox "Access denied due to wrong login credentials.", vbInformation, App.Title & " : Login Failed", Me
    
    TxtPassword(&H1).Text = VBA.vbNullString  'Clear the confirmed one
    TxtPassword(&H0).SetFocus 'Move focus to the specified control
    
    Set vRs = Nothing
    
    'If the number of Trials is 3 then...
    If Trials = &H3 Then
        
        'Inform User
        vMsgBox "The maximum attempt has been reached.", vbExclamation, App.Title & " : Login Failed", Me
        
        PwdChangeOK = True 'Denote that the Password change process is successful
        
        Unload Me 'Unload this Form from the Memory
        
    End If 'Close respective IF..THEN block statement
    
Exit_ShpBttnOK_Click:
    
    'Allow controls to respond to user-generated events
    ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True
    
    If PwdChangeOK Then Unload Me 'Unload this Form from the Memory
    
    Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_ShpBttnOK_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Clearing Entries - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume ' Exit_ShpBttnOK_Click
    
End Sub

Private Sub Form_Activate()
    If TxtPassword(&H0).Locked Then TxtPassword(&H1).SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Caption = App.Title & " : Change Password"
    TxtPassword(&H0).BackColor = VBA.IIf(TxtPassword(&H0).Locked, &H8000000F, &HC0FFFF)
    If Not DontConfirmFromDatabase Then Call DisplayUserDetails Else ImgDBPhoto.Picture = Nothing: ImgVirtualPhoto.Picture = Nothing
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    If Not PwdChangeOK Then If Not ShpBttnCancel.Visible Then Cancel = True Else Cancel = Not CloseFrm(Me, , False)
    If PwdChangeOK And Not DontConfirmFromDatabase Then PwdChangeOK = False
End Sub

Private Sub txtPassword_GotFocus(Index As Integer)
    TxtPassword(Index).SelStart = &H0: TxtPassword(Index).SelLength = VBA.Len(TxtPassword(Index).Text)
End Sub
