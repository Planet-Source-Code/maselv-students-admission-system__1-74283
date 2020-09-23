VERSION 5.00
Begin VB.Form Frm_Login 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Login"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "Frm_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerCapLock 
      Interval        =   50
      Left            =   -600
      Top             =   480
   End
   Begin VB.PictureBox picCapsON 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1800
      ScaleHeight     =   270
      ScaleWidth      =   1665
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label lblCapsON 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAPS LOCK ON"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   150
         TabIndex        =   13
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.CheckBox ChkRememberMe 
      BackColor       =   &H00CFE1E2&
      Caption         =   "&Remember Me"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1960
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Fra_Login 
      BackColor       =   &H00CFE1E2&
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   4575
      Begin VB.TextBox TxtPassword 
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
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtUserName 
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
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin Stadmis.ShapeButton ShpBttnCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "Frm_Login.frx":08CA
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
         Left            =   3360
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "Frm_Login.frx":0C64
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
      Begin VB.Label LblLogin 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   750
      End
      Begin VB.Label LblLogin 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Label lblForgotPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot your password?"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Click here if you have forgotten your account password"
      Top             =   2160
      Width           =   1560
   End
   Begin VB.Label LblLogin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   1560
   End
   Begin VB.Label LblLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter user name and password to connect to the server ..."
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   3615
   End
   Begin VB.Image ImgUser 
      Height          =   240
      Index           =   1
      Left            =   4080
      Picture         =   "Frm_Login.frx":11FE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   240
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
      Left            =   3480
      TabIndex        =   8
      Top             =   2070
      Width           =   960
   End
   Begin VB.Image ImgFooter 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_Login.frx":1AC8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Image ImgUser 
      Height          =   600
      Index           =   0
      Left            =   3840
      Picture         =   "Frm_Login.frx":22BE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image ImgHeader 
      Height          =   855
      Left            =   0
      Picture         =   "Frm_Login.frx":3188
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_Login"
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


'********************************************************************************************************
'*                                       APPROVED USERS                                                 *
'********************************************************************************************************
'*  Another technique to help prevent abuse and misuse of computer data is to limit the use of          *
'*  computers and data files to approved persons. Security software can verify the identity of          *
'*  computer users and limit their privileges to use, view, and alter files. The software also          *
'*  securely records their actions to establish accountability. Military organizations give access      *
'*  rights to classified, confidential, secret, or top-secret information according to the              *
'*  corresponding security clearance level of the user. Other types of organizations also classify      *
'*  information and specify different degrees of protection.                                            *
'********************************************************************************************************
'*                                         PASSWORDS                                                    *
'********************************************************************************************************
'*  Passwords are confidential sequences of characters that allow approved persons to make use of       *
'*  specified computers, software, or information. To be effective, passwords must be difficult to      *
'*  guess and should not be found in dictionaries. Effective passwords contain a variety of             *
'*  characters and symbols that are not part of the alphabet. To thwart imposters, computer systems     *
'*  usually limit the number of attempts and restrict the time it takes to enter the correct password.  *
'********************************************************************************************************


Option Explicit
Option Compare Binary

Private Trials%
Private LoginSucceeded As Boolean

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Function ApplyTheme() As Boolean

    '**************************************************************
    'Apply Theme Settings
    
    Me.BackColor = tTheme.tBackColor
    Fra_Login.BackColor = Me.BackColor
    
    ImgHeader.Picture = tTheme.tImagePicture
    ImgFooter.Picture = ImgHeader.Picture
    
    ShpBttnCancel.ForeColor = tTheme.tButtonForeColor
    ShpBttnCancel.BackColor = tTheme.tButtonBackColor
    
    ShpBttnOK.ForeColor = tTheme.tButtonForeColor
    ShpBttnOK.BackColor = tTheme.tButtonBackColor
    
    For vIndex(&H0) = &H0 To LblLogin.UBound Step &H1
        LblLogin(vIndex(&H0)).ForeColor = tTheme.tForeColor
    Next vIndex(&H0)
    
    LblTrials.ForeColor = tTheme.tWarningForeColor
    lblCapsON.ForeColor = tTheme.tWarningForeColor
    
    ChkRememberMe.ForeColor = tTheme.tForeColor
    lblForgotPassword.ForeColor = tTheme.tWarningForeColor
    
    '**************************************************************
    
End Function

Private Sub ShpBttnCancel_Click()
    
    'Denote that the User has not successfully logged in
    LoginSucceeded = False
    Unload Me 'Unload this Form from the Memory
    
End Sub

Private Sub ShpBttnOK_Click()
    
    'If the User Name has not been entered then...
    If VBA.Trim$(TxtUserName.Text) = VBA.vbNullString Then
        
        'Inform User
        vMsgBox "Please enter User Name", vbInformation, App.Title & " : Unspecified User Name", Me
        TxtUserName.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    ShpBttnOK.Enabled = False: ShpBttnCancel.Enabled = False
    
    TxtUserName.Text = CapAllWords(VBA.Trim$(TxtUserName.Text))
    
    ConnectDB 'Call procedure in Mdl_DataManipulators to create connection to Database
    With vRs 'Executes a series of statements for vRs Recordset
        
        'Retrieve Records of all Users with the Specified entries
        .Open "SELECT * FROM [Qry_Users] ORDER BY [Hierarchy] ASC", vAdoCNN, adOpenKeyset, adLockReadOnly
        
        'If a User Record has been found then...
        If Not (.BOF And .EOF) Then
            
            .Find "[User Name] = '" & VBA.Trim(Replace(TxtUserName.Text, "'", "''")) & "'"
            
            If Not .EOF Then
                
                If .AbsolutePosition > SoftwareSetting.Licences.Max_Users + &H2 Then
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Warn User
                    vMsgBox "Your User Account is not among the accounts licensed by this Program. The Program Licence allows a maximum of only " & SoftwareSetting.Licences.Max_Users & " Software User" & VBA.IIf(SoftwareSetting.Licences.Max_Users = &H1, VBA.vbNullString, "s") & ". Please contact Administrator for Licence Renewal.", vbExclamation, App.Title & " : Limited Capabilities", Me
                    
                    TxtPassword.Text = VBA.vbNullString 'Discard the entered password
                    TxtUserName.SetFocus 'Move focus to the specified control
                    ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True 'Enable buttons
                    
                    Exit Sub 'Quit this Procedure
                    
                End If 'Close respective IF..THEN block statement
                
            Else
                
                GoTo AccessDenied 'Branch unconditionally to a specified line
                
            End If 'Close respective IF..THEN block statement
            
            Dim iPwd$
            
            If Not VBA.IsNull(![Password]) Then iPwd = ![Password]
            
            'If a password has been specified when the Account is not password protected then...
            If iPwd = VBA.vbNullString And TxtPassword.Text <> VBA.vbNullString Then
                
                GoTo AccessDenied 'Branch unconditionally to a specified line
                
            'If a password has not been specified when the Account is not password protected then...
            ElseIf iPwd = VBA.vbNullString And TxtPassword.Text = VBA.vbNullString Then
                
                'Do nothing. Credentials are OK
                
            Else 'Otherwise
                
                'If the Account is not password protected then if the entered password does not match the account password then Branch unconditionally to a specified line
                If iPwd <> VBA.vbNullString Then If iPwd <> VBA.vbNullString Then If iPwd <> SmartEncrypt(TxtPassword.Text) Then GoTo AccessDenied
                
            End If 'Close respective IF..THEN block statement
            
            If ![Hierarchy] = &H0 Then
                
                User.Privileges = SmartDecrypt(![User Privileges])
                vBuffer(&H0) = VBA.GetSetting(App.Title, "Settings", "Admin Lock", VBA.vbNullString)
                
                If iPwd <> vBuffer(&H0) Then
                    
                    If IsNull(![Security Question 1]) Or IsNull(![Security Ans 1]) Or IsNull(![Security Question 2]) Or IsNull(![Security Ans 2]) Then
                        GoTo ResetEntry
                    Else
                        'Do nothing
                    End If 'Close respective IF..THEN block statement
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Check If the User has insufficient privileges to perform this Operation
                    vMsgBox "The specified User information has not been verified by " & App.Title & " System. A system other than this program altered the details. Please confirm Security Questions?", vbCritical, App.Title & " : Software Data Breach!!!", Me
                    
                    Frm_SecurityCheck.iConfirmingUser = True
                    CenterForm Frm_SecurityCheck, Me
                    
                    If vBuffer(&H0) = "Succeeded" Then
                        
                        VBA.SaveSetting App.Title, "Settings", "Admin Lock", VBA.IIf(VBA.IsNull(![Password]), "", ![Password])
                        GoTo ContinueLogginIn
                        
                    End If 'Close respective IF..THEN block statement
                    
                    End 'Halt the Application
                    
ResetEntry:
                    
                    'Indicate that a process or operation is in progress.
                    Screen.MousePointer = vbHourglass
                    
                    .Close 'Close the opened object and any dependent objects
                    
                    'Delete the breached User Account
                    .Open "DELETE * FROM [Tbl_Users] WHERE [National ID] = '24616804'", vAdoCNN, adOpenKeyset, adLockPessimistic
                    
                    End 'Halt the Application
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
ContinueLogginIn:
            
            vBuffer(&H0) = VBA.vbNullString: vBuffer(&H1) = VBA.vbNullString
            iPwd = VBA.vbNullString
            
            'If the User Status has been specified then...
            If Not VBA.IsNull(![User Status]) Then
                
                'If the User has been disabled then...
                If ![User Status] <> VBA.vbNullString Then
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Warn User
                    vMsgBox "Your account has been disabled. Please contact Software Administrator", vbExclamation, App.Title & " : Access Denied", Me
                    
                    TxtPassword.Text = VBA.vbNullString 'Discard the entered password
                    TxtUserName.SetFocus 'Move focus to the specified control
                    ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True 'Enable buttons
                    
                    Exit Sub 'Quit this Procedure
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            User.Parental_Control_ON = ![Parental Control]
            
            'If the User Logon Time Limit has been specified then...
            If User.Parental_Control_ON Then
                
                If Not VBA.IsNull(![Start Time]) And Not VBA.IsNull(![End Time]) Then
                    
                    User.Start_Time = ![Start Time]
                    User.End_Time = ![End Time]
                    
                    If VBA.DateDiff("n", User.Start_Time, VBA.FormatDateTime(VBA.Now, vbShortTime)) >= &H0 And VBA.DateDiff("n", User.End_Time, VBA.FormatDateTime(VBA.Now, vbShortTime)) >= &H0 Then
                        
                        'Indicate that a process or operation is complete.
                        Screen.MousePointer = vbDefault
                        
                        'Warn User
                        vMsgBox "Your account has a limited logon time to operate between " & User.Start_Time & " and " & User.End_Time & ". It's now " & VBA.Format$(VBA.Now, "hh:nn:ss AMPM") & ". Please contact Software Administrator", vbExclamation, App.Title & " : Access Denied", Me
                        
                        TxtUserName.SetFocus 'Move focus to the specified control
                        ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True 'Enable buttons
                        
                        Exit Sub 'Quit this Procedure
                        
                    End If 'Close respective IF..THEN block statement
                    
                Else
                    
                    User.Parental_Control_ON = False
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            User.User_ID = ![User ID]: User.Login_Name = ![User Name]
            
            'Check settings if the User should change his/her password
            vIndex(&H0) = VBA.Val(VBA.GetSetting(App.Title, "Settings", User.User_ID & " Request Password change after Login", "0"))
            
            'If so then...
            If vIndex(&H0) = &H1 Then
                
                .Close 'Close the opened object and any dependent objects
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "You have been requested to change your Account password. Please continue.", vbInformation, App.Title & " : Account Specifications", Me
                
                'Display the 'Change Password' Form for User to make the Entries
                
                With Frm_ChangePassword
                    
                    .ShpBttnCancel.Visible = False: .TxtPassword(&H0).Locked = True
                    .TxtPassword(&H0).Text = TxtPassword.Text
                    .LblTrials.Visible = False
                    
                    CenterForm Frm_ChangePassword, Me
                    
                End With 'Close WITH block statement
                
                'Indicate that a process or operation is in progress.
                Screen.MousePointer = vbHourglass
                
                VBA.DeleteSetting App.Title, "Settings", User.User_ID & " Request Password change after Login"
                
            End If 'Close respective IF..THEN block statement
            
            'Clear the entered Password if any
            TxtPassword.Text = VBA.vbNullString
            
            Me.Hide: CenterForm Frm_PleaseWait, , False: VBA.DoEvents
            Frm_PleaseWait.ImgProgressBar.Width = &H0
            
            'ConnectDB 'Call procedure in Mdl_DataManipulators to create connection to Database
            
            If vRs.State = adStateOpen Then vRs.Close
            
            'Retrieve Records of all Users with the Specified entries
            vRs.Open "SELECT * FROM [Qry_Users] WHERE [User Name] = '" & VBA.Trim(Replace(TxtUserName.Text, "'", "''")) & "'", vAdoCNN, adOpenKeyset, adLockReadOnly
            
            'Assign User Information
            With User
                
                If Not VBA.IsNull(VBA.Trim(vRs![Hierarchy])) Then .Hierarchy = vRs![Hierarchy]
                If Not VBA.IsNull(VBA.Trim(vRs![Full Name])) Then .User_Name = vRs![Full Name]
                If Not VBA.IsNull(VBA.Trim(vRs![User Privileges])) Then .Privileges = SmartDecrypt(vRs![User Privileges])
                
            End With 'Close WITH block statement
            
            LoginSucceeded = True 'Denote that the User has successfully logged in
            
            vRs.Close 'Close the opened object and any dependent objects
            
            vBuffer(&H0) = VBA.GetSetting(App.Title, "Settings", "Last User", VBA.vbNullString)
            
            'If the User wants the Software to auto fill the Username details then Save to Settings
            If ChkRememberMe.Value = vbChecked Then
                VBA.SaveSetting App.Title, "Settings", "Last User", User.Login_Name
            Else
                If vBuffer(&H0) <> VBA.vbNullString Then VBA.DeleteSetting App.Title, "Settings", "Last User"
            End If
            
            'Limit Login Records to the specified Maximum Records required
            vRs.Open "DELETE * FROM [Tbl_Login] WHERE [Login No] NOT IN (SELECT TOP " & SoftwareSetting.Max_Login_Records - &H1 & " [Login No] FROM [Tbl_Login] ORDER BY [Login Date] DESC)", vAdoCNN, adOpenKeyset, adLockPessimistic
            
            Dim vUser$, vLoginDate$
            Dim vSuccessful_Logout As Boolean
            
            Dim vDrive
            Dim ObjNet As Object
            
            Set ObjNet = CreateObject("WScript.Network")
            Set vDrive = vFso.GetDrive(vFso.GetDriveName(GetWindowsDir))
            
            User.Device_Name = ObjNet.ComputerName
            User.Device_Account_Name = ObjNet.UserName
            User.Device_Serial_No = vDrive.SerialNumber
            
            Set ObjNet = Nothing: Set vDrive = Nothing
            
            ConnectDB 'Call procedure in Mdl_DataManipulators to create connection to Database
            
            'Check if the Application terminated correctly when it was last opened
            vRs.Open "SELECT TOP 1 * FROM [Qry_Login] WHERE [Device Serial] = '" & User.Device_Serial_No & "' ORDER BY [Login Date] DESC", vAdoCNN, adOpenKeyset, adLockReadOnly
            
            vSuccessful_Logout = True
            
            'If the Application terminated correctly when it was last opened
            If vRs.RecordCount Then
                
                vUser = vRs![User Name]
                vLoginDate = vRs![Login Date]
                vSuccessful_Logout = vRs![Successful Logout]
                
            End If 'Close respective IF..THEN block statement
            
            vRs.Close 'Close the opened object and any dependent objects
            
            vRs.Open "SELECT * FROM [Tbl_Login] ORDER BY [Login Date] DESC", vAdoCNN, adOpenKeyset, adLockPessimistic
            vRs.AddNew
            
            vRs![User ID] = VBA.IIf(VirtualUser.User_Name <> VBA.vbNullString, VirtualUser.User_ID, User.User_ID)
            vRs![Device Serial] = User.Device_Serial_No
            vRs![Device Name] = VBA.Trim(User.Device_Name)
            vRs![Device Account] = VBA.Trim(User.Device_Account_Name)
            
            vRs.Update
            vRs.UpdateBatch adAffectAllChapters
            
            User.Login_ID = vRs![Login No]
            vRs.Close 'Close the opened object and any dependent objects
            
            User.Login_Date = VBA.Now: Set vRs = Nothing
            
            LoginSucceeded = True
            
            Me.Hide
            Frm_PleaseWait.ImgProgressBar.Width = Frm_PleaseWait.lblProgressBar.Width / &H2
            Frm_PleaseWait.lblStatus.Caption = "50% Complete": VBA.DoEvents
            Frm_PleaseWait.lblInfo.Caption = "Loading...": Frm_PleaseWait.ShpBttnCancel.Visible = False
            
            vWait = True: vStartupComplete = False
            Frm_Main.Show
            
            ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True
            
            'Wait for the Main Form to load completely
            Do While vWait
                VBA.DoEvents 'Yield execution so that the operating system can process other events
            Loop
            
            Frm_PleaseWait.ImgProgressBar.Width = Frm_PleaseWait.lblProgressBar.Width 'Set to Maximum
            Frm_PleaseWait.lblStatus.Caption = "100% Complete": VBA.DoEvents
            
            Unload Frm_PleaseWait 'Unload this Form from the Memory
            Unload Frm_Login 'Unload this Form from the Memory
            
            'If the Software was not correctly closed the last time it was opened then...
            If Not vSuccessful_Logout Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "This Software was not correctly closed the last time it was opened by '" & vUser & "' on {" & VBA.Format$(vLoginDate, "ddd dd MMM yyyy HH:nn:ss") & "}.", vbExclamation, App.Title & " : Improper Shutdown", Frm_Main
                
                'Indicate that a process or operation is in progress.
                Screen.MousePointer = vbHourglass
                
            End If 'Close respective IF..THEN block statement
            
            Dim Frm As Form
            
            'Unload all accessory forms properly ensuring restoration of resources
            For Each Frm In VB.Forms
                
                'Close all other open Forms in this Application without alerts, apart from this Main one
                If Frm.Name <> "Frm_Main" Then vSilentClosure = True: Unload Frm
                
            Next Frm 'Move to the next open Form
            
            vSilentClosure = False 'Allow closure alerts to be displayed to the User
            vStartupComplete = True
            
            VBA.DoEvents 'Yield execution so that the operating system can process other events.
            
        Else 'If a User Record has not been found then...
            
AccessDenied:
            
            'Denote that the User has not successfully logged in and increment the number of Password trials by 1
            LoginSucceeded = False: Trials = Trials + &H1
            
            LblTrials.Caption = &H3 - Trials & " Attempts"
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Inform User
            vMsgBox "Access denied for User (" & VBA.Trim$(TxtUserName.Text) & ")", vbExclamation, App.Title & " : Login Failed", Me
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
            ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True
            
            TxtPassword.SetFocus 'Move focus to the specified control
            TxtPassword.SelStart = &H0: TxtPassword.SelLength = VBA.Len(TxtPassword.Text)
            Set vRs = Nothing
            
            'If the number of Trials is 3 then...
            If Trials = &H3 Then
                
                LblTrials.Visible = False
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "The maximum Login attempts has been reached. The Software will shut down in £ seconds..", vbExclamation, App.Title & " : Login Failed", Me, , , &HA
                
                End 'Halt the Application
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End With 'End WITH Statement
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Activate()
On Local Error Resume Next
    If TxtUserName.Text <> VBA.vbNullString Then TxtPassword.SetFocus
End Sub

Private Sub Form_Load()
On Local Error GoTo Handle_Form_Load_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Me.Caption = App.Title & " : Login"
    
    'Call ApplyTheme
    
    Dim vAdminAccountID&
    Dim vDevAccountID(&H1) As String
    
    vBuffer(&H2) = VBA.vbNullString 'Initialize variable
    
    'If no accounts have been defined in the Database then create one for Administrators
    
    ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve Records of all Software Users
        .Open "SELECT * FROM [Tbl_Users]", vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If there are Records then...
        If (.BOF And .EOF) Then
            
            vRsTmp.Open "SELECT * FROM [Tbl_Users] WHERE [National ID] = '24616804'", vAdoCNN, adOpenKeyset, adLockPessimistic
            
            'If there are Records then...
            If (vRsTmp.BOF And vRsTmp.EOF) Then
                vRsTmp.AddNew 'Create a new record for the Recordset object.
            Else
                vRsTmp.Update
            End If 'Close respective IF..THEN block statement
            
            vRsTmp![Date Employed] = VBA.DateSerial(1986, &H3, &H9)
            vRsTmp![Title] = "Mr."
            vRsTmp![Surname] = "Masika"
            vRsTmp![Other Names] = "Elvas .S."
            vRsTmp![Gender] = "Male"
            vRsTmp![User Name] = "Admin"
            vRsTmp![Password] = SmartEncrypt("Maselv...")
            vRsTmp![Account Name] = "Software Developer"
            vRsTmp![User Privileges] = SmartEncrypt(def_Privileges)
            vRsTmp![Hierarchy] = &H0
            vRsTmp![National ID] = 24616804
            vRsTmp![Postal Address] = "P.O Box 137, Bungoma 50200, Kenya"
            vRsTmp![Location] = "Bungoma"
            vRsTmp![Phone No] = "254724688172" & VBA.vbCrLf & "254751041184"
            vRsTmp![E-mail Address] = "masika_elvas@programmer.net"
            vRsTmp![Marital Status] = "Single"
            vRsTmp![Birth Date] = VBA.DateSerial(1986, &H3, &H9)
            vRsTmp![Occupation] = "Software Developer, Lexeme Kenya Ltd"
            vRsTmp![Position] = "Software Manager"
            
            vRsTmp![Security Question 1] = "What is the first name of your favourite aunt?"
            vRsTmp![Security Ans 1] = SmartEncrypt("kate")
            vRsTmp![Security Question 2] = "What is your first pet's name??"
            vRsTmp![Security Ans 2] = SmartEncrypt("sea dog")
            
            vRsTmp![PIN No] = Null
            vRsTmp![NSSF No] = Null
            vRsTmp![NHIF No] = Null
            vRsTmp![Report To] = Null
            vRsTmp![Bank Acc No] = "00109204459100"
            vRsTmp![Bank Name] = "Cooperative Bank of Kenya, Migori Branch"
            vRsTmp![Deceased] = False
            vRsTmp![Virtual Entry] = True
            vRsTmp![User ID] = &H0
            
            'Save/Modify Record
            vRsTmp.Update
            vRsTmp.UpdateBatch adAffectAllChapters
            
            VBA.SaveSetting App.Title, "Settings", "Admin Lock", vRsTmp![Password]
            
            vRsTmp.Close 'Close the opened object and any dependent objects
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_Form_Load:
    
    User.Login_ID = &H0
    User.Login_Name = VBA.vbNullString
    User.User_ID = &H0
    User.User_Name = VBA.vbNullString
    User.Hierarchy = &H0
    User.Privileges = VBA.vbNullString
    
    'Retrieve Settings to remember the User who has just logged in
    vBuffer(&H0) = VBA.GetSetting(App.Title, "Settings", "Last User", VBA.vbNullString)
    
    If vBuffer(&H0) <> VBA.vbNullString Then TxtUserName.Text = vBuffer(&H0)
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_Form_Load_Error:
    
    'AutomatiOn Local Error for Reports
    If Err.Number = -2147024770 Then Resume
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Form Load Error - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_Form_Load
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If the User has not successfully logged in then confirm application exit
    If Not LoginSucceeded Then Cancel = Not CloseFrm(Me, , False): If Not Cancel Then End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Erase vBuffer 'Initialize variable
    
    If LoginSucceeded Then Exit Sub
    
    'Description: Unloads an application properly ensuring restoration of resources
    '--------------------------------------------------------------------------------
    
    'For all the open Forms in this Software...
    For Each vFrm(&H1) In Forms
        
        'If the form is not this Main Form then unload it
        If vFrm(&H1).Name <> Me.Name Then vSilentClosure = True: Unload vFrm(&H1)
        
    Next vFrm(&H1) 'Move to the next open Form if any
    '--------------------------------------------------------------------------------
    
    vSilentClosure = False
    
    'Call Procedure in Mdl_Stadmis Module to free memory and system resources
    Call PerformMemoryCleanup
    
    vStartupComplete = True
    
End Sub

Private Sub ImgFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblForgotPassword.FontUnderline Then lblForgotPassword.FontUnderline = False
End Sub

Private Sub lblForgotPassword_Click()
    CenterForm Frm_SecurityCheck, Me
End Sub

Private Sub lblForgotPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblForgotPassword.Appearance = &H1: lblForgotPassword.BorderStyle = &H1: lblForgotPassword.ForeColor = &HC00000
End Sub

Private Sub lblForgotPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lblForgotPassword.FontUnderline Then lblForgotPassword.FontUnderline = True
End Sub

Private Sub lblForgotPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblForgotPassword.Appearance = &H0: lblForgotPassword.BorderStyle = &H0: lblForgotPassword.ForeColor = &HC00000
End Sub

Private Sub TimerCapLock_Timer()
    'Inform User if CAPS LOCK is ON
    picCapsON.Visible = (GetKeyState(vbKeyCapital) <> &H0)
End Sub

Private Sub txtPassword_GotFocus()
    TxtPassword.SelStart = &H0: TxtPassword.SelLength = VBA.Len(TxtPassword.Text)
End Sub

Private Sub TxtUserName_GotFocus()
    TxtUserName.SelStart = &H0: TxtUserName.SelLength = VBA.Len(TxtUserName.Text)
End Sub
