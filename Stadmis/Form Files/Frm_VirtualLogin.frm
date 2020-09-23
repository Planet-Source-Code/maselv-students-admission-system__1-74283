VERSION 5.00
Begin VB.Form Frm_VirtualLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Virtual Login"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "Frm_VirtualLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
      Left            =   1680
      ScaleHeight     =   270
      ScaleWidth      =   1665
      TabIndex        =   10
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
         TabIndex        =   11
         Top             =   0
         Width           =   1365
      End
   End
   Begin VB.Frame Fra_Login 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         Picture         =   "Frm_VirtualLogin.frx":08CA
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
         Picture         =   "Frm_VirtualLogin.frx":0C64
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
         TabIndex        =   1
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter user name and password to connect to the server ..."
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   3615
   End
   Begin VB.Image ImgUser 
      Height          =   240
      Index           =   1
      Left            =   4080
      Picture         =   "Frm_VirtualLogin.frx":11FE
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
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   960
   End
   Begin VB.Image ImgFooter 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_VirtualLogin.frx":1AC8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Image ImgUser 
      Height          =   600
      Index           =   0
      Left            =   3840
      Picture         =   "Frm_VirtualLogin.frx":22BE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image ImgHeader 
      Height          =   855
      Left            =   0
      Picture         =   "Frm_VirtualLogin.frx":3188
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_VirtualLogin"
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
Option Compare Binary

Public mySetNo&, mySetIndex&
Public myWholeSet As Boolean

Private Trials%
Private LoginSucceeded As Boolean

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Sub ShpBttnCancel_Click()
    Unload Me
End Sub

Private Sub ShpBttnOK_Click()
    
    'If the User Name has not been entered then...
    If VBA.Trim$(TxtUserName.Text) = VBA.vbNullString Then
        
        'Inform User
        vMsgBox "Please enter User Name", vbInformation, App.Title & " : Unspecified User Name", Me
        TxtUserName.SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Sub-Procedure
        
    End If 'Close respective IF..THEN block statement
    
    'Change Mouse Pointer to show Processing state
    Screen.MousePointer = vbHourglass
    
    ShpBttnOK.Enabled = False: ShpBttnCancel.Enabled = False
    
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
            
            'If the User Logon Time Limit has been specified then...
            If ![Parental Control] Then
                
                If Not VBA.IsNull(![Start Time]) And Not VBA.IsNull(![End Time]) Then
                    
                    If VBA.DateDiff("n", ![Start Time], VBA.FormatDateTime(VBA.Now, vbShortTime)) >= &H0 And VBA.DateDiff("n", ![End Time], VBA.FormatDateTime(VBA.Now, vbShortTime)) >= &H0 Then
                        
                        'Indicate that a process or operation is complete.
                        Screen.MousePointer = vbDefault
                        
                        'Warn User
                        vMsgBox "Your account has a limited logon time to operate between " & ![Start Time] & " and " & ![End Time] & ". It's now " & VBA.Format$(VBA.Now, "hh:nn:ss AMPM") & ". Please contact Software Administrator", vbExclamation, App.Title & " : Access Denied", Me
                        
                        TxtUserName.SetFocus 'Move focus to the specified control
                        ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True 'Enable buttons
                        
                        Exit Sub 'Quit this Procedure
                        
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            '------------------------------------------------------------------------------------------
            'Group privileges should supersede User privileges. The following codes ensure that
            '------------------------------------------------------------------------------------------
            
            VirtualUser.Privileges = SmartDecrypt(vRs![User Privileges])
            
            If (vRs![Hierarchy] = &H1) Then vBuffer(&H0) = def_Privileges Else vBuffer(&H0) = VirtualUser.Privileges
            
            Dim myAccessRightsArray() As String
            myAccessRightsArray = VBA.Split(vBuffer(&H0), "|")
            
            If (mySetNo - &H1 > UBound(myAccessRightsArray)) And Not (vRs![Hierarchy] = &H1) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "You have also insufficient privileges to perform the Operation", vbExclamation, App.Title & " : Access Denied", Me
                
                ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True
                
                Exit Sub 'Quit this Sub-Procedure
                
            End If 'Close respective IF..THEN block statement
            
            If myWholeSet Then
                vBuffer(&H1) = VBA.Int((VBA.Replace(myAccessRightsArray(mySetNo - &H1), "1", VBA.vbNullString) = VBA.vbNullString))
            Else
                If mySetNo - &H1 <= UBound(myAccessRightsArray) Then If mySetIndex <= VBA.Len(myAccessRightsArray(mySetNo - &H1)) Then vBuffer(&H1) = VBA.Int((VBA.Mid$(myAccessRightsArray(mySetNo - &H1), mySetIndex, &H1) = &H1))
            End If 'Close respective IF..THEN block statement
            
            If Not VBA.CBool(VBA.Val(vBuffer(&H1))) And Not (vRs![Hierarchy] = &H1) Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "You also have insufficient privileges to perform the Operation", vbExclamation, App.Title & " : Access Denied", Me
                
                ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True
                
                Exit Sub 'Quit this Sub-Procedure
                
            End If 'Close respective IF..THEN block statement
            
            VirtualUser.User_ID = ![User ID]
            VirtualUser.User_Name = ![User Name]
            VirtualUser.Privileges = vBuffer(&H0)
            
            Unload Me 'Unload this Form from the Memory
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            Exit Sub 'Quit this Sub-Procedure
            
        Else 'If a User Record has not been found then...
            
AccessDenied:
            
            'Denote that the User has not successfully logged in and increment the number of Password trials by 1
            LoginSucceeded = False: Trials = Trials + &H1
            
            LblTrials.Caption = &H3 - Trials & " Attempts"
            
            Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
            
            'Inform User
            vMsgBox "Access denied for User (" & VBA.Trim$(TxtUserName.Text) & ")", vbInformation, App.Title & " : Login Failed", Me
            
            ShpBttnOK.Enabled = True: ShpBttnCancel.Enabled = True
            
            TxtPassword.SetFocus 'Move focus to the specified control
            TxtPassword.SelStart = &H0: TxtPassword.SelLength = VBA.Len(TxtPassword.Text)
            
            Set vRs = Nothing
            
            'If the number of Trials is 3 then...
            If Trials = &H3 Then
                
                LblTrials.Visible = False
                
                'Inform User
                vMsgBox "The maximum Login attempts has been reached...", vbExclamation, App.Title & " : Login Failed", Me
                
                Unload Me 'Unload this Form from the Memory
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End With 'End WITH Statement
    
    Screen.MousePointer = vbDefault 'Change Mouse Pointer to show end of Processing state
    
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title & " : Virtual Login"
    VirtualUser.User_ID = &H0: VirtualUser.User_Name = VBA.vbNullString
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
