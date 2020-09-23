VERSION 5.00
Begin VB.Form Frm_SecurityCheck 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Security Verification"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "Frm_SecurityCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
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
      ItemData        =   "Frm_SecurityCheck.frx":09EA
      Left            =   240
      List            =   "Frm_SecurityCheck.frx":09EC
      TabIndex        =   3
      Top             =   1680
      Width           =   4815
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
      ItemData        =   "Frm_SecurityCheck.frx":09EE
      Left            =   240
      List            =   "Frm_SecurityCheck.frx":09F0
      TabIndex        =   7
      Top             =   2640
      Width           =   4815
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2040
      Width           =   3615
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
      Index           =   1
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3000
      Width           =   3615
   End
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3600
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
      Picture         =   "Frm_SecurityCheck.frx":09F2
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
      Left            =   4200
      TabIndex        =   10
      Top             =   3600
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
      Picture         =   "Frm_SecurityCheck.frx":0D8C
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
      TabIndex        =   14
      Top             =   3720
      Width           =   960
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label lblSecurityQuestion1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security Question 1:"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Label lblSecurityQuestion2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Security Question 2:"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1785
   End
   Begin VB.Label lblSecurityQuestion2Ans 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Answer:"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Label lblSecurityQuestion1Ans 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Answer:"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please fill in the correct security details in order to reset your forgotten password."
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
      Height          =   435
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password"
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
      TabIndex        =   12
      Top             =   120
      Width           =   1530
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00808080&
      Height          =   2535
      Left            =   120
      Top             =   960
      Width           =   5055
   End
   Begin VB.Image ImgFooter 
      Height          =   615
      Left            =   -120
      Picture         =   "Frm_SecurityCheck.frx":1326
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Image ImgHeader 
      Height          =   975
      Left            =   0
      Picture         =   "Frm_SecurityCheck.frx":1BC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Frm_SecurityCheck"
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
Option Compare Binary

Public iConfirmingUser As Boolean

Private Trials%
Private VerificationSucceeded As Boolean

Private Sub Form_Load()
    
    Me.Caption = App.Title & " : Security Verification"
    
    'Fill Security Questions
    
    vArrayList = VBA.Split(def_SecurityQuestions1, "|")
    
    For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
        cboSecurityQuestion1.AddItem vArrayList(vIndex(&H0))
    Next vIndex(&H0)
    
    vArrayList = VBA.Split(def_SecurityQuestions2, "|")
    
    For vIndex(&H0) = &H0 To UBound(vArrayList) Step &H1
        cboSecurityQuestion2.AddItem vArrayList(vIndex(&H0))
    Next vIndex(&H0)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If the User has not successfully specified correct credentials then confirm Form exit
    If Not VerificationSucceeded Then vBuffer(&H0) = "Cancelled": Cancel = Not CloseFrm(Me, , False) Else vBuffer(&H0) = "Succeeded"
End Sub

Private Sub ShpBttnCancel_Click()
    Unload Me 'Unload this Form from the Memory
End Sub

Private Sub ShpBttnOK_Click()
On Local Error GoTo Handle_ShpBttnOK_Click_Error
    
    'If the User has not specified the first Security Question then...
    If cboSecurityQuestion1.Text = VBA.vbNullString And Not cboSecurityQuestion1.Locked Then
        
        'Warn User to select an existing Record
        vMsgBox "Please specify the first Security Question.", vbExclamation, App.Title & " : Record not Selected", Me
        cboSecurityQuestion1.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not answered the first Security Question then...
    If txtSecurityQuestionAns(&H0).Text = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please enter the answer for the first Security Question.", vbExclamation, App.Title & " : Record not Selected", Me
        txtSecurityQuestionAns(&H0).SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not specified the second Security Question then...
    If cboSecurityQuestion2.Text = VBA.vbNullString And Not cboSecurityQuestion2.Locked Then
        
        'Warn User to select an existing Record
        vMsgBox "Please specify the second Security Question.", vbExclamation, App.Title & " : Record not Selected", Me
        cboSecurityQuestion2.SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the User has not answered the User Name then...
    If txtSecurityQuestionAns(&H1).Text = VBA.vbNullString Then
        
        'Warn User to select an existing Record
        vMsgBox "Please enter the answer for the second Security Question.", vbExclamation, App.Title & " : Record not Selected", Me
        txtSecurityQuestionAns(&H1).SetFocus 'Move the focus to the specified control
        Exit Sub 'Quit this saving procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        'Retrieve the Record with the entered criteria from the specified table
        .Open "SELECT * FROM [Tbl_Users] WHERE [User Name] = '" & VBA.Replace(TxtUserName.Text, "'", "''") & "'", vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If a User account with the entered details exist then...
        If Not (.BOF And .EOF) Then
            
            .Filter = "[Security Question 1] = '" & VBA.Trim$(VBA.Replace(cboSecurityQuestion1.Text, "'", "''")) & "' AND [Security Ans 1] = '" & SmartEncrypt(VBA.Replace(VBA.Trim$(txtSecurityQuestionAns(&H0).Text), "'", "''")) & "' AND [Security Question 2] = '" & VBA.Trim$(VBA.Replace(cboSecurityQuestion2.Text, "'", "''")) & "' AND [Security Ans 2] = '" & SmartEncrypt(VBA.Replace(VBA.Trim$(txtSecurityQuestionAns(&H1).Text), "'", "''")) & "'"
            
            'If a User account with the entered details exist then...
            If Not (.BOF And .EOF) Then
                
                'Denote that the User has successfully verified the Account
                VerificationSucceeded = True
                
                If Not iConfirmingUser Then
                    
                    'Allow User to enter a new Account Password
                    User.User_ID = ![User ID]: User.User_Name = ![User Name]
                    If Not VBA.IsNull(![Password]) Then Frm_ChangePassword.TxtPassword(&H0).Text = SmartDecrypt(![Password])
                    Frm_ChangePassword.TxtPassword(&H0).Locked = True
                    CenterForm Frm_ChangePassword, Me
                    
                    'Initialize variables
                    User.User_ID = &H0: User.User_Name = VBA.vbNullString
                    
                End If 'Close respective IF..THEN block statement
                
                Unload Me 'Unload the Form from the Memory
                
            Else
                GoTo TrialCounter
            End If 'Close respective IF..THEN block statement
            
        Else 'If no User account exists with the entered details then...
            
TrialCounter:
            
            Static iTrials%
            
            iTrials = iTrials + &H1
            
            LblTrials.Caption = &H3 - iTrials & " Attempts"
            
            'If the number of Trials is 3 then...
            If iTrials < &H3 Then
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "The entered Security Information is incorrect.", vbExclamation, App.Title & " : Forgot Password", Me
                
            Else
                
                LblTrials.Visible = False
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "The maximum attempts has been reached. The Software will shut down in £ seconds..", vbExclamation, App.Title & " : Login Failed", Me, , , &HA
                
                vSilentClosure = True: Unload Me
                
                End 'Halt the Application
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        If .State = adStateOpen Then .Close  'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_ShpBttnOK_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_ShpBttnOK_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error - " & Err.Number, Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnOK_Click
    
End Sub
