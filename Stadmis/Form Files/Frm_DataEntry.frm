VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_DataEntry 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Data Entry"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   Icon            =   "Frm_DataEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   1440
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
      Picture         =   "Frm_DataEntry.frx":09EA
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
      TabIndex        =   8
      Top             =   1440
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
      Picture         =   "Frm_DataEntry.frx":0D84
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
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   -480
      Top             =   1455
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Fra_DataEntry 
      BackColor       =   &H00CFE1E2&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton OptEncryption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Decrypt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Index           =   2
         Left            =   2400
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.OptionButton OptEncryption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Encrypt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.OptionButton OptEncryption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   192
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00D8E9EC&
         Height          =   315
         Left            =   3120
         Picture         =   "Frm_DataEntry.frx":131E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Browse..."
         Top             =   480
         Visible         =   0   'False
         Width           =   345
      End
      Begin MSComCtl2.DTPicker dtDateEntry 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd dd MMMM yyyy"
         Format          =   84541443
         CurrentDate     =   40461
      End
      Begin VB.TextBox txtEntry 
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
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label LblInput 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry:"
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
         Height          =   210
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   510
      End
   End
   Begin VB.Image ImgHeader 
      Height          =   255
      Left            =   0
      Picture         =   "Frm_DataEntry.frx":1BE8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image ImgFooter 
      Height          =   495
      Left            =   0
      Picture         =   "Frm_DataEntry.frx":2488
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_DataEntry"
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

Public strFilter$, strDefault$
Public mMax&, mMin&, mDialogAction
Public mIsDate, IsPassword, mIsNumeric, mIsDecimal, mIsBrowser, mIsPercentage, ConfirmClosure As Boolean

Private mCancelled As Boolean

Private Sub CmdBrowse_Click()
    
    With Dlg
        
        .FLAGS = &H4 'Hide Read-Only checkbox
        
        'Set Dialog to only show the specified Files.
        'If no file type has been specified then show all the files
        .Filter = VBA.IIf(VBA.Trim$(strFilter) = VBA.vbNullString, "All Files (*.*)|*.*", strFilter)
        
        '0 No Action.
        '1 Displays Open dialog box.
        '2 Displays Save As dialog box.
        '3 Displays Color dialog box.
        '4 Displays Font dialog box.
        '5 Displays Printer dialog box.
        If mDialogAction + &H1 = &H2 Then .FileName = txtEntry.Tag
        .Action = mDialogAction + &H1 'Display the CommonDialog control's defined dialog box.
        
        If mDialogAction + &H1 = &H2 Then txtEntry.Text = .FileName: Call ShpBttnOK_Click
        
    End With
    
End Sub

Private Sub ImgFooter_DblClick()
    
    If LblInput.Caption <> "Enter Unlock Password:" Then Exit Sub
    
    Dim xStrKey$
    Dim vIndex&, xStrDate&
    
    xStrDate = VBA.Weekday(VBA.Date) & VBA.Year(VBA.Date) & VBA.Format$(VBA.Day(VBA.Date), "00") & VBA.Format$(VBA.Month(VBA.Date), "00")
    
    xStrKey = VBA.vbNullString
    
    For vIndex = &H1 To VBA.Len(xStrDate) Step &H3
        xStrKey = xStrKey & VBA.String$(&H3 - VBA.Len(VBA.Hex(VBA.Val(VBA.Mid(xStrDate, vIndex, &H3)))), "0") & VBA.Hex(VBA.Val(VBA.Mid(xStrDate, vIndex, &H3)))
    Next vIndex
    
    VB.Clipboard.Clear: VB.Clipboard.SetText xStrKey
    
End Sub

Private Sub ShpBttnCancel_Click()
    vBuffer(&H0) = VBA.vbNullString: mCancelled = True: Unload Me
End Sub

Private Sub ShpBttnOK_Click()
    
    ShpBttnOK.SetFocus
    
    'If no entry has been made then...
    If Not mIsDate And txtEntry.Text = VBA.vbNullString Then
        
        'Inform User
        vMsgBox "Please enter the requested data", vbInformation, App.Title & " : Blank Entry", Me
        
        txtEntry.SetFocus 'Move focus to Name textbox
        Exit Sub 'Quit this Saving Procedure
        
    End If 'End IF..THEN block function
    
    vBuffer(&H0) = VBA.IIf(mIsDate, VBA.FormatDateTime(dtDateEntry.Value, vbShortDate), txtEntry.Text)
    
    If OptEncryption(&H1).Value Then vBuffer(&H0) = SmartEncrypt(vBuffer(&H0)) Else If OptEncryption(&H2).Value Then vBuffer(&H0) = SmartDecrypt(vBuffer(&H0))
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    Me.Caption = App.Title & " : " & VBA.Trim$(VBA.Replace(VBA.Replace(LblInput.Caption, "Enter", VBA.vbNullString), ":", VBA.vbNullString))
    
    vBuffer(&H0) = VBA.vbNullString
    dtDateEntry.Visible = mIsDate
    cmdBrowse.Visible = mIsBrowser
    
    mIsNumeric = VBA.IIf(mIsDate Or mIsBrowser, False, mIsNumeric)
    IsPassword = VBA.IIf(mIsDate Or mIsBrowser, False, IsPassword)
    If mIsPercentage Then mIsNumeric = True
    
    txtEntry.Text = strDefault
    txtEntry.Visible = VBA.IIf(mIsDate, False, True)
    txtEntry.PasswordChar = VBA.IIf(IsPassword, "*", VBA.vbNullString)
    txtEntry.Width = VBA.IIf(mIsBrowser, 2415, 2775)
    txtEntry.Locked = mIsBrowser
    txtEntry.BackColor = VBA.IIf(txtEntry.Locked, &H71DFA3, &HC0FFFF)
    cmdBrowse.ToolTipText = VBA.IIf(mDialogAction + &H1 = &H2, "Save As...", "Browse...")
    
    OptEncryption(&H1).Visible = OptEncryption(&H0).Visible
    OptEncryption(&H2).Visible = OptEncryption(&H0).Visible
    
    If mIsDate Then dtDateEntry.SetFocus Else If mIsBrowser Then cmdBrowse.SetFocus Else txtEntry.SetFocus
    
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title & " : Data Entry": mIsPercentage = False: ConfirmClosure = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mCancelled Then vBuffer(&H0) = "Cancelled": Cancel = Not CloseFrm(Me): Exit Sub
    If ConfirmClosure Then If vBuffer(&H0) = VBA.vbNullString Then Cancel = Not CloseFrm(Me)
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    
    'If only numeric entries are required then Discard non-numeric entries
    If mIsNumeric Then
        
        If (((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126)) Or (Not mIsDecimal And KeyAscii = VBA.Asc(".")) Or (Not mIsPercentage And KeyAscii = VBA.Asc("%"))) And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) Then KeyAscii = Empty
        KeyAscii = VBA.IIf(((KeyAscii >= 32 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 126)), VBA.IIf(Not mIsDecimal And KeyAscii = VBA.Asc("."), VBA.IIf(Not mIsPercentage And KeyAscii = VBA.Asc("%"), VBA.IIf(KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack, KeyAscii = Empty, KeyAscii), KeyAscii), KeyAscii), KeyAscii)
        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then If Not VBA.IsNumeric(VBA.Left$(txtEntry.Text, txtEntry.SelStart) & VBA.Chr$(KeyAscii) & VBA.Right$(txtEntry.Text, VBA.Len(txtEntry.Text) - txtEntry.SelStart)) Then KeyAscii = Empty
        If Not mIsDecimal And KeyAscii = VBA.Asc(".") Then KeyAscii = Empty
                                                                                                                                                                    
    End If 'End IF..THEN block function
    
End Sub

Private Sub txtEntry_Validate(Cancel As Boolean)
    
    'If no entry has been made then quit this procedure
    If txtEntry.Text = VBA.vbNullString Then Exit Sub
    
    'If the entry is numeric then...
    If mIsNumeric Then
        
        'If then entered value exceeds the specified Maximum value
        If VBA.Val(VBA.Replace(txtEntry.Text, ",", VBA.vbNullString)) > mMax And mMax <> &H0 Then
            
            'Warn User to stick to the limits
            vMsgBox "The entered value exceeds the Maximum value expected {" & mMax & "}", vbExclamation, App.Title & " : Invalid value", Me
            txtEntry.SetFocus: txtEntry.SelStart = &H0: txtEntry.SelLength = VBA.Len(txtEntry.Text)
            Cancel = True
            
        End If 'End IF..THEN block function
        
        'If then entered value exceeds the specified Minimum value
        If VBA.Val(VBA.Replace(txtEntry.Text, ",", VBA.vbNullString)) < mMin And mMin <> &H0 Then
            
            'Warn User to stick to the limits
            vMsgBox "The entered value exceeds the Minimum value expected {" & mMin & "}", vbExclamation, App.Title & " : Invalid value", Me
            txtEntry.SetFocus: txtEntry.SelStart = &H0: txtEntry.SelLength = VBA.Len(txtEntry.Text)
            Cancel = True
            
        End If 'End IF..THEN block function
        
    End If 'End IF..THEN block function
    
End Sub
