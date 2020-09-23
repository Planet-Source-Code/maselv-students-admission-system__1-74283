VERSION 5.00
Begin VB.Form Frm_SoftwarePatent 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Software Protection"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "Frm_SoftwarePatent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_Login 
      BackColor       =   &H00CFE1E2&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   5895
      Begin VB.TextBox txtUsers 
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtSerialCode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   4920
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSerialCode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSerialCode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSerialCode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1335
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSerialCode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   120
         MaxLength       =   5
         TabIndex        =   0
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSerialKey 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtDeviceSerial 
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtLicenseCode 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   372
      End
      Begin VB.TextBox txtExpiry 
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Users:"
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
         Index           =   5
         Left            =   4320
         TabIndex        =   23
         Top             =   285
         Width           =   525
      End
      Begin VB.Line Line 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Index           =   3
         X1              =   4680
         X2              =   4800
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Index           =   2
         X1              =   3480
         X2              =   3600
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Index           =   1
         X1              =   2295
         X2              =   2415
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Index           =   0
         X1              =   1080
         X2              =   1215
         Y1              =   1455
         Y2              =   1470
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Code:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Key:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device Serial:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "License Code:"
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
         Index           =   3
         Left            =   2040
         TabIndex        =   13
         Top             =   705
         Width           =   1140
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry:"
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
         Index           =   4
         Left            =   3720
         TabIndex        =   12
         Top             =   705
         Width           =   570
      End
   End
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   6
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
      Caption         =   "      Cancel"
      AccessKey       =   "C"
      Picture         =   "Frm_SoftwarePatent.frx":0EE2
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
      Left            =   5040
      TabIndex        =   5
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
      Caption         =   "    OK"
      AccessKey       =   "O"
      Picture         =   "Frm_SoftwarePatent.frx":127C
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
   Begin VB.Label lblDeveloper 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail: customerinfo@lexeme-kenya.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblDeveloper 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone: (+254) 724 688 172 / (+254) 724 585 279"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_SoftwarePatent.frx":1816
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First time running"
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
      TabIndex        =   18
      Top             =   120
      Width           =   1695
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
      TabIndex        =   17
      Top             =   3120
      Width           =   960
   End
   Begin VB.Image ImgHeader 
      Height          =   1215
      Left            =   0
      Picture         =   "Frm_SoftwarePatent.frx":18C1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6495
   End
   Begin VB.Image ImgFooter 
      Height          =   855
      Left            =   0
      Picture         =   "Frm_SoftwarePatent.frx":20B7
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   6495
   End
End
Attribute VB_Name = "Frm_SoftwarePatent"
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


'****************************************************************************************************
'*                                  COPY PROTECTION                                                 *
'****************************************************************************************************
'*  This is a software “lock” placed on the program by the developer to prevent the product from    *
'*  being copied and distributed without approval or authorization.                                 *
'****************************************************************************************************

Option Explicit
Option Compare Binary

Private Trials%
Private myArray() As String
Private IsLoading As Boolean
Private CopyrightVerified As Boolean
Private nBuffer$, Device_Serial_No$
Private vFso As New FileSystemObject

Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Function GetWindowsDir() As String
    
    Dim ret As Long
    Dim Temp As String
    
    Const MAX_LENGTH = 145

    Temp = VBA.String$(MAX_LENGTH, &H0)
    ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = VBA.Left$(Temp, ret)
    GetWindowsDir = VBA.IIf(Temp <> VBA.vbNullString And VBA.Right$(Temp, &H1) <> "\", Temp & "\", Temp)
    
End Function

Private Function EncodeSerial(vDate As Date, vWeekday%, vSerialKey&, vUsers&) As String
On Local Error GoTo Handle_EncodeSerial_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim nIndex&
    Dim nBuffer$, nBuffer1$
    
    'Users 00: Day 00: Month 00: Serial 0000: Year 00
    
    VBA.Randomize 'Initializes the random-number generator.
    
    'Txt(&H0).Text = VBA.Format$(vUsers, "00") & VBA.Format$(VBA.Day(vDate), "00") & VBA.Format$(VBA.Month(VBA.Date), "00") & vSerialKey & VBA.Right(VBA.Year(VBA.Date), &H2)
    nBuffer1 = VBA.StrReverse(VBA.Format$(vUsers, "00")) & VBA.StrReverse(VBA.Format$(VBA.Day(vDate), "00")) & VBA.StrReverse(VBA.Format$(VBA.Month(vDate), "00")) & VBA.StrReverse(vSerialKey) & VBA.StrReverse(VBA.Right(VBA.Year(vDate), &H2))
    
    'For each character in the provided Serial Code
    For nIndex = &H1 To VBA.Len(nBuffer1) Step &H1
        
        'Returns an Integer representing the code for each character
        nBuffer = nBuffer & VBA.Asc(VBA.Mid$(nBuffer1, nIndex, &H1))
        
    Next nIndex 'Increment counter variable by value in the STEP option
    
    'Attach the System's weekday to the generated value
    nBuffer = VBA.StrReverse(nBuffer) & vWeekday
    
    nBuffer1 = VBA.vbNullString 'Initialize variable
    
    'For each character in the generated value
    For nIndex = &H1 To VBA.Len(nBuffer) Step &H1
        
        'Return a String containing the character associated with the each character code
        nBuffer1 = nBuffer1 & VBA.Chr$(65 + (VBA.Mid$(nBuffer, nIndex, &H1)))
        
    Next nIndex 'Increment counter variable by value in the STEP option
    
    'Reverse the character order of the specified string
    nBuffer1 = VBA.StrReverse(nBuffer1)
    
    nBuffer = VBA.vbNullString 'Initialize variable
    
    'For each character in the generated value
    For nIndex = &H1 To VBA.Len(nBuffer1) Step &H5
        
        'Split the value into 5 characters separated by -
        nBuffer = nBuffer & VBA.Mid$(nBuffer1, nIndex, &H5) & "-"
        
    Next nIndex 'Increment counter variable by value in the STEP option
    
    'Assign the final generated value to this Function
    EncodeSerial = VBA.Left(nBuffer, VBA.Len(nBuffer) - &H1)
    
Exit_EncodeSerial:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Function
    
Handle_EncodeSerial_Error:
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Encoding - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_EncodeSerial
    
End Function

Private Function DecodeSerial(SerialCode$) As String
On Local Error GoTo Handle_DecodeSerial_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim nIndex&
    Dim nBuffer$, nBuffer1$
    Dim vWeekday&, vDay&, vMonth&, vYear&, vSerial&, vUsers&
    
    'Reverse the character order of the specified string and combine characters in the value by removing by -
    nBuffer = VBA.StrReverse(VBA.Replace(SerialCode, "-", VBA.vbNullString))
    
    nBuffer1 = VBA.vbNullString 'Initialize variable
    
    'For each character in the specified value
    For nIndex = &H1 To VBA.Len(nBuffer) Step &H1
        
        'Return a String containing the character codes associated with the each character in the value
        nBuffer1 = nBuffer1 & (VBA.Asc((VBA.Mid$(nBuffer, nIndex, &H1))) - 65)
        
    Next nIndex 'Increment counter variable by value in the STEP option
    
    vWeekday = VBA.Right$(nBuffer1, &H1) 'Extract the System's Weekday from the regenerated Serial Code
    
    'Reverse the character order of the regenerated Serial Code without the System's Weekday
    nBuffer1 = VBA.StrReverse(VBA.Left$(nBuffer1, VBA.Len(nBuffer1) - &H1))
    
    nBuffer = VBA.vbNullString 'Initialize variable
    
    'For each character in the specified value
    For nIndex = &H1 To VBA.Len(nBuffer1) Step &H2
        
        'Return a String containing the character associated with the each double-character code
        nBuffer = nBuffer & VBA.Chr$(VBA.Mid$(nBuffer1, nIndex, &H2))
        
    Next nIndex 'Increment counter variable by value in the STEP option
    
    'Assign the final generated value to this Function
    DecodeSerial = VBA.StrReverse(VBA.Left(nBuffer, &H2)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(nBuffer, &H3, &H2)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(nBuffer, &H5, &H2)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(nBuffer, &H7, &H4)) & "|" & _
                    VBA.StrReverse(VBA.Mid$(nBuffer, &HB, &H2)) & "|" & vWeekday
    
Exit_DecodeSerial:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Function 'Quit this Function
    
Handle_DecodeSerial_Error:
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Decoding - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_DecodeSerial
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    vRegistering = False
End Sub

Private Sub ShpBttnCancel_Click()
    Unload Me 'Unload this Form from the Memory
End Sub

Private Sub ShpBttnOK_Click()
On Local Error GoTo Handle_ShpBttnOK_Click_Error
    
    'If the Serial Code has not been entered then...
    If VBA.LenB(VBA.Trim$(txtSerialCode(&H0).Text)) = &H0 Then
        
        'Inform User
        vMsgBox "Please enter Serial Code", vbInformation, App.Title & " : Unspecified Serial Code", Me
        txtSerialCode(&H0).SetFocus 'Move focus to the specified control
        Exit Sub 'Quit this Sub-Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim nIndex&
    Dim sSerialKey$, sSerialCode$
    
    For nIndex = &H0 To &H4 Step &H1
        sSerialCode = sSerialCode & txtSerialCode(nIndex).Text & "-"
    Next nIndex
    sSerialCode = VBA.Left$(sSerialCode, VBA.Len(sSerialCode) - &H1)
    
    sSerialKey = DecodeSerial(sSerialCode)
    
    If VBA.LenB(VBA.Trim$(sSerialKey)) = &H0 Then GoTo VerificationFailed
    
    Dim sArrayList() As String
    
    sArrayList = VBA.Split(sSerialKey, "|")
    
    If UBound(sArrayList) < &H5 Then GoTo VerificationFailed
    
    If Not VBA.IsDate(VBA.DateSerial(sArrayList(&H4), sArrayList(&H2), sArrayList(&H1))) Then GoTo VerificationFailed Else txtExpiry.Tag = VBA.DateSerial(sArrayList(&H4), sArrayList(&H2), sArrayList(&H1))
    If Not VBA.IsNumeric(sArrayList(&H0)) Then GoTo VerificationFailed Else txtLicenseCode.Tag = sArrayList(&H0) 'Max Users
    If Not VBA.IsNumeric(sArrayList(&H3)) Then GoTo VerificationFailed
    
    If sArrayList(&H5) <> VBA.Weekday(VBA.Date) Then GoTo VerificationFailed
    If sArrayList(&H3) <> VBA.Val(txtSerialKey.Text) Then GoTo VerificationFailed
    If sArrayList(&H5) <> VBA.Weekday(VBA.Date) Then GoTo VerificationFailed
    
    vRegistered = (txtExpiry.Tag = VBA.DateSerial(&H7C2, &H3, &H9))
    
    'If the entered Licence period has already expired then...
    If VBA.DateDiff("d", txtExpiry.Tag, VBA.Date) > &H0 And Not vRegistered Then
        
        'Warn User
        vMsgBox "The entered License expired on " & VBA.Format$(txtExpiry.Tag, "ddd dd MMM yyyy"), vbExclamation, App.Title & " : Invalid Licence", Me
        GoTo VerificationFailed
        
    End If 'Close respective IF..THEN block statement
    
    txtExpiry.Text = VBA.Format$(txtExpiry.Tag, "ddd dd MMM yyyy")
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    VB.Clipboard.Clear 'Clear clipboard contents
    
    VBA.SaveSetting App.Title, "Copyright Protection", "License Code", SmartEncrypt(VBA.StrReverse(VBA.Format$(txtLicenseCode.Text, "00")), False)
    SoftwareSetting.Licences.License_Encrypted = SmartEncrypt(sSerialCode)
    VBA.SaveSetting App.Title, "Copyright Protection", "License Encrypted", SmartEncrypt(sSerialCode & "|" & txtDeviceSerial.Text)
    VBA.SaveSetting App.Title, "Copyright Protection", "Device Serial", SmartEncrypt(SmartEncrypt(VBA.StrReverse(txtSerialKey.Text), False) & "-" & VBA.StrReverse(VBA.Format$(VBA.Replace(txtDeviceSerial.Text, "-", VBA.vbNullString), "000000000000")))
    
    CopyrightVerified = True 'Denote that the program running has been verified
    
    'Continue with loading the program
    SoftwareSetting.Licences.Expiry_Date = txtExpiry.Tag
    SoftwareSetting.Licences.Key = txtSerialKey.Text
    SoftwareSetting.Licences.License_Code = sSerialCode
    SoftwareSetting.Licences.Max_Users = VBA.Val(txtLicenseCode.Tag) '* &HA
    txtUsers.Text = txtLicenseCode.Tag
    
    'If not registering on demand then...
    If Not vRegistering Then
        
        'Inform User
        vMsgBox "The Serial Code has successfully been verified. Proceeding with loading.", vbInformation, App.Title & " : Software Copyright", Me
        If SoftwareSetting.Show_Splash_Screen And Not vRegistered Then Frm_Splash.Show Else Call WindowsMinimizeAll: Frm_Main.Show
        
    End If 'Close respective IF..THEN block statement
    
    Unload Me: GoTo Exit_ShpBttnOK_Click
    
VerificationFailed:
    
    'Increment the number of Password trials by 1
    Trials = Trials + &H1
    
    LblTrials.Caption = &H3 - Trials & " Attempts"
    
    'If the entered Licence period has already expired then...
    If VBA.DateDiff("d", txtExpiry.Tag, VBA.Date) > &H0 And Not vRegistered Then
        'Do nothing;User already warned
    Else
        
        'Warn User
        vMsgBox "Invalid Serial Code. Please Retry", vbExclamation, App.Title & " : Software Copyright", Me
        
    End If 'Close respective IF..THEN block statement
    
    'Highlight the entered Serial Code contents
    txtSerialCode(&H0).SetFocus 'Move focus to the specified control
    txtSerialCode(&H0).SelStart = &H0: txtSerialCode(&H0).SelLength = VBA.LenB(txtSerialCode(&H0).Text)
    
    'If the number of Trials is 3 then...
    If Trials = &H3 Then
        
        LblTrials.Visible = False 'Hide the control
        
        'Inform User
        vMsgBox "The maximum attempts has been reached. The Software will shut down.", vbExclamation, App.Title & " : Login Failed", Me
        
        'Minimize hacking by deleting the currently entered copyright settings
        If Not vRegistering Then VBA.DeleteSetting App.Title, "Copyright Protection"
        
        vRegistered = False: Unload Me 'Halt the Application
        
        Exit Sub 'Quit this Sub-Procedure
        
    End If 'Close respective IF..THEN block statement
    
Exit_ShpBttnOK_Click:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_ShpBttnOK_Click_Error:
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Activating Form - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnOK_Click
    
End Sub

Private Sub Form_Activate()
On Local Error GoTo Handle_Form_Activate_Error
    
    If Not IsLoading Then Exit Sub 'If the Form has already loaded then Quit this Procedure
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    IsLoading = False
    
    'If the program has been copyrighted then continue with loading it
    If CopyrightVerified Then
        
        vRegistered = (txtExpiry.Tag = VBA.DateSerial(&H7C2, &H3, &H9))
        
        If VBA.DateDiff("d", VBA.Date, txtExpiry.Tag) < &H1 And Not vRegistered Then
            CopyrightVerified = False: lblHeading.Caption = "Licence Expired"
            lblInfo(&H1).Caption = "The License period expired" & VBA.IIf(txtExpiry.Tag = VBA.Date, " Today", " on " & VBA.Format$(txtExpiry.Tag, "ddd dd MMM yyyy")) & " at 12:00:00 AM. Please contact the Software Developer for renewal of the Licence."
            VBA.DoEvents
        Else
            
            Dim nIndex&
            Dim sSerialCode$
            
            sSerialCode = VBA.vbNullString
            For nIndex = &H0 To &H4 Step &H1
                sSerialCode = sSerialCode & txtSerialCode(nIndex).Text & "-"
            Next nIndex
            sSerialCode = VBA.Left$(sSerialCode, VBA.Len(sSerialCode) - &H1)
            
            'Continue with loading the program
            SoftwareSetting.Licences.Expiry_Date = txtExpiry.Tag
            SoftwareSetting.Licences.Key = txtSerialKey.Text
            SoftwareSetting.Licences.License_Code = sSerialCode
            SoftwareSetting.Licences.Max_Users = VBA.Val(txtLicenseCode.Tag) '* &HA
            
            If Not vRegistered And Not vRegistering Then
                
                Call GetSoftwareSettings(True)
                If SoftwareSetting.Show_Splash_Screen Then Frm_Splash.Show Else Call WindowsMinimizeAll: Frm_Main.Show
                Unload Me: Exit Sub
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    Me.Show 'Display this Form to the User
    
Exit_Form_Activate:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Form_Activate_Error:
    
    If Err.Number = 401 Then Resume Next
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Activating Form - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_Form_Activate
    
End Sub

Private Sub Form_Load()
On Local Error GoTo Handle_Form_Load_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    IsLoading = True: vSilentClosure = False
    
    Me.Caption = App.Title & " : Software Protection"
    lblInfo(&H1).Caption = App.Title & " running on a limited Licence Period. Please request for Licence before expiration of the current one."
    
    'VBA.DeleteSetting App.Title, "Copyright Protection"
    
    Dim vDrive
    Dim sTrials%
    Dim vString$
    Dim ObjNet As Object
    
    Set ObjNet = CreateObject("WScript.Network")
    Set vDrive = vFso.GetDrive(vFso.GetDriveName(GetWindowsDir))
    
    Device.Name = ObjNet.ComputerName
    Device.Account_Name = ObjNet.UserName
    Device.Serial_No = vDrive.SerialNumber
    txtDeviceSerial.Tag = Device.Serial_No
    Set ObjNet = Nothing: Set vDrive = Nothing
    
    Device_Serial_No = VBA.Format$(VBA.Replace(txtDeviceSerial.Tag, "-", VBA.vbNullString), "000000000000")
    
    vString = VBA.GetSetting(App.Title, "Copyright Protection", "Device Serial", "0")
    
    If vString <> "0" Then
        
        myArray = VBA.Split(SmartDecrypt(vString), "-")
    
        'If the program is running for the first time on the Computer then...
        If UBound(myArray) < &H0 Then sTrials = &H3: GoTo RequestPwd
        
    Else
        sTrials = &H3: GoTo RequestPwd
    End If 'Close respective IF..THEN block statement
        
    'If the program is running for the first time on the Computer then...
    If VBA.StrReverse(VBA.Replace(myArray(&H1), VBA.vbCr, VBA.vbNullString)) <> Device_Serial_No Then
        
        lblHeading.Caption = App.Title & " has detected that it is running on this device for the first time. Due to copyright protection, please fill appropriately the below information to unlock it."
        
        Dim sPwd$
        Dim sArray As Variant
        
        If vString <> "0" Then VBA.DeleteSetting App.Title, "Copyright Protection"
        sArray = VBA.GetAllSettings(App.Title, "Main Form")
        If UBound(sArray) >= &H0 Then VBA.DeleteSetting App.Title, "Main Form"
        sArray = VBA.GetAllSettings(App.Title, "Settings")
        If UBound(sArray) >= &H0 Then VBA.DeleteSetting App.Title, "Settings"
        
        sTrials = &H3
        
RequestPwd:
        
        Load Frm_DataEntry
        
        vBuffer(&H0) = VBA.vbNullString
        Frm_DataEntry.Caption = App.Title & " : First Time Use"
        Frm_DataEntry.OptEncryption(&H0).Visible = True
        Frm_DataEntry.IsPassword = True
        Frm_DataEntry.LblInput.Caption = "Enter Unlock Password:"
        Call CenterForm(Frm_DataEntry)
        
        'If the User has cancelled then confirm halting the Application
        If vBuffer(&H0) = "Cancelled" Or vBuffer(&H0) = VBA.vbNullString Then vMsgBox "The Application will automatically terminate after £ seconds!!", vbExclamation, App.Title & " : Unspecified Unlock Password", , , , &H8: GoTo QuitApp
        
        'IF Sun, Tue, Thu & Sat THEN Decrypt entry ELSE IF Mon, Wed, Fri THEN Encrypt entry
        sPwd = SmartDecrypt(vBuffer(&H0), True)
        
        Dim xStrKey$
        Dim vIndex&, xStrDate&, xStrCode&
        
        xStrDate = VBA.Date
        xStrCode = VBA.Weekday(xStrDate) & VBA.Year(xStrDate) & VBA.Format$(VBA.Day(xStrDate), "00") & VBA.Format$(VBA.Month(xStrDate), "00")
        
        xStrKey = VBA.vbNullString
        
        For vIndex = &H1 To VBA.Len(xStrCode) Step &H3
            xStrKey = xStrKey & VBA.String$(&H3 - VBA.Len(VBA.Hex(VBA.Val(VBA.Mid(xStrCode, vIndex, &H3)))), "0") & VBA.Hex(VBA.Val(VBA.Mid(xStrCode, vIndex, &H3)))
        Next vIndex
        
        If xStrKey <> sPwd Then
            
            'If the maximum number of trials has not been reached then...
            If sTrials >= &H1 Then
                
                'Warn User
                vMsgBox "Invalid Password. Remaining " & sTrials & ".", vbExclamation, App.Title & " : Invalid Password"
                sTrials = sTrials - &H1: GoTo RequestPwd 'Request for entry again
                
            Else 'If the maximum number of trials has been reached then...
                
                'Inform User
                vMsgBox "The maximum attempts has been reached. The Software will shut down.", vbExclamation, App.Title & " : Verification Failed"
QuitApp:
                End 'Halt the Application
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        VBA.Randomize 'Initialize the random-number generator.
        
        'Generate a random number between 1000 and 9999
        vBuffer(&H0) = VBA.StrReverse(VBA.Int((1000 - 9999 + 1000) * VBA.Rnd + 9999))
        
        Dim iStr(&H3) As String
        
        iStr(&H0) = EncodeSerial(VBA.DateAdd("d", 31, VBA.Date), VBA.Weekday(VBA.Date), VBA.Val(vBuffer(&H0)), &H1)
        
        VBA.SaveSetting App.Title, "Copyright Protection", "License Code", SmartEncrypt(VBA.StrReverse(VBA.Format$(VBA.Weekday(VBA.Date), "00")), False)
        iStr(&H1) = SmartEncrypt(iStr(&H0) & "|" & txtDeviceSerial.Tag)
        VBA.SaveSetting App.Title, "Copyright Protection", "License Encrypted", iStr(&H1)
        
        'Assign the serial key of the Computer
        VBA.SaveSetting App.Title, "Copyright Protection", "Device Serial", SmartEncrypt(SmartEncrypt(VBA.StrReverse(vBuffer(&H0)), False) & "-" & VBA.StrReverse(VBA.Format$(VBA.Replace(Device_Serial_No, "-", VBA.vbNullString), "000000000000")))
        
        txtExpiry.Tag = VBA.DateAdd("d", 31, VBA.Date)
        txtExpiry.Text = VBA.Format$(txtExpiry.Tag, "ddd dd MMM yyyy")
        
    End If 'Close respective IF..THEN block statement
    
    myArray = VBA.Split(SmartDecrypt(VBA.GetSetting(App.Title, "Copyright Protection", "Device Serial", VBA.vbNullString)), "-")
    
    txtDeviceSerial.Text = txtDeviceSerial.Tag
    txtSerialKey.Text = VBA.StrReverse(SmartDecrypt(myArray(&H0), False))
    
    If VBA.LenB(VBA.Trim$(txtSerialKey.Text)) = &H0 Then
        
        VBA.Randomize 'Initialize the random-number generator.
        vBuffer(&H0) = VBA.StrReverse(VBA.Int((1000 - 9999 + 1000) * VBA.Rnd + 9999))
        iStr(&H0) = EncodeSerial(VBA.DateAdd("d", &H1, VBA.Date), VBA.Weekday(VBA.Date), VBA.Val(vBuffer(&H0)), 30)
        VBA.SaveSetting App.Title, "Copyright Protection", "License Code", SmartEncrypt(VBA.StrReverse(VBA.Format$(VBA.Weekday(VBA.Date), "00")), False)
        iStr(&H1) = SmartEncrypt(iStr(&H0) & "|" & txtDeviceSerial.Tag)
        VBA.SaveSetting App.Title, "Copyright Protection", "License Encrypted", iStr(&H1)
        
        'Assign the serial key of the Computer
        VBA.SaveSetting App.Title, "Copyright Protection", "Device Serial", SmartEncrypt(SmartEncrypt(VBA.StrReverse(vBuffer(&H0)), False) & "-" & VBA.StrReverse(VBA.Format$(VBA.Replace(Device_Serial_No, "-", VBA.vbNullString), "000000000000")))
        
        myArray = VBA.Split(SmartDecrypt(VBA.GetSetting(App.Title, "Copyright Protection", "Device Serial", VBA.vbNullString)), "-")
        
    End If 'Close respective IF..THEN block statement
    
    txtDeviceSerial.Text = txtDeviceSerial.Tag
    txtSerialKey.Text = VBA.StrReverse(SmartDecrypt(myArray(&H0), False))
    
    txtLicenseCode.Text = VBA.Format$(VBA.Weekday(VBA.Date), "00")
    
    Dim sSerialCode$, sSerialKey$
    
    'Validate existing License Code if it exists
    sSerialCode = SmartDecrypt(VBA.GetSetting(App.Title, "Copyright Protection", "License Encrypted", VBA.vbNullString))
    
    If VBA.LenB(VBA.Trim$(sSerialCode)) = &H0 Then GoTo ActivateForm
    
    'If the existing code is not up to to required number of characters then...
    If VBA.Len(VBA.Replace(VBA.Replace(sSerialCode, "|-", "|"), "|" & VBA.Replace(txtDeviceSerial.Tag, "-", VBA.vbNullString), VBA.vbNullString)) < 25 Then GoTo FakeSerialCode
    
    txtLicenseCode.Text = VBA.StrReverse(SmartDecrypt(VBA.GetSetting(App.Title, "Copyright Protection", "License Code", VBA.vbNullString), False))
    
    sSerialKey = DecodeSerial(VBA.Replace(VBA.Replace(sSerialCode, "|-", "|"), "|" & VBA.Replace(txtDeviceSerial.Tag, "-", VBA.vbNullString), VBA.vbNullString))
    
    myArray = VBA.Split(sSerialCode, "|")
    If UBound(myArray) <> &H1 Then GoTo FakeSerialCode
    If VBA.Replace(myArray(&H1), "-", VBA.vbNullString) <> VBA.Replace(txtDeviceSerial.Tag, "-", VBA.vbNullString) Then GoTo FakeSerialCode
    
    myArray = VBA.Split(sSerialKey, "|")
    
    If VBA.LenB(VBA.Trim$(txtLicenseCode.Text)) = &H0 Then GoTo ActivateForm
    If UBound(myArray) < &H5 Then GoTo FakeSerialCode
    
    If Not VBA.IsDate(VBA.DateSerial(myArray(&H4), myArray(&H2), myArray(&H1))) Then GoTo FakeSerialCode Else txtExpiry.Tag = VBA.DateSerial(myArray(&H4), myArray(&H2), myArray(&H1))
    If Not VBA.IsNumeric(myArray(&H0)) Then GoTo FakeSerialCode Else txtLicenseCode.Tag = myArray(&H0) 'Max Users
    If Not VBA.IsNumeric(myArray(&H3)) Then GoTo FakeSerialCode
    
    txtUsers.Text = txtLicenseCode.Tag
    txtExpiry.Text = VBA.Format$(txtExpiry.Tag, "ddd dd MMM yyyy")
    
    If VBA.Val(myArray(&H5)) <> VBA.Val(txtLicenseCode.Text) Then GoTo ActivateForm
    If VBA.Val(myArray(&H3)) <> VBA.Val(txtSerialKey.Text) Then GoTo ActivateForm
    txtLicenseCode.Text = VBA.Format$(myArray(&H5), "00")
    
    Dim nIndex&
    Dim nArray() As String
    
    nArray = VBA.Split(VBA.Replace(VBA.Replace(sSerialCode, "|-", "|"), "|" & VBA.Replace(txtDeviceSerial.Tag, "-", VBA.vbNullString), VBA.vbNullString), "-")
    
    For nIndex = &H0 To &H4 Step &H1
        txtSerialCode(nIndex).Text = nArray(nIndex)
    Next nIndex
    
    CopyrightVerified = True 'Denote that the program running has been verified
    
ActivateForm:
    
    Me.Hide
    Call Form_Activate
    GoTo Exit_Form_Load
    
FakeSerialCode:
    
    'If the Software has been licensed with a fake serial code then...
    vMsgBox "This Software has been licensed with a Fake License Code. The Licence will be revoked. Please contact the Software Administrator.", vbCritical, App.Title & " : Fake Licence"
    
    VBA.DeleteSetting App.Title, "Copyright Protection"
    
    End 'Close Program and Quit this Procedure
    
Exit_Form_Load:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Form_Load_Error:
    
    If Err.Number = &H5 Then Resume Next
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Loading Form - " & Err.Number
    
    'Resume execution at the specified Label
    Resume Exit_Form_Load
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If the program running has not successfully been copyrighted then confirm application exit
    If Not CopyrightVerified Then Cancel = Not CloseFrm(Me, , False): If Not Cancel And Not vRegistered Then End
End Sub

Private Sub lblInfo_DblClick(Index As Integer)
On Local Error Resume Next
    'Revoke all Copyright protection settings of the device
    VBA.DeleteSetting App.Title, "Copyright Protection"
End Sub

Private Sub txtDeviceSerial_GotFocus()
On Local Error GoTo Handle_txtDeviceSerial_GotFocus_Error
    
    If vRegistering Then Exit Sub
    
    Static iGotFocus As Boolean
    
    Dim nIndex&
    Dim sSerialCode$, nSerialCode$, sSerialKey$
    
    If iGotFocus Then Exit Sub
    
    iGotFocus = True
    
    nSerialCode = VBA.vbNullString
    For nIndex = &H0 To &H4 Step &H1
        nSerialCode = nSerialCode & txtSerialCode(nIndex).Text & "-"
    Next nIndex
    
    nSerialCode = VBA.Replace(VBA.Left$(nSerialCode, VBA.LenB(nSerialCode) - &H1), "----", VBA.vbNullString)
    
    'Validate existing License Code if it exists
    sSerialCode = SmartDecrypt(VBA.GetSetting(App.Title, "Copyright Protection", "License Encrypted", VBA.vbNullString))
    
    If sSerialCode = nSerialCode Then GoTo Exit_txtDeviceSerial_GotFocus
    
    If VBA.LenB(VBA.Trim$(sSerialCode)) = &H0 Then GoTo ActivateForm
    
    'If the existing code is not up to to required number of characters then...
    If VBA.LenB(VBA.Replace(VBA.Replace(sSerialCode, "|" & txtDeviceSerial.Tag, VBA.vbNullString), "-", VBA.vbNullString)) < 25 Then GoTo FakeSerialCode
    
    txtLicenseCode.Text = VBA.StrReverse(SmartDecrypt(VBA.GetSetting(App.Title, "Copyright Protection", "License Code", VBA.vbNullString), False))
    
    sSerialKey = DecodeSerial(VBA.Replace(sSerialCode, "|" & VBA.Replace(txtDeviceSerial.Tag, "-", VBA.vbNullString), VBA.vbNullString))
    
    myArray = VBA.Split(sSerialKey, "|")
    
    If txtLicenseCode.Text = VBA.vbNullString Then GoTo ActivateForm
    If UBound(myArray) < &H5 Then GoTo FakeSerialCode
    
    If Not VBA.IsDate(VBA.DateSerial(myArray(&H4), myArray(&H2), myArray(&H1))) Then GoTo FakeSerialCode Else txtExpiry.Tag = VBA.DateSerial(myArray(&H4), myArray(&H2), myArray(&H1))
    If Not VBA.IsNumeric(myArray(&H0)) Then GoTo FakeSerialCode Else txtLicenseCode.Tag = myArray(&H0) 'Max Users
    If Not VBA.IsNumeric(myArray(&H3)) Then GoTo FakeSerialCode
    
    If VBA.Val(myArray(&H5)) <> VBA.Val(txtLicenseCode.Text) Then GoTo ActivateForm
    If VBA.Val(myArray(&H3)) <> VBA.Val(txtSerialKey.Text) Then GoTo ActivateForm
    
    txtExpiry.Text = VBA.Format$(txtExpiry.Tag, "ddd dd MMM yyyy")
    
    Dim nArray() As String
    
    nArray = VBA.Split(sSerialCode, "-")
    
    For nIndex = &H0 To &H4 Step &H1
        txtSerialCode(nIndex).Text = nArray(nIndex)
    Next nIndex
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    txtLicenseCode.SetFocus
    
    vMsgBox "A new Licence has been detected. Proceeding with Licence verification", vbInformation, App.Title & " : Licence Renewal", Me
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    CopyrightVerified = True 'Denote that the program running has been verified
    
ActivateForm:
    
    Call ShpBttnOK_Click
    GoTo Exit_txtDeviceSerial_GotFocus
    
FakeSerialCode:
    
    'If the Software has been licensed with a fake serial code then...
    vMsgBox "This Software has been licensed with a Fake License Code. The Licence will be revoked. Please contact the Software Administrator.", vbCritical, App.Title & " : Fake Licence", Me
    
    'Revoke all Copyright protection settings of the device
    VBA.DeleteSetting App.Title, "Copyright Protection"
    
    VBA.Randomize 'Initialize the random-number generator.
    
    'Regenerate new Serial Key for the device
    VBA.SaveSetting App.Title, "Copyright Protection", "Device Serial", SmartEncrypt(VBA.StrReverse(VBA.Int((1000 - 9999 + 1000) * VBA.Rnd + 9999)) & "-" & VBA.StrReverse(Device_Serial_No))
    
Exit_txtDeviceSerial_GotFocus:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_txtDeviceSerial_GotFocus_Error:
    
    If Err.Number = &H9 Or Err.Number = 402 Or Err.Number = &H5 Then Resume Next
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Activating Form - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_txtDeviceSerial_GotFocus
    
End Sub

Private Sub txtSerialCode_Change(Index As Integer)
On Local Error GoTo Handle_txtSerialCode_Change_Error
    
    Dim nIndex&
    
    'If the contents of the current  texbox have been deleted then...
    If txtSerialCode(Index).Text = VBA.vbNullString Then
        
        'Delete the contents of all successive textboxes
        For nIndex = Index + &H1 To txtSerialCode.UBound Step &H2
            txtSerialCode(nIndex).Text = VBA.vbNullString
        Next nIndex
        
    End If 'Close respective IF..THEN block statement
    
    If Index < txtSerialCode.UBound And VBA.Len(txtSerialCode(Index).Text) = &H5 Then txtSerialCode(Index + &H1).SetFocus
    
    If VBA.Len(VBA.Trim$(nBuffer)) = &H0 Or Index <> &H0 Then nBuffer = VBA.vbNullString: Exit Sub
    
    Dim nArray() As String
    
    nArray = VBA.Split(nBuffer, "-")
    nBuffer = VBA.vbNullString
    
    'Assign the keygen to respective controls
    For nIndex = &H0 To &H4 Step &H1
        txtSerialCode(nIndex).Text = nArray(nIndex)
    Next nIndex
    txtSerialCode(&H4).SetFocus: txtSerialCode(&H4).SelStart = VBA.Len(txtSerialCode(&H4).Text)
    
Exit_txtSerialCode_Change:
    
    'Set Mouse pointer to indicate end of this process or operation
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_txtSerialCode_Change_Error:
    
    If Err.Number = &H5 Then Resume Next
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Activating Form - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_txtSerialCode_Change
    
End Sub

Private Sub txtSerialCode_GotFocus(Index As Integer)
    txtSerialCode(Index).SelStart = &H0: txtSerialCode(Index).SelLength = VBA.Len(txtSerialCode(Index).Text)
End Sub

Private Sub txtSerialCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyRight And Index < &H4 And (txtSerialCode(Index).SelStart = VBA.Len(txtSerialCode(Index).Text) Or txtSerialCode(Index).SelLength = VBA.LenB(txtSerialCode(Index).Text)) Then txtSerialCode(Index + &H1).SetFocus: Exit Sub
    If KeyCode = vbKeyLeft And Index > &H0 And txtSerialCode(Index).SelStart = &H0 Then txtSerialCode(Index - &H1).SetFocus: Exit Sub
    
End Sub

Private Sub txtSerialCode_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If (KeyAscii >= 97 And KeyAscii <= 122) Then KeyAscii = VBA.Asc(VBA.UCase$(VBA.Chr$(KeyAscii)))
    
    nBuffer = VBA.vbNullString
    
    'If pasting Keygen from Clipboard then...
    If (KeyAscii = &H3 Or KeyAscii = 22) And Index = &H0 Then
        
        nBuffer = VB.Clipboard.GetText 'Get the Keygen from Clipboard
        
        'Validate to check if the retrieved Keygen is valid
        If VBA.Len(VBA.Replace(nBuffer, "-", VBA.vbNullString)) <> 25 Then KeyAscii = Empty: nBuffer = VBA.vbNullString: Exit Sub
        
    Else
        
        If VBA.Len(txtSerialCode(Index)) > &H4 And KeyAscii <> vbKeyBack Then
            
            If Index <> &H4 Then
                
                If txtSerialCode(Index + &H1).Text = VBA.vbNullString Then txtSerialCode(Index + &H1).Text = VBA.Chr$(KeyAscii)
                txtSerialCode(Index + &H1).SetFocus: txtSerialCode(Index + &H1).SelStart = VBA.LenB(txtSerialCode(Index + &H1).Text)
                
            End If 'Close respective IF..THEN block statement
            
            KeyAscii = Empty
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
End Sub
