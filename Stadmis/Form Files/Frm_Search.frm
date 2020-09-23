VERSION 5.00
Begin VB.Form Frm_Search 
   BackColor       =   &H00CFE1E2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "App Title : Search"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "Frm_Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkOption 
      BackColor       =   &H00CFE1E2&
      Caption         =   "Search Hidden Cols"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Search items that do not confirm with the search criteria"
      Top             =   2980
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox ChkOption 
      BackColor       =   &H00CFE1E2&
      Caption         =   "Exclude Criteria"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Search items that do not confirm with the search criteria"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdEscapeClose 
      Height          =   255
      Left            =   -360
      TabIndex        =   16
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton cmdEnterSearch 
      Height          =   255
      Left            =   -360
      TabIndex        =   15
      Top             =   3000
      Width           =   255
   End
   Begin VB.Frame Fra_Details 
      BackColor       =   &H00CFE1E2&
      Caption         =   "Search Condition:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   3975
      Begin VB.CheckBox ChkOption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Filter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Exclude items that do not match the entered search criteria"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox ChkOption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "Match Case"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Case-Sensitivity"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox ChkOption 
         BackColor       =   &H00CFE1E2&
         Caption         =   "&Exact Match"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Search for items with the whole of the entered search criteria"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton OptSearch 
         BackColor       =   &H80000013&
         Caption         =   "&Exact"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   2310
         Width           =   855
      End
      Begin VB.OptionButton OptSearch 
         BackColor       =   &H80000013&
         Caption         =   "&Any"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   2310
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox CboSearch 
         BackColor       =   &H00CFE1E2&
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "Frm_Search.frx":076A
         Left            =   240
         List            =   "Frm_Search.frx":076C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "0"
         ToolTipText     =   "Select the search column."
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox TxtSearch 
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
         Left            =   240
         TabIndex        =   1
         Text            =   "Type and press Enter key"
         ToolTipText     =   "Type Search string and press Enter key"
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblLookIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Look In:"
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
         TabIndex        =   2
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search by:"
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
         Top             =   240
         Width           =   885
      End
   End
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2760
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
      Picture         =   "Frm_Search.frx":076E
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
   Begin Stadmis.ShapeButton ShpBttnSearch 
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2760
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
      Caption         =   "      Search"
      AccessKey       =   "S"
      Picture         =   "Frm_Search.frx":0B08
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
   Begin VB.Image Image3 
      Height          =   495
      Left            =   -120
      Picture         =   "Frm_Search.frx":0EA2
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label LblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_Search.frx":1698
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   600
      TabIndex        =   14
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image ImgSearch 
      Height          =   375
      Left            =   120
      Picture         =   "Frm_Search.frx":1723
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "Frm_Search.frx":1E8D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Frm_Search"
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

Public TargetLv As ListView

Private Sub Form_Load()
    Me.Caption = App.Title & " : Search"
    If CboSearch.List(&H0) <> "[All]" Then CboSearch.AddItem "[All]", &H0
End Sub

Private Sub ShpBttnCancel_Click()
    Unload Me
End Sub

Private Sub ShpBttnSearch_Click()
On Local Error GoTo Handle_ShpBttnSearch_Click_Error
    
    'If the User has not typed anything then...
    If TxtSearch.Text = "Type and press Enter key" Or VBA.Trim$(TxtSearch.Text) = VBA.vbNullString Then
        
        'Warn the User
        vMsgBox "Please enter a string to search", vbExclamation, App.Title & " : No Search String", Me
        Exit Sub 'Quit this Procedure
        
    End If 'Close respective IF..THEN block statement
    
    Dim Lst
    Dim nArray() As String
    Dim sArray() As String
    Dim sLookFor$, sRowData$, sStr$
    Dim sRow&, sCol&, sStartCol&, sEndCol&, sMatchCnt&, sFirstMatch&, sNum&
    Dim sExactMatch, sMatchCase, sFilter, sMatchFound, sExclude, sVisibleDataOnly As Boolean
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    TargetLv.Visible = False 'Hide Listview for faster items addition
    
    'Get set values
    sExactMatch = (ChkOption(&H0).Value = vbChecked)
    sMatchCase = (ChkOption(&H1).Value = vbChecked)
    sFilter = (ChkOption(&H2).Value = vbChecked)
    sExclude = (ChkOption(&H3).Value = vbChecked)
    sLookFor = VBA.Replace(VBA.IIf(sMatchCase, TxtSearch.Text, VBA.LCase$(TxtSearch.Text)), "  ", " ")
    sVisibleDataOnly = (ChkOption(&H4).Value = vbUnchecked)
    
    sStartCol = VBA.IIf(CboSearch.ListIndex = &H0, &H1, CboSearch.ListIndex)
    sEndCol = VBA.IIf(CboSearch.ListIndex = &H0, CboSearch.ListCount - &H1, sStartCol)
    
    'Execute a series of statements on the specified Listview
    With TargetLv
        
        nArray = VBA.Split(.Parent.iSearchIDs, "|")
        
        'Set defaults
        .SelectedItem = Nothing: .MultiSelect = True: .Parent.iSearchIDs = VBA.vbNullString
        
        'For each list item in the specified Listview control...
        For sRow = &H1 To .ListItems.Count Step &H1
            
            'If filtered to the last item then Quit this FOR..LOOP block statement
            If sRow > .ListItems.Count Then Exit For
            
            If .ListItems(sRow).Tag = "Function" Then If sFilter Then .ListItems.Remove sRow: sRow = sRow - &H1: GoTo NextRow
            
            sRowData = "|" 'Initialize variable
            
            'For each column in the specified Listview control...
            For sCol = sStartCol To sEndCol Step &H1
                
                'If search in columns that have visible data only then...
                If sVisibleDataOnly Then
                    
                    'Search in columns that have visible data only
                    If .ColumnHeaders(sCol).Width > 80 + VBA.IIf(Not Nothing Is .SmallIcons, 250, &H0) + VBA.IIf(.Checkboxes, 250, &H0) Then
                        
                        If .ListItems(sRow).ListSubItems.Count = (sCol - &H1) Then
                            If sCol = &H1 Then Set Lst = .ListItems(sRow) Else Set Lst = .ListItems(sRow).ListSubItems(sCol - &H1)
                            sRowData = sRowData & Lst.Text & "|" 'Assign the Row data into one string
                        Else
                            sRowData = sRowData & "|"  'Assign the Row data into one string
                        End If 'Close respective IF..THEN block statement
                        
                    End If 'Close respective IF..THEN block statement
                    
                Else 'If not search in columns that have visible data only then...
                    
                    If .ListItems(sRow).ListSubItems.Count >= (sCol - &H1) Then
                        If sCol = &H1 Then Set Lst = .ListItems(sRow) Else Set Lst = .ListItems(sRow).ListSubItems(sCol - &H1)
                        sRowData = sRowData & Lst.Text & "|" 'Assign the Row data into one string
                    Else
                        sRowData = sRowData & "|"  'Assign the Row data into one string
                    End If 'Close respective IF..THEN block statement
                                        
                End If 'Close respective IF..THEN block statement
                
            Next sCol 'Move to the next Column Header
            
            sMatchFound = False 'Denote, by default, that the row does not match the search criteria
            
            'If case-sensitive then...
            If sMatchCase Then
                
                'If search exactly as entered then...
                If sExactMatch Then
                    sMatchFound = (VBA.InStr(sRowData, "|" & sLookFor & "|") <> &H0)
                Else 'If search any occurrence of the entered value then...
                    
                    sArray = VBA.Split(sLookFor, " ")
                    
                    sMatchFound = True 'Denote that the row matches the search criteria
                    
                    'Reverse the search string too
                    For sCol = &H0 To UBound(sArray) Step &H1
                        If Not sMatchFound Then Exit For
                        sMatchFound = sRowData Like "*" & sArray(sCol) & "*"
                    Next sCol 'Move to the next sArray element
                    
                End If 'Close respective IF..THEN block statement
                
            Else 'If not case-sensitive then...
                
                'If search exactly as entered then...
                If sExactMatch Then
                    sMatchFound = (VBA.InStr(VBA.LCase$(sRowData), "|" & sLookFor & "|") <> &H0)
                Else 'If search any occurrence of the entered value then...
                    
                    sArray = VBA.Split(sLookFor, " ")
                    
                    sMatchFound = True 'Denote that the row matches the search criteria
                    
                    'Reverse the search string too
                    For sCol = &H0 To UBound(sArray) Step &H1
                        If Not sMatchFound Then Exit For
                        sMatchFound = VBA.LCase$(sRowData) Like "*" & VBA.LCase$(sArray(sCol)) & "*"
                    Next sCol 'Move to the next sArray element
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            'If the records with the entered Search criteria are to be excluded then
            If sExclude Then sMatchFound = Not sMatchFound
            
            'If the row data does not confirm with the specified search criteria then...
            If Not sMatchFound Then
                
                'Deselect the item if selected
                .ListItems(sRow).Selected = False
                
                'If the items are to be selectively screened out in the Listview then...
                If sFilter Then
                    
                    'Enable the TOTAL row to be visible and with the correct summations after filter
                    If .ListItems(.ListItems.Count).Tag = "Function" Then
                        
                        'For each column
                        For sNum = &H1 To .ColumnHeaders.Count Step &H1
                            
                            'If it is the first column..
                            If sNum = &H1 Then
                                
                                'If the column sum is not blank then deduct the column amount of the row to be filtered out
                                If .ListItems(.ListItems.Count).ListSubItems.Count >= (sNum - &H1) Then If .ListItems(.ListItems.Count).Text <> VBA.vbNullString Then .ListItems(.ListItems.Count).Text = VBA.FormatNumber(VBA.Val(VBA.Replace(.ListItems(.ListItems.Count).Text, ",", "")) - VBA.Val(VBA.Replace(.ListItems(sRow).Text, ",", "")), &H2)
                                
                            Else 'If it is not the first column..
                                
                                'If the column sum is not blank then deduct the column amount of the row to be filtered out
                                If .ListItems(.ListItems.Count).ListSubItems.Count >= (sNum - &H1) Then If .ListItems(.ListItems.Count).ListSubItems(sNum - &H1).Text <> VBA.vbNullString Then .ListItems(.ListItems.Count).ListSubItems(sNum - &H1).Text = VBA.FormatNumber(VBA.Val(VBA.Replace(.ListItems(.ListItems.Count).ListSubItems(sNum - &H1).Text, ",", "")) - VBA.Val(VBA.Replace(.ListItems(sRow).ListSubItems(sNum - &H1).Text, ",", "")), &H2)
                                
                            End If 'Close respective IF..THEN block statement
                            
                        Next sNum 'Increment counter variable by value in the Step option of the FOR..LOOP
                        
                    End If 'Close respective IF..THEN block statement
                    
                    'Remove the row that is to be filtered out
                    .ListItems.Remove sRow: sRow = sRow - &H1
                    
                    If UBound(nArray) >= &H0 Then nArray(sRow) = VBA.vbNullString: sStr = VBA.Join(nArray, "|")
                    
                    'If the data starts with '|' character then remove it
                    If VBA.Left$(sStr, &H1) = "|" Then sStr = VBA.Right$(sStr, VBA.Len(sStr) - &H1)
                    
                    'If the data ends with '|' character then remove it
                    If VBA.Right$(sStr, &H1) = "|" Then sStr = VBA.Left$(sStr, VBA.Len(sStr) - &H1)
                    
                    nArray = VBA.Split(sStr, "|")
                    
                End If 'Close respective IF..THEN block statement
                
            Else 'If the row data confirms with the specified search criteria then...
                
                sMatchCnt = sMatchCnt + &H1 'Increment the variable that denotes no of matching rows
                sFirstMatch = VBA.IIf(sMatchCnt = &H1, sRow, sFirstMatch) 'Get the row index of the first matching row
                .ListItems(sRow).Selected = True 'Select the item
                .Parent.iSearchIDs = .Parent.iSearchIDs & "|" & .ListItems(sRow).ListSubItems(&H1).Text
                
            End If 'Close respective IF..THEN block statement
            
NextRow:
            'If the next row is the last and is the Total Row then leave it in the Listview
            If .ListItems.Count > 0 Then If sRow + &H1 = .ListItems.Count And .ListItems(.ListItems.Count).Tag = "Function" Then sRow = sRow + &H1
            
        Next sRow 'Move to the next List Item
        
        If sMatchCnt > &H0 Then .ListItems(.SelectedItem.Index).EnsureVisible
        
    End With 'Close the WITH block statements
    
    Unload Me 'Unload this form from the memory.
    
    'If Match found then display to the user the first Record that matched the search criteria
    If sFirstMatch > &H0 And TargetLv.ListItems.Count >= sFirstMatch + &H1 Then TargetLv.ListItems(sFirstMatch).Selected = True: TargetLv.ListItems(sFirstMatch).EnsureVisible
    
    TargetLv.Visible = True 'Display Listview
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox VBA.IIf(sMatchCnt = &H0, "No", sMatchCnt) & " Records found matching the specified criteria", vbInformation, App.Title & " : Search Results", TargetLv.Parent
    
    TargetLv.Visible = True 'Display Listview
    TargetLv.SetFocus 'Move focus to the Listview control
    
Exit_ShpBttnSearch_Click:
    
    TargetLv.Parent.iSearchIDs = sStr
    
    'If the data starts with '|' character then remove it
    If VBA.Left$(TargetLv.Parent.iSearchIDs, &H1) = "|" Then TargetLv.Parent.iSearchIDs = VBA.Right$(TargetLv.Parent.iSearchIDs, VBA.Len(TargetLv.Parent.iSearchIDs) - &H1)
    
    TargetLv.Visible = True 'Display Listview
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Procedure
    
Handle_ShpBttnSearch_Click_Error:
    
    If Err.Number = &H5 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Searching - " & Err.Number, Me
    
    TargetLv.Visible = True 'Display Listview
    TxtSearch.SetFocus 'Move focus to the textbox control
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnSearch_Click
    
End Sub

Private Sub TxtSearch_Change()
    'If the search input box is blank then assign default
    If TxtSearch.Text = VBA.vbNullString Then TxtSearch.Text = "Type and press Enter key": TxtSearch.SelStart = &H0: TxtSearch.SelLength = VBA.Len(TxtSearch.Text)
End Sub

Private Sub TxtSearch_GotFocus()
    'If the search input box gets focus then highlight all its contents
    TxtSearch.SelStart = &H0: TxtSearch.SelLength = VBA.Len(TxtSearch.Text)
End Sub
