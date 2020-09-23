VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_PickDetails 
   BackColor       =   &H00CFE1E2&
   Caption         =   "App Title"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   Icon            =   "Frm_PickDetails.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Stadmis.AutoSizer AutoSizer 
      Left            =   120
      Top             =   120
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Frame Fra_Photo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Image ImgDBPhoto 
         Height          =   1215
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Image ImgVirtualPhoto 
         Height          =   135
         Left            =   240
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape shpFraPhotoBorder 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   1455
         Left            =   0
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CheckBox ChkDisplayPhoto 
      BackColor       =   &H00CFE1E2&
      Caption         =   "&Display Photo"
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
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Tag             =   "AutoSizer:XY"
      Top             =   4440
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComctlLib.ListView Lv 
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Tag             =   "AutoSizer:WH"
      Top             =   600
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "ilRecordIco"
      SmallIcons      =   "ImgLst"
      ForeColor       =   128
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   -720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":1164
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":1B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":2A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":31B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":354C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":443E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":47D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":51D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":5624
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":5A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm_PickDetails.frx":5EC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   5
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
      Caption         =   "      Cancel"
      AccessKey       =   "C"
      Picture         =   "Frm_PickDetails.frx":631A
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
      Left            =   2880
      TabIndex        =   6
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
      Caption         =   "    OK"
      AccessKey       =   "O"
      Picture         =   "Frm_PickDetails.frx":66B4
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Tag             =   "AutoSizer:Y"
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
      Caption         =   "      Search"
      AccessKey       =   "S"
      Picture         =   "Frm_PickDetails.frx":6C4E
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
   Begin Stadmis.ShapeButton ShpBttnCheck 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   120
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
      Caption         =   "Check All"
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
   Begin Stadmis.ShapeButton ShpBttnCheck 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   9
      Top             =   120
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
      Caption         =   "UnCheck All"
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
   Begin VB.Label lblTotalRecords 
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
      TabIndex        =   3
      Tag             =   "AutoSizer:Y"
      Top             =   4440
      Width           =   1395
   End
   Begin VB.Label lblSelectedCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   -120
      Picture         =   "Frm_PickDetails.frx":6FE8
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Image ImgHeader 
      Height          =   600
      Left            =   0
      Picture         =   "Frm_PickDetails.frx":77DE
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Frm_PickDetails"
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

Public iSelColPos&
Public iTargetFields$, iLvwItemDataType$, iPhotoSpecifications$, iSearchIDs$

Private Xaxis&
Private DragPhoto, iSelectionComplete As Boolean

Private Function ApplyTheme() As Boolean
    
    '**************************************************************
    'Apply Theme Settings
    
    Me.BackColor = tTheme.tBackColor
    Fra_Photo.BackColor = Me.BackColor
    
    ImgHeader.Picture = tTheme.tImagePicture
    ImgFooter.Picture = ImgHeader.Picture
    
    lblSelectedCategory.ForeColor = tTheme.tForeColor
    
    ShpBttnCheck(&H0).ForeColor = tTheme.tButtonForeColor
    ShpBttnCheck(&H0).BackColor = tTheme.tButtonBackColor
    
    ShpBttnCheck(&H1).ForeColor = tTheme.tButtonForeColor
    ShpBttnCheck(&H1).BackColor = tTheme.tButtonBackColor
    
    lblTotalRecords.ForeColor = tTheme.tForeColor
    
    ChkDisplayPhoto.ForeColor = tTheme.tForeColor
    ChkDisplayPhoto.BackColor = tTheme.tBackColor
    
    ShpBttnSearch.ForeColor = tTheme.tButtonForeColor
    ShpBttnSearch.BackColor = tTheme.tButtonBackColor
    
    ShpBttnCancel.ForeColor = tTheme.tButtonForeColor
    ShpBttnCancel.BackColor = tTheme.tButtonBackColor
    
    ShpBttnOK.ForeColor = tTheme.tButtonForeColor
    ShpBttnOK.BackColor = tTheme.tButtonBackColor
    
    '**************************************************************
    
End Function

Private Sub ChkDisplayPhoto_Click()
    If Fra_Photo.Visible And ChkDisplayPhoto.Value = vbUnchecked Then Fra_Photo.Visible = False
End Sub

Private Sub ShpBttnCheck_Click(Index As Integer)
    
    If Not Lv.Checkboxes Then Exit Sub
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    For vIndex(&H0) = &H1 To Lv.ListItems.Count Step &H1
        Lv.ListItems(vIndex(&H0)).Checked = Not VBA.CBool(Index)
    Next vIndex(&H0)
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Activate()
On Local Error Resume Next
    vMultiSelectedData = VBA.vbNullString: Lv.Visible = True
    iSelColPos = VBA.IIf(iSelColPos = &H0, &H2, iSelColPos): Lv.SetFocus
End Sub

Private Sub Form_Load()
    
    Me.Caption = App.Title
    Call AutoSizer.GetInitialPositions
    If vFrmWidth <> &H0 Then Me.Width = vFrmWidth
    vFrmWidth = &H0
    Call AutoSizer.AutoResize
    'Call ApplyTheme
    
End Sub

Private Sub Fra_Photo_Click()
    Call ImgDBPhoto_Click
End Sub

Private Sub ImgDBPhoto_Click()
    Call PhotoClicked(ImgDBPhoto, ImgVirtualPhoto, Fra_Photo, True)
End Sub

Private Sub Lv_Click()
    Lv.MultiSelect = False 'Remove Multi-selection
End Sub

Private Sub Lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Static PrevCol&
    
    'If the same column is clicked twice then toggle Sort order else sort ascending
    SortListview Lv, ColumnHeader.Position, VBA.IIf(PrevCol = ColumnHeader.Position, VBA.IIf(Lv.SortOrder = &H0, &H1, &H0), &H0)
    PrevCol = ColumnHeader.Position
    
End Sub

Private Sub Lv_DblClick()
    
    'If their are no Records to select then quit this Procedure
    If Lv.ListItems.Count = &H0 Then Exit Sub
    
    If Lv.Checkboxes Then Lv.SelectedItem.Checked = True
    
    'Call Click event of the specified control
    Call ShpBttnOK_Click
    
End Sub

Private Sub Lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Local Error GoTo Handle_Lv_ItemClick_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'If the Photo Table has not been specified or the Photo shouldn't be displayed then Quit this Procedure
    If iPhotoSpecifications = VBA.vbNullString Or ChkDisplayPhoto.Value = vbUnchecked Then GoTo Exit_Lv_ItemClick
    
    Dim iPhotoFld$
    Dim iArray() As String
    Dim nRs As New ADODB.Recordset
        
    iArray = VBA.Split(VBA.Replace(iPhotoSpecifications, ";", "|"), "|")
    
    iPhotoFld = "Photo" 'Assign default name of Photo field
    
    ConnectDB , False 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set nRs = New ADODB.Recordset
    With nRs 'Execute a series of statements on nRs recordset
        
        .Open "SELECT * FROM " & iArray(&H0) & VBA.Replace(iArray(&H1), "$", Item.Text), vAdoCNN, adOpenKeyset, adLockReadOnly
        
DisplayPhoto:
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then
            
            Dim iPhotoName$
            
            If UBound(iArray) >= &H2 Then
                
                If Lv.ColumnHeaders.Count > iArray(&H2) Then
                    
                    iPhotoName = VBA.Trim$(Lv.ColumnHeaders(VBA.Val(iArray(&H2))).Text)
                    
                    'Display field values in their respective controls
                    If Not VBA.IsNull(nRs(iPhotoName)) Then iPhotoName = nRs(iPhotoName)
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            'If the Record contains the Person's Photo then...
            If Not VBA.IsNull(nRs(iPhotoFld)) Then
                
                'Display Person's Photo
                Set ImgVirtualPhoto.DataSource = Nothing
                ImgVirtualPhoto.DataField = iPhotoFld
                Set ImgVirtualPhoto.DataSource = nRs
                ImgVirtualPhoto.Refresh
                ImgVirtualPhoto.ToolTipText = VBA.IIf(iPhotoName = VBA.vbNullString, Item.ListSubItems(&H1).Text, iPhotoName) & "'s " & iPhotoFld
                
                'Call Procedure to Fit image to image holder
                Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_Photo)
                
                'Position the Photo Frame directly below the selected item
                Fra_Photo.Top = Lv.Top + Item.Top + Item.Height + 50
                
                Fra_Photo.Visible = True 'Display the Photo Frame
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
Exit_Lv_ItemClick:
    
    'Re-initialize elements of the fixed-size array and release dynamic-array storage space.
    Erase vIndex
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Procedure
    
Handle_Lv_ItemClick_Error:
    
    If Err.Number = 3265 Then If iPhotoFld <> "Logo" Then iPhotoFld = "Logo": Resume DisplayPhoto Else Resume Exit_Lv_ItemClick
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Clicking Item - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_Lv_ItemClick
    
End Sub

Private Sub Lv_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not Nothing Is Lv.SelectedItem Then
        
        Select Case KeyCode
            
            Case vbKeyUp: 'If the Up key is pressed an its on the first item them jump to the last item in the list
                
                If Lv.SelectedItem.Index = &H1 Then Lv.ListItems(Lv.ListItems.Count).Selected = True: Lv.ListItems(Lv.ListItems.Count).EnsureVisible: KeyCode = Empty
                
            Case vbKeyDown: 'If the Down key is pressed an its on the last item them jump to the first item in the list
                
                If Lv.SelectedItem.Index = Lv.ListItems.Count Then Lv.ListItems(&H1).Selected = True: Lv.ListItems(&H1).EnsureVisible: KeyCode = Empty
                
        End Select 'Close SELECT..CASE block statement
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub Lv_KeyPress(KeyAscii As Integer)
    
    Dim iLst As ListItem
    
    Set iLst = FindLvItemByFirstItemChar(Lv, KeyAscii, iSelColPos)
    If Not Nothing Is iLst Then Call Lv_ItemClick(iLst)
    
End Sub

Private Sub Lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim ColDistance&
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'For each column in the Listview...
    For vIndex(&H0) = &H1 To Lv.ColumnHeaders.Count Step &H1
        
        'Get the column distance from the left side of the Listview
        ColDistance = ColDistance + Lv.ColumnHeaders(vIndex(&H0)).Width
        
        'If the distance exceeds the clicked position then pick the current column position
        If ColDistance > x Then iSelColPos = vIndex(&H0): Exit For
        
    Next vIndex(&H0) 'Move to the next column
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ShpBttnCancel_Click()
    Unload Me 'Unload this form from memory.
End Sub

Private Sub ShpBttnOK_Click()
    
    Dim Lst
    Dim vRow&, vCol&
    Dim iSelRowCnt&, iSelColCnt&
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    iSelRowCnt = &H0 'Initialize Row Counter
    vMultiSelectedData = VBA.vbNullString
    
    Dim nValidRow As Boolean
    Dim nMultiSelectedData$
    
    'For each list item in the Listview control...
    For vRow = &H1 To Lv.ListItems.Count Step &H1
        
        If Lv.Checkboxes Then nValidRow = Lv.ListItems(vRow).Checked Else nValidRow = Lv.ListItems(vRow).Selected
        
        'If the List item has been selected or checked then...
        If nValidRow Then
            
            'For each column in the Listview control...
            For vCol = &H1 To Lv.ColumnHeaders.Count Step &H1
                
                'If the Field is not hidden then...
                If VBA.InStr("|" & VBA.Replace(iTargetFields, ";", "|") & "|", "|" & vCol & "|") <> &H0 Then
                    
                    If vCol = &H1 Then Set Lst = Lv.ListItems(vRow) Else Set Lst = Lv.ListItems(vRow).ListSubItems(vCol - &H1)
                    
                    'Separate Columns by "~"
                    nMultiSelectedData = nMultiSelectedData & "~" & Lst.Text 'Assign column value
                    
                End If 'Close respective IF..THEN block statement
                
            Next vCol 'Move to the next column
            
            'If the data starts with '~' character then remove it
            If VBA.Left$(nMultiSelectedData, &H1) = "~" Then nMultiSelectedData = VBA.Right$(nMultiSelectedData, VBA.Len(nMultiSelectedData) - &H1)
            
            'Separate Rows by "|"
            vMultiSelectedData = vMultiSelectedData & nMultiSelectedData & "|"
            nMultiSelectedData = VBA.vbNullString
            
        End If 'Close respective IF..THEN block statement
        
    Next vRow 'Move to the next row
    
    'If the data ends with '|' character then remove it
    If VBA.Right$(vMultiSelectedData, &H1) = "|" Then vMultiSelectedData = VBA.Left$(vMultiSelectedData, VBA.Len(vMultiSelectedData) - &H1)
    
    'If at least 1 item has been selected then denote that the selection is true
    If iSelColCnt <> &H0 Then iSelectionComplete = True
    
    Unload Me 'Unload this form from memory.
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub ShpBttnSearch_Click()
    
    Fra_Photo.Visible = False 'Hide the Photo Frame
    
    'Call Function in Mdl_Stadmis to perform Search
    Call SearchLv(Lv, , iSelColPos)
    
End Sub


'------------------------------------------------------------------------------------------
'The Following Codes enable the User to Drag the displayed Photo across the Listview Item

Private Sub DragMe(CurrX!)
    
    If DragPhoto = True Then 'Check if it is in drag mode
    
        'If the Photo is dragged past the Listview left position
        If Fra_Photo.Left < Lv.Left + VBA.IIf(Not Nothing Is Lv.SmallIcons, 300, &H0) Then
        
            'Set the Photo to the left position of the sheet
            Fra_Photo.Left = Lv.Left + VBA.IIf(Not Nothing Is Lv.SmallIcons, 300, &H0)
            DragPhoto = False 'Disable drag mode
            Exit Sub 'Quit this procedure
            
        Else 'If the Photo is dragged within the Listview then...
        
            'If the Photo is dragged past the Listview right position
            If Fra_Photo.Left > (Lv.Width + Lv.Left) - Fra_Photo.Width Then
                
                'Set the Photo to the right position of the Listview
                Fra_Photo.Left = (Lv.Width + Lv.Left) - Fra_Photo.Width
                DragPhoto = False 'Disable drag mode
                Exit Sub 'Quit this procedure
                
            Else
                
                'If the Photo frame is dragged within the Listview
                'Move the Photo to the mouse pointer position
                Fra_Photo.Move (Fra_Photo.Left + CurrX - Xaxis)
                
            End If 'Close respective IF..THEN block statement
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
End Sub

Private Sub Fra_Photo_MouseDown(Button%, Shift%, x!, y!)
    
    DragPhoto = True 'Enable dragging of the Photo
    
    'Set the mouse pointer to cross to show drag mode
    Screen.MousePointer = vbSizeAll
    
    Xaxis = x 'Capture the current X-axis position
    
End Sub

Private Sub Fra_Photo_MouseMove(Button%, Shift%, x!, y!)
    DragMe x 'Call procedure
End Sub

Private Sub Fra_Photo_MouseUp(Button%, Shift%, x!, y!)
    DragPhoto = False 'Disable drag mode
    Screen.MousePointer = vbDefault 'Set mouse pointer to default pointer to portray end of processing
End Sub

Private Sub ImgDBPhoto_MouseDown(Button%, Shift%, x!, y!)
    DragPhoto = True 'Enable dragging of the Photo
    'Set the mouse pointer to cross to show drag mode
    Screen.MousePointer = vbSizeAll
    Xaxis = x 'Capture the current X-axis position
End Sub

Private Sub ImgDBPhoto_MouseMove(Button%, Shift%, x!, y!)
    DragMe x 'Call procedure
End Sub

Private Sub ImgDBPhoto_MouseUp(Button%, Shift%, x!, y!)
    DragPhoto = False 'Disable drag mode
    Screen.MousePointer = vbDefault 'Set mouse pointer to default pointer to portray end of processing
End Sub

