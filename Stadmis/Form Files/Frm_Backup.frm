VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frm_Backup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "App Title : Set up Backup"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   Icon            =   "Frm_Backup.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9855
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.AutoSizer AutoSizer 
      Left            =   720
      Top             =   5400
      _ExtentX        =   661
      _ExtentY        =   661
      MinWidth        =   9945
      MinHeight       =   6390
   End
   Begin VB.PictureBox picBuffer 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   -720
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   1200
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame Fra_Backup 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   2640
      TabIndex        =   3
      Tag             =   "AutoSizer:WH"
      Top             =   720
      Width           =   7095
      Begin MSComctlLib.ListView Lv 
         Height          =   3000
         Left            =   120
         TabIndex        =   0
         Tag             =   "AutoSizer:WH"
         Top             =   960
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         Icons           =   "ImgLst"
         SmallIcons      =   "ImgLst"
         ForeColor       =   16711680
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
      Begin VB.Frame Fra_MenuOptions 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Tag             =   "AutoSizer:W"
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox chkDisplayAllFolders 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Display all Folders"
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
            Height          =   255
            Left            =   2880
            TabIndex        =   17
            Tag             =   "AutoSizer:XY"
            Top             =   150
            Width           =   1575
         End
         Begin VB.CheckBox ChkDisplayFiles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Display Files"
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
            Height          =   255
            Left            =   1680
            TabIndex        =   16
            Tag             =   "AutoSizer:XY"
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton cmdMenuOptions 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Index           =   3
            Left            =   5400
            Picture         =   "Frm_Backup.frx":09EA
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "AutoSizer:X"
            ToolTipText     =   "Up One Level"
            Top             =   120
            Width           =   345
         End
         Begin VB.CommandButton cmdMenuOptions 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Index           =   2
            Left            =   5760
            Picture         =   "Frm_Backup.frx":0F74
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "AutoSizer:X"
            ToolTipText     =   "Delete Folder"
            Top             =   120
            Width           =   345
         End
         Begin VB.CommandButton cmdMenuOptions 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Index           =   1
            Left            =   6120
            Picture         =   "Frm_Backup.frx":12FE
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "AutoSizer:X"
            ToolTipText     =   "Create New Folder"
            Top             =   120
            Width           =   345
         End
         Begin VB.CommandButton cmdMenuOptions 
            BackColor       =   &H00D8E9EC&
            Height          =   315
            Index           =   0
            Left            =   6480
            Picture         =   "Frm_Backup.frx":1688
            Style           =   1  'Graphical
            TabIndex        =   12
            Tag             =   "AutoSizer:X"
            ToolTipText     =   "View"
            Top             =   120
            Width           =   345
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save Backup on:"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   165
            Width           =   1335
         End
      End
      Begin VB.Label lblSelectedPath 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "My Computer"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Tag             =   "AutoSizer:YW"
         ToolTipText     =   "Click here if you have forgotten your account password"
         Top             =   3960
         Width           =   6840
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "We recommend that you save your Backups on an external hard drive."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   600
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Tag             =   "AutoSizer:W"
         Top             =   120
         Width           =   6855
         WordWrap        =   -1  'True
      End
   End
   Begin Stadmis.ShapeButton ShpBttnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Tag             =   "AutoSizer:XY"
      Top             =   5400
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
      Picture         =   "Frm_Backup.frx":1A12
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
   Begin Stadmis.ShapeButton ShpBttnBackup 
      Default         =   -1  'True
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Tag             =   "AutoSizer:XY"
      Top             =   5400
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
      Caption         =   "      Backup"
      AccessKey       =   "B"
      Picture         =   "Frm_Backup.frx":1DAC
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
   Begin Stadmis.ShapeButton ShpBttnRefresh 
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Tag             =   "AutoSizer:Y"
      ToolTipText     =   "Refresh"
      Top             =   5400
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
      Caption         =   "      Refresh"
      AccessKey       =   "R"
      Picture         =   "Frm_Backup.frx":2146
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
   Begin MSComctlLib.ImageList ImgLstFiles 
      Left            =   120
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lblReadme 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Why save Backup? (Click here)"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   4080
      TabIndex        =   7
      Tag             =   "AutoSizer:XY"
      ToolTipText     =   "Click here if you have forgotten your account password"
      Top             =   5475
      Width           =   2175
   End
   Begin VB.Image ImgFooter 
      Height          =   615
      Left            =   2640
      Picture         =   "Frm_Backup.frx":24E0
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   5280
      Width           =   7095
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select where you want to save your Backup"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   0
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   5325
   End
   Begin VB.Image ImgHeader 
      Height          =   615
      Left            =   2640
      Picture         =   "Frm_Backup.frx":2CD6
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   120
      Width           =   7095
   End
   Begin VB.Image ImgBackupUtility 
      Height          =   6255
      Left            =   -120
      Picture         =   "Frm_Backup.frx":34CC
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   8055
   End
   Begin VB.Menu MnuListviewViews 
      Caption         =   "Listview Views"
      Visible         =   0   'False
      Begin VB.Menu MnuListviewView 
         Caption         =   "Thumbnails"
         Index           =   0
      End
      Begin VB.Menu MnuListviewView 
         Caption         =   "Icons"
         Index           =   1
      End
      Begin VB.Menu MnuListviewView 
         Caption         =   "List"
         Index           =   2
      End
      Begin VB.Menu MnuListviewView 
         Caption         =   "Details"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Frm_Backup"
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

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal FLAGS&) As Long
Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Enum iIconSize
    
    [16_x_16] = 16
    [32_x_32] = 32
    
End Enum

Private Type typSHFILEINFO
    
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
    
End Type

Private iStatus$
Private iDriveSpace#, iLevel#
Private iArrayList() As String
Private iItemClicked As Boolean
Private FileInfo As typSHFILEINFO
Private iFolderIconIndex&, iFileIconIndex&
Private iDrives, iDrive, iFldr, iFldrs, iSubFldrs, iSubFldr, iFl, iFls

Private Function IconExtraction(iImgLst As ImageList, iPath$, IconSize As iIconSize) As Long
    
    Dim R As Integer
    
    R = ExtractIcon(iPath, iImgLst, picBuffer, VBA.CInt(IconSize))
    
    If R = &H0 Then Exit Function
    
    IconExtraction = R
    
End Function

Private Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), FLAGS Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), FLAGS Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> &H0 Then
        
        With PictureBox
            
            .Height = 15 * PixelsXY
            .Width = 15 * PixelsXY
            .ScaleHeight = 15 * PixelsXY
            .ScaleWidth = 15 * PixelsXY
            .Picture = LoadPicture("")
            .AutoRedraw = True
            
            SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hDc, 0, 0, ILD_TRANSPARENT)
            .Refresh
            
        End With
        
        IconIndex = AddtoImageList.ListImages.Count + &H1
        Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
        ExtractIcon = IconIndex
        
    End If
    
End Function

Private Function GetActualSize(iSize#) As String
On Local Error GoTo Handle_GetActualSize_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim iCnt&
    Dim iIndex#, vIndex#
    
    iIndex = iIndex + 1024
    iCnt = &H0
    vIndex = VBA.Val(iSize)
    
    Do While vIndex >= iIndex
        
        iIndex = iIndex * 1024
        iCnt = iCnt + &H1
        
    Loop
    
    GetActualSize = VBA.Replace(VBA.FormatNumber(vIndex / (1024 ^ iCnt), &H3), ".000", VBA.vbNullString)
    
    Select Case iCnt
        
        Case &H0: 'B
            GetActualSize = GetActualSize & " B"
        Case &H1: 'KB
            GetActualSize = GetActualSize & " KB"
        Case &H2: 'MB
            GetActualSize = GetActualSize & " MB"
        Case &H3: 'GB
            GetActualSize = GetActualSize & " GB"
        Case &H4: 'TB
            GetActualSize = GetActualSize & " TB"
                        
    End Select 'Close SELECT..CASE block statement
    
Exit_GetActualSize:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_GetActualSize_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Getting Actual Drive Size - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_GetActualSize
    
End Function

Private Function ListDrives() As Boolean
On Local Error GoTo Handle_ListDrives_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Initialize controls' properties
    If ImgLst.ListImages.Count > &H0 Then Lv.Icons = Nothing: Lv.SmallIcons = Nothing: ImgLst.ListImages.Clear
    
    Lv.ColumnHeaders.Clear: Lv.ListItems.Clear
    
    'Create Column Headers
    Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Backup Destination", 2150
    Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Free Space", 1100, lvwColumnRight
    Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Total Size", 1100, lvwColumnRight
    Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Drive Type", 1500, lvwColumnRight
    Lv.ColumnHeaders.Add Lv.ColumnHeaders.Count + &H1, , "Attributes", 800, lvwColumnRight
    
    VBA.DoEvents: VBA.DoEvents
    
    'Return a Drives collection consisting of all Drive objects available on the local machine
    Set iDrives = vFso.Drives
    
    Dim vNum&, iCounter&, iCnt&
    
    iCounter = iDrives.Count
    Frm_PleaseWait.ImgProgressBar.Width = &H0
    Frm_PleaseWait.ImgProgressBar.Left = Frm_PleaseWait.lblProgressBar.Left
    Frm_PleaseWait.ImgProgressBar.Visible = True
    
    'Show the 'Please Wait' Form to the User
    CenterForm Frm_PleaseWait, , False
    
'    Lv.Visible = False
    
    For Each iDrive In iDrives
        
        Frm_PleaseWait.lblInfo.Caption = iDrive.Path & "\"
        
        'Retrieve the icon associated with the drive
        vNum = IconExtraction(ImgLst, iDrive.Path & "\", [16_x_16])
        
        'Set the ImageList control associated with the Icon and SmallIcon views of the ListView control
        If Nothing Is Lv.SmallIcons Then Lv.Icons = ImgLst: Lv.SmallIcons = ImgLst
        
        'If the type of the drive is...
        Select Case iDrive.DriveType
            
            Case &H0: vBuffer(&H0) = "Unknown"
            Case &H1: vBuffer(&H0) = "Removable"
            Case &H2: vBuffer(&H0) = "Fixed"
            Case &H3: vBuffer(&H0) = "Network"
            Case &H4: vBuffer(&H0) = "CD-ROM"
            Case &H5: vBuffer(&H0) = "RAM Disk"
            
        End Select 'Close SELECT..CASE block statement

        'If the appropriate drive is inserted and ready for access then...
        If iDrive.IsReady Then
             
            'If the drive has space for writing on then...
            If FormatNumber((iDrive.FreeSpace / 1024) / 1024, &H0) <> &H0 Then
                
                vBuffer(&H1) = VBA.vbNullString 'Initialize variable
                
                'If it's a Network drive {This includes drives shared anywhere on a network}
                If iDrive.DriveType = Remote Then
                    
                    'Retrieve the network share name for the drive
                    vBuffer(&H1) = iDrive.ShareName
                    
                Else
                    
                    'Retrieve the volume name of the drive
                    vBuffer(&H1) = VBA.IIf(iDrive.VolumeName = VBA.vbNullString, "Local Disk", iDrive.VolumeName)
                    
                End If 'Close respective IF..THEN block statement
                
                Lv.ListItems.Add Lv.ListItems.Count + &H1, , vBuffer(&H1) & " (" & iDrive.DriveLetter & ")", vNum, vNum
                Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H1, , GetActualSize(iDrive.FreeSpace)
                Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H2, , GetActualSize(iDrive.TotalSize)
                Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H3, , vBuffer(&H0) & " Drive"
                Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H4, , " "
                iDriveSpace = iDrive.FreeSpace
                
            End If 'Close respective IF..THEN block statement
            
        Else
            
            Lv.ListItems.Add Lv.ListItems.Count + &H1, , vBuffer(&H0) & " (" & iDrive.DriveLetter & ")", vNum, vNum
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H1, , "Not Ready"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H2, , "Not Ready"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H3, , vBuffer(&H0) & " Drive"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H4, , " "
            Lv.ListItems(Lv.ListItems.Count).ForeColor = &HFF&
            
        End If 'Close respective IF..THEN block statement
        
        vBuffer(&H0) = VBA.vbNullString  'Initialize variable
        
        Lv.ListItems(Lv.ListItems.Count).Tag = iDrive.Path & "\"
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H1).Tag = "Drive"
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H1).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H2).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H3).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        
        iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
        Frm_PleaseWait.ImgProgressBar.Width = iCnt
        Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
        
        VBA.DoEvents: VBA.DoEvents
        
    Next iDrive 'Move to the next available drive
    
    Dim Lst As ListItem
    If lblSelectedPath.Caption <> "My Computer" Then Set Lst = Lv.FindItem(vFso.GetBaseName(lblSelectedPath.Caption)): If Not Nothing Is Lst Then Lst.EnsureVisible: Lst.Selected = True
    
    'Initialize control property
    lblSelectedPath.Caption = "My Computer"
    
    'Retrieve the icon associated with the drive
    iFolderIconIndex = IconExtraction(ImgLst, App.Path & "\Tools", [16_x_16])
    iFileIconIndex = IconExtraction(ImgLst, App.Path & "\Tools\File 28x28.ico", [16_x_16])
    
Exit_ListDrives:
    
    Lv.Visible = True: lblSelectedPath.Tag = "AutoSizer:YW|Drive"
    Unload Frm_PleaseWait
    
    vBuffer(&H0) = VBA.vbNullString
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Function
    
Handle_ListDrives_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Listing Drives - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ListDrives
    
End Function

Private Function ListFolders(Lst As ListItem) As Boolean
On Local Error GoTo Handle_ListFolders_Error
    
    Dim iCounter&, iCnt&
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
iDisplayPathFolders:
    
    'Set Column Widths
    Lv.ColumnHeaders(&H1).Width = 3350
    Lv.ColumnHeaders(&H2).Width = 1100
    Lv.ColumnHeaders(&H3).Width = 1100
    Lv.ColumnHeaders(&H4).Width = &H0
    Lv.ColumnHeaders(&H5).Width = 1000
    
'    Lv.Visible = False
    
    'Return a folder in the specified path.
    Set iFldr = vFso.GetFolder(lblSelectedPath.Caption)
    
    lblSelectedPath.Caption = lblSelectedPath.Caption
    
    'Return a collection consisting of all folders contained in the specified path,
    'including those with Hidden and System file attributes set.
    Set iSubFldrs = iFldr.SubFolders
    Set iFls = iFldr.Files
    
    iCounter = &H2 + iSubFldrs.Count + VBA.IIf(ChkDisplayFiles.Value = vbChecked, iFls.Count, &H0)
    Frm_PleaseWait.ImgProgressBar.Width = &H0
    Frm_PleaseWait.lblInfo.Caption = "Checking " & VBA.IIf(ChkDisplayFiles.Value = vbChecked, "Files and ", VBA.vbNullString) & "Folders"
    Frm_PleaseWait.ImgProgressBar.Left = Frm_PleaseWait.lblProgressBar.Left
    Frm_PleaseWait.ImgProgressBar.Visible = True
    
    'Show the 'Please Wait' Form to the User
    CenterForm Frm_PleaseWait, Me, False
    
    Lv.ListItems.Clear 'Remove all the items in the Listview control
    
    'Initialize controls' properties
    If ImgLst.ListImages.Count > &H0 Then Lv.Icons = Nothing: Lv.SmallIcons = Nothing: ImgLst.ListImages.Clear
    
    iFolderIconIndex = IconExtraction(ImgLst, App.Path & "\Tools", [16_x_16])
    
    'Set the ImageList control associated with the Icon and SmallIcon views of the ListView control
    If Nothing Is Lv.SmallIcons Then Lv.Icons = ImgLst: Lv.SmallIcons = ImgLst
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    Dim iAddtoLv As Boolean
    
    'For every subfolder...
    For Each iSubFldr In iSubFldrs
        
        Frm_PleaseWait.lblInfo.Caption = "Checking.." & VBA.Replace(vFso.GetParentFolderName(iSubFldr.Path) & "\" & iSubFldr.Name, "\\", "\") & "'"
        
        VBA.DoEvents: VBA.DoEvents
        
        iAddtoLv = False
        
        Select Case iSubFldr.Attributes
            
            Case 1, 3, 4, 5, 18, 19, 22, 23:
                
            Case Else:
                iAddtoLv = True
                
        End Select 'Close SELECT..CASE block statement
        
        If Not iAddtoLv And chkDisplayAllFolders.Value = vbUnchecked Then GoTo NextSubFolder
        
        'Display the details in the Listview control
        Lv.ListItems.Add Lv.ListItems.Count + &H1, , iSubFldr.Name, iFolderIconIndex, iFolderIconIndex
        
        If Not iAddtoLv Then
            
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H1, , "Not Ready"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H2, , "Not Ready"
            Lv.ListItems(Lv.ListItems.Count).ForeColor = &HFF&
            
        Else
            
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H1, , "N/A"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H2, , GetActualSize(iSubFldr.Size)
            Lv.ListItems(Lv.ListItems.Count).ForeColor = Lv.ForeColor
            
        End If
        
        Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H3, , " "
        Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H4, , iSubFldr.Attributes
        Lv.ListItems(Lv.ListItems.Count).Tag = iSubFldr.Path
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H1).Tag = "Folder"
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H1).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H2).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H3).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H4).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        
        iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
        Frm_PleaseWait.ImgProgressBar.Width = iCnt
        Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
        
NextSubFolder:
        
    Next iSubFldr
    
    Dim vNum&
    
    If iArrayList(iLevel - &H1) <> VBA.vbNullString And iStatus = "Moving Up.." Then
        
        Set Lst = Lv.FindItem(vFso.GetBaseName(iArrayList(iLevel - &H1)))
        If Not Nothing Is Lst Then Lst.EnsureVisible: Lst.Selected = True
        
        iLevel = iLevel - &H1
        If iLevel <> &H0 Then
        ReDim Preserve iArrayList(iLevel - &H1) As String
        End If
        
    End If
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    lblSelectedPath.Caption = iFldr.Path
    If ChkDisplayFiles.Value = vbUnchecked Then GoTo Exit_ListFolders
    
    'For every file...
    For Each iFl In iFls
        
        'Retrieve the icon associated with the drive
        vNum = IconExtraction(ImgLst, iFl.Path & "\", [16_x_16])
        
        iAddtoLv = False
        
        Select Case iFl.Attributes
            
            Case 1, 3, 4, 5, 18, 19, 22, 23:
                iAddtoLv = False
            Case Else:
                iAddtoLv = True
                
        End Select 'Close SELECT..CASE block statement
        
        'Display the details in the Listview control
        Lv.ListItems.Add Lv.ListItems.Count + &H1, , iFl.Name, vNum, vNum
        
        If Not iAddtoLv Then
            
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H1, , "Not Ready"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H2, , "Not Ready"
            Lv.ListItems(Lv.ListItems.Count).ForeColor = &HFF&
            
        Else
            
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H1, , "N/A"
            Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H2, , GetActualSize(iFl.Size)
            Lv.ListItems(Lv.ListItems.Count).ForeColor = Lv.ForeColor
            
        End If
        
        Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H3, , " "
        Lv.ListItems(Lv.ListItems.Count).ListSubItems.Add &H4, , iFl.Attributes
        Lv.ListItems(Lv.ListItems.Count).Tag = iFl.Path
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H1).Tag = "File"
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H1).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H2).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H3).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        Lv.ListItems(Lv.ListItems.Count).ListSubItems(&H4).ForeColor = Lv.ListItems(Lv.ListItems.Count).ForeColor
        
        iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
        Frm_PleaseWait.ImgProgressBar.Width = iCnt
        Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
        
NextFile:
        
    Next iFl
    
Exit_ListFolders:
    
    iStatus = VBA.vbNullString
    
    Lv.Visible = True: Lv.SetFocus
    
    lblSelectedPath.Tag = "AutoSizer:YW|Folder": Unload Frm_PleaseWait
    
    Me.Enabled = True
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Function 'Quit this Sub procedure
    
Handle_ListFolders_Error:
    
    If Err.Number = &H5 Then Resume Next
    
    Dim ErrNo&
    
    ErrNo = Err.Number
    Resume
    Unload Frm_PleaseWait
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Listing SubFolders - " & ErrNo, Me
    
    If ErrNo = 76 Then Call ListDrives
    
    'Resume execution at the specified Label
    Resume Exit_ListFolders
    
End Function

Private Sub chkDisplayAllFolders_Click()
    If lblSelectedPath.Caption <> "My Computer" Then Call ShpBttnRefresh_Click
End Sub

Private Sub ChkDisplayFiles_Click()
    If lblSelectedPath.Caption <> "My Computer" Then Call ShpBttnRefresh_Click
End Sub

Private Sub Form_Load()
    
    Me.Caption = App.Title & " : Backup Utility"
    
    ReDim Preserve iArrayList(&H0) As String
    
    Call ListDrives
    
    'Inform User on importance of Backing up data
    lblReadme.Tag = "The Backup utility helps you create a copy of the information on your hard disk. In the event that the original data on your hard disk is accidentally erased or overwritten, or becomes inaccessible because of a hard disk malfunction, you can use the copy to restore your lost or damaged data."
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Call Procedure in Mdl_Stadmis Module to confirm Form closure
    Cancel = Not CloseFrm(Me)
End Sub

Private Sub Fra_Backup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblReadme.FontUnderline Then lblReadme.FontUnderline = False
End Sub

Private Sub ImgBackupUtility_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblReadme.FontUnderline Then lblReadme.FontUnderline = False
End Sub

Private Sub ImgFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblReadme.FontUnderline Then lblReadme.FontUnderline = False
End Sub

Private Sub ImgHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblReadme.FontUnderline Then lblReadme.FontUnderline = False
End Sub

Private Sub cmdMenuOptions_Click(Index As Integer)
On Local Error GoTo Handle_cmdMenuOptions_Click_Error
    
    Static iMnuStatus$
    
    Select Case Index
        
        Case &H0: 'Change Listview View
            Me.PopupMenu MnuListviewViews
            
        Case &H1: 'Create New Folder
            
RequestFolderName:
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
            Load Frm_DataEntry
            Frm_DataEntry.LblInput.Caption = "New Folder Name"
            Frm_DataEntry.strDefault = App.Title & " Backup"
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            Frm_DataEntry.Show vbModal
            
            'If the new folder name has already been given to another folder in the specified location then...
            If vFso.FolderExists(lblSelectedPath.Caption & "\" & VBA.Trim$(vBuffer(&H0))) Then
                
                Dim vAns%
                
                'Confirm next action from User
                vAns = vMsgBox("A folder with the entered name already exists in the specified location.", vbQuestion + vbCustomButtons + vbDefaultButton1, App.Title & " : Folder Exists", Me, , , , , , "&Replace|Retry|&Cancel")
                
                'If User's response is to..
                Select Case vAns
                    
                    Case &H0: 'Cancel creation of a new folder then..
                        
                        'Branch execution to the specified Label
                        GoTo Exit_cmdMenuOptions_Click
                        
                    Case &H1: 'Request User to enter a different folder name
                        
                        'Branch execution to the specified Label
                        GoTo RequestFolderName
                        
                    Case &H2: 'Replace the existing folder with the new one
                        
                        iMnuStatus = "Replacing File..."
                        Call cmdMenuOptions_Click(&H2) 'Delete the Folder
                        iMnuStatus = VBA.vbNullString 'Initialize variable
                        
                        'If the Folder still exists then...
                        If vFso.FolderExists(lblSelectedPath.Caption & "\" & VBA.Trim$(vBuffer(&H0))) Then
                            
                            'Warn User of failed replacement
                            vMsgBox "The Folder could not be replaced. Please use a different name", vbExclamation, App.Title & " : Replace Operation Failed!!", Me
                            
                            'Branch execution to the specified Label
                            GoTo Exit_cmdMenuOptions_Click
                            
                        End If 'Close respective IF..THEN block statement
                        
                End Select
                
            End If 'Close respective IF..THEN block statement
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
            If vBuffer(&H0) <> VBA.vbNullString Then vFso.CreateFolder lblSelectedPath.Caption & "\" & VBA.Trim$(vBuffer(&H0)): Call ShpBttnRefresh_Click
            
            Erase vBuffer 'Initialize variable
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
        Case &H2: 'Delete Folder
            
            'If the item to be deleted is a Drive then...
            If lblSelectedPath.Tag = "AutoSizer:YW|Drive" Then
                
                'Warn User
                vMsgBox "Cannot delete a Drive!!", vbExclamation, App.Title & " : Operation Failed", Me
            
                'Branch to the specified Label
                GoTo Exit_cmdMenuOptions_Click
                
            End If 'Close respective IF..THEN block statement
            
            'If the folder to be deleted has not been selected then...
            If Nothing Is Lv.SelectedItem Then
                
                'Warn User
                vMsgBox "Please select a Folder to be deleted", vbExclamation, App.Title & " : Folder not specified", Me
                
                'Branch to the specified Label
                GoTo Exit_cmdMenuOptions_Click
                
            End If 'Close respective IF..THEN block statement
            
            'If the selected item is in red then...
            If Lv.SelectedItem.ForeColor = &HFF& Then
                
                'Warn User
                vMsgBox "The selected Folder {" & Lv.SelectedItem.Text & "} cannot be deleted!! It is either a System or a Readonly Folder", vbExclamation, App.Title & " : Operation Failed", Me
                
                'Branch to the specified Label
                GoTo Exit_cmdMenuOptions_Click
                
            End If 'Close respective IF..THEN block statement
            
            If VBA.Val(Lv.SelectedItem.ListSubItems(&H2)) <> &H0 Then
                
                'Confirm deletion
                If vMsgBox("The selected Folder {" & Lv.SelectedItem.Text & "} contains data. Are you sure you want to DELETE it?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then
                    
                    'Branch to the specified Label
                    GoTo Exit_cmdMenuOptions_Click
                    
                End If 'Close respective IF..THEN block statement
                
            Else
                
                If iMnuStatus <> "Replacing File..." Then
                    
                    'Confirm deletion
                    If vMsgBox("Are you sure you want to DELETE the selected Folder? {" & Lv.SelectedItem.Text & "}", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then
                        
                        'Branch to the specified Label
                        GoTo Exit_cmdMenuOptions_Click
                        
                    End If 'Close respective IF..THEN block statement
                    
                End If 'Close respective IF..THEN block statement
                
            End If 'Close respective IF..THEN block statement
            
            'Check if the specified folder exists. If so then...
            If vFso.FolderExists(VBA.Replace(lblSelectedPath.Caption & "\" & Lv.SelectedItem.Text, "\\", "\")) Then
                
                'Delete the selected Folder
                vFso.DeleteFolder VBA.Replace(lblSelectedPath.Caption & "\" & Lv.SelectedItem.Text, "\\", "\"), True
                Call ShpBttnRefresh_Click 'Refresh details
                
                If iMnuStatus <> "Replacing File..." Then
                    
                    'Indicate that a process or operation is complete.
                    Screen.MousePointer = vbDefault
                    
                    'Inform User
                    vMsgBox "Folder succefully deleted", vbInformation, App.Title & " : Folder Deletion", Me
                    
                End If 'Close respective IF..THEN block statement
                
            Else 'If the specified folder does not exist, and initially it existed then...
                
                Call ShpBttnRefresh_Click 'Refresh details
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "It seems like the Folder has already been deleted from the location", vbInformation, App.Title & " : Folder Deletion", Me
                
            End If 'Close respective IF..THEN block statement
            
        Case &H3: 'Move One Level Up
            Call Lv_KeyDown(vbKeyBack, &H0)
            
    End Select
    
Exit_cmdMenuOptions_Click:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub procedure
    
Handle_cmdMenuOptions_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    If Err.Number = 76 And Index = &H1 Then
        
        'Warn User
        vMsgBox "Cannot create Folder on the specified Location", vbExclamation, App.Title & " : Error Creating Folder - " & Err.Number, Me
        
    ElseIf Err.Number = 52 And Index = &H1 Then
        
        'Warn User
        vMsgBox "Folder name cannot contain the following special characters: \ / : * ? "" < > ", vbExclamation, App.Title & " : Error Creating Folder - " & Err.Number, Me
        
    Else
        
        'Warn User
        vMsgBox Err.Description, vbExclamation, App.Title & " : Error on Icon '" & cmdMenuOptions(Index).ToolTipText & "' - " & Err.Number, Me
        
    End If 'Close respective IF..THEN block statement
    
    'Resume execution at the specified Label
    Resume Exit_cmdMenuOptions_Click
    
End Sub

Private Sub lblInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If lblReadme.FontUnderline Then lblReadme.FontUnderline = False
End Sub

Private Sub lblReadme_Click()
    
    lblReadme.FontUnderline = True
    
    'Call Function in Mdl_Stadmis to display Backup Utility Information
    Call OpenNotes(lblReadme, True, "Backup Utility", , "Why Backup data?")
    
End Sub

Private Sub lblReadme_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblReadme.BorderStyle = &H1 'Create a click effect
End Sub

Private Sub lblReadme_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lblReadme.FontUnderline Then lblReadme.FontUnderline = True
End Sub

Private Sub lblReadme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblReadme.BorderStyle = &H0 'Remove the click effect
End Sub

Private Sub lblSelectedPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim iNum&
    
    For iNum = &H1 To VBA.Len(lblSelectedPath.Caption) Step &H1
        If Me.TextWidth(VBA.Left$(lblSelectedPath.Caption, iNum)) >= X Then GoTo CheckText
    Next iNum
    
    Exit Sub
    
CheckText:
    
    vArrayListTmp = VBA.Split(lblSelectedPath.Caption, "\")
    vArrayList = VBA.Split(VBA.Left$(lblSelectedPath.Caption, iNum), "\")
    
    ReDim Preserve vArrayListTmp(UBound(vArrayList)) As String
    lblSelectedPath.Caption = VBA.Join(vArrayListTmp, "\")
    Call ShpBttnRefresh_Click: Me.Enabled = True
    
End Sub

Private Sub Lv_DblClick()
On Local Error GoTo Handle_Lv_DblClick_Error
    
    Me.Enabled = False
    
MoveUp:
    
    Dim iPrevPath$
    Dim Lst As ListItem
    
    If iStatus = "Refreshing.." Then
        
        If lblSelectedPath.Tag = "AutoSizer:YW|Drive" Then
            
            Call ListDrives 'call Sub procedure to list all the drives in the Computer
            Exit Sub 'Quit this Sub procedure
            
        Else
            
            iStatus = VBA.vbNullString
            Call ListFolders(Lst) 'Display contents of a folder one level up
            
        End If 'Close respective IF..THEN block statement
        
    End If 'Close respective IF..THEN block statement
    
    If iStatus = "Moving Up.." Then
        
        Dim iArray() As String
        
        iPrevPath = lblSelectedPath.Caption
        iArray = VBA.Split(lblSelectedPath.Caption, "\")
        
        'If the parent is a folder then...
        If UBound(iArray) >= &H1 Then
            
            If iArray(&H1) = VBA.vbNullString Then
                
                Call ListDrives 'call Sub procedure to list all the drives in the Computer
                
                iLevel = &H1
                ReDim Preserve iArrayList(iLevel - &H1) As String
                
                Set Lst = Lv.FindItem(vFso.GetBaseName(iArrayList(iLevel - &H1)))
                If Not Nothing Is Lst Then Lst.EnsureVisible: Lst.Selected = True: Lv.Visible = True: Lv.SetFocus
                
            Else
                
                ReDim Preserve iArray(UBound(iArray) - &H1)
                If UBound(iArray) > &H0 Then lblSelectedPath.Caption = VBA.Join(iArray, "\") Else lblSelectedPath.Caption = iArray(&H0)
                lblSelectedPath.Caption = lblSelectedPath.Caption & VBA.IIf(UBound(iArray) = &H0, "\", VBA.vbNullString)
                
                Call ListFolders(Lst) 'Display contents of a folder one level up
                
            End If 'Close respective IF..THEN block statement
            
        Else 'If the parent is a drive then...
            
            If lblSelectedPath.Tag <> "Drive" Then Call ListDrives  'call Sub procedure to list all the drives in the Computer
            Lv.SetFocus
            
        End If 'Close respective IF..THEN block statement
        
        Me.Enabled = True
        
        Exit Sub 'Quit this Sub procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If no item has been double-clicked then quit this Sub procedure
    If Nothing Is Lv.SelectedItem Then Exit Sub
    
    'If a file has been double-clicked then quit this Sub procedure
    If Lv.SelectedItem.ListSubItems(&H1).Tag = "File" Then Exit Sub
    
    'If the selected item is in red then...
    If Lv.SelectedItem.ForeColor = &HFF& Then
        
        vMsgBox "The selected " & Lv.SelectedItem.ListSubItems(&H1).Tag & " is inaccessible", vbExclamation, App.Title & " : Invalid Backup Location", Me
        Me.Enabled = True
        Exit Sub 'Quit this Sub procedure
        
    End If 'Close respective IF..THEN block statement
    
    iLevel = iLevel + &H1
    
    If UBound(iArrayList) < iLevel - &H1 Then ReDim Preserve iArrayList(iLevel - &H1) As String
    
    iArrayList(iLevel - &H1) = Lv.SelectedItem.Text
    lblSelectedPath.Caption = Lv.SelectedItem.Tag
    
    Dim iCounter&, iCnt&
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Call ListFolders(Lst)
    
Exit_Lv_DblClick:
    
    iStatus = VBA.vbNullString
    
    Me.Enabled = True
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub procedure
    
Handle_Lv_DblClick_Error:
    
    If Err.Number = &H5 Then Resume Next
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Listing SubFolders - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_Lv_DblClick
    
End Sub

Private Sub Lv_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Call Sub procedure to display contents of a folder one level up
    If KeyCode = vbKeyBack Then iStatus = "Moving Up..": Call Lv_DblClick: KeyCode = Empty
    
    'Display the Subfolders contained in the selected drive/folder
    If KeyCode = vbKeyReturn Then Stop: Call Lv_DblClick
    
End Sub

Private Sub Lv_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Stop
End Sub

Private Sub Lv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Stop
End Sub

Private Sub Lv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iStatus = VBA.IIf(Nothing Is Lv.HitTest(X, Y), "Moving Up..", VBA.vbNullString)
End Sub

Private Sub Lv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblReadme.FontUnderline Then lblReadme.FontUnderline = False
End Sub

Private Sub MnuListviewView_Click(Index As Integer)
    Lv.View = Index
End Sub

Private Sub ShpBttnBackup_Click()
On Local Error GoTo Handle_ShpBttnBackup_Click_Error
    
    'If the folder to be deleted has not been selected then...
    If Nothing Is Lv.SelectedItem Then
        
        'Warn User
        vMsgBox "Please select a Backup Folder", vbExclamation, App.Title & " : Folder not specified", Me
        Exit Sub 'Quit this Sub procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the selected item is in red then...
    If Lv.SelectedItem.ForeColor = &HFF& Then
        
        vMsgBox "The selected " & Lv.SelectedItem.ListSubItems(&H1).Tag & " is inaccessible", vbExclamation, App.Title & " : Invalid Backup Location", Me
        Exit Sub 'Quit this Sub procedure
        
    End If 'Close respective IF..THEN block statement
    
    'If the selected location is the Software's location for its data storage then...
    If VBA.Left$(lblSelectedPath.Caption, VBA.Len(App.Path & "\Application Data")) = App.Path & "\Application Data" Then
        
        'Warn User
        vMsgBox "The selected location {" & lblSelectedPath.Caption & "} has been reserved for the Software. Please choose another backup location", vbExclamation, App.Title & " : Invalid Backup Location", Me
        Exit Sub 'Quit this Sub procedure
        
    End If 'Close respective IF..THEN block statement
    
    'Confirm deletion. If No then quit this Sub procedure
    If vMsgBox("Are you sure you want to Backup data to the specified location '" & vFso.GetParentFolderName(VBA.Replace(lblSelectedPath.Caption & "\" & Lv.SelectedItem.Text, "\\", "\")) & "'?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirmation", Me) = vbNo Then Exit Sub
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Dim vNum&, iCounter&, iCnt&
    
    iCounter = &H8
    Frm_PleaseWait.ImgProgressBar.Width = &H0
    Frm_PleaseWait.ImgProgressBar.Left = Frm_PleaseWait.lblProgressBar.Left
    Frm_PleaseWait.ImgProgressBar.Visible = True
    
    'Show the 'Please Wait' Form to the User
    CenterForm Frm_PleaseWait, , False
    
    '--------------------------------------------------------------------------------------------
    '                                Retrieve Software Settings
    '--------------------------------------------------------------------------------------------
    
    Dim vRow&
    Dim iSoftwareSettings$
    
    vBuffer(&H0) = "Settings"
    vArrayList = VBA.GetAllSettings(App.Title, vBuffer(&H0))
    For vRow = &H0 To UBound(vArrayList) Step &H1
        vBuffer(&H0) = vBuffer(&H0) & "|" & vArrayList(vRow, &H0) & "-" & vArrayList(vRow, &H1)
    Next vRow
    
    iSoftwareSettings = vBuffer(&H0) & "||"
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    vBuffer(&H0) = "Main Form"
    vArrayList = VBA.GetAllSettings(App.Title, vBuffer(&H0))
    For vRow = &H0 To UBound(vArrayList) Step &H1
        vBuffer(&H0) = vBuffer(&H0) & "|" & vArrayList(vRow, &H0) & "-" & vArrayList(vRow, &H1)
    Next vRow
    
    iSoftwareSettings = iSoftwareSettings & vBuffer(&H0) & "||"
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    vBuffer(&H0) = "Copyright Protection"
    vArrayList = VBA.GetAllSettings(App.Title, vBuffer(&H0))
    For vRow = &H0 To UBound(vArrayList) Step &H1
        vBuffer(&H0) = vBuffer(&H0) & "|" & vArrayList(vRow, &H0) & "-" & vArrayList(vRow, &H1)
    Next vRow
    
    iSoftwareSettings = iSoftwareSettings & vBuffer(&H0) & "||"
    vBuffer(&H0) = VBA.vbNullString
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    If VBA.Right$(iSoftwareSettings, &H2) = "||" Then iSoftwareSettings = VBA.Left$(iSoftwareSettings, VBA.Len(iSoftwareSettings) - &H2)
    
    '--------------------------------------------------------------------------------------------
    '                                  Save Software Settings
    '--------------------------------------------------------------------------------------------
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    If VBA.LenB(VBA.Trim$(vAdoCNN.ConnectionString)) = &H0 Then Call ConnectDB 'Call Function in Mdl_Stadmis Module to connect to Software's Database
    Set vRs = New ADODB.Recordset
    With vRs 'Execute a series of statements on vRs recordset
        
        .Open "SELECT * FROM [Tbl_Setup]", vAdoCNN, adOpenKeyset, adLockPessimistic
        
        'If there are records in the table then...
        If Not (.BOF And .EOF) Then .Update Else .AddNew
        
        ![Backup Settings] = iSoftwareSettings
        
        .Update
        .UpdateBatch adAffectAllChapters
        
        .Close 'Close the opened object and any dependent objects
        
    End With 'Close the WITH block statements
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    Dim iDBPath$, iDBName$, iDestDBName$, iBkDBName$
    
    iDBPath = vAdoCNN.Properties("Data Source")
    
    Call PerformMemoryCleanup
    
    iDBName = App.Title & " Backup - " & VBA.Format$(VBA.Now, "yyyyMMddHHnnss")
    iBkDBName = App.Path & "\Application Data\Copies of Backups\" & iDBName & " - Saved by " & User.Full_Name & "." & def_BackupDatabaseExt
    iDestDBName = VBA.Replace(lblSelectedPath.Caption & "\" & iDBName & "." & def_BackupDatabaseExt, "\\", "\")
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    If vFso.FileExists(iDBPath) Then
        
        If Not vFso.FolderExists(vFso.GetParentFolderName(iDestDBName)) Then Call CreatePath(iDestDBName)
        vFso.CopyFile iDBPath, iDestDBName, True
        
        'If the database has not been backed up then...
        If Not vFso.FileExists(iDestDBName) Then
            
            VBA.DoEvents: VBA.DoEvents
            
            Unload Frm_PleaseWait
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Inform User
            vMsgBox "The Software Backup has not successfully been saved", vbExclamation, App.Title & " : Data Backup Failed!!", Me
            GoTo Exit_ShpBttnBackup_Click 'Resume execution at the specified Label
            
        End If 'Close respective IF..THEN block statement
        
        If Not vFso.FolderExists(vFso.GetParentFolderName(iBkDBName)) Then Call CreatePath(iBkDBName)
        vFso.CopyFile iDBPath, iBkDBName, True
        
    End If 'Close respective IF..THEN block statement
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    Call ShpBttnRefresh_Click 'Refresh details to show the newly backed up data
    
    iCnt = iCnt + (Frm_PleaseWait.lblProgressBar.Width / iCounter)
    Frm_PleaseWait.ImgProgressBar.Width = iCnt
    Frm_PleaseWait.lblStatus.Caption = VBA.Int(((VBA.CLng(iCnt) * 100) / Frm_PleaseWait.lblProgressBar.Width)) & "% Complete"
    
    VBA.DoEvents: VBA.DoEvents
    
    Unload Frm_PleaseWait
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Inform User
    vMsgBox "Software Backup successfully saved at '" & iBkDBName & "'", vbInformation, App.Title & " : Data Backup", Me
    
Exit_ShpBttnBackup_Click:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub procedure
    
Handle_ShpBttnBackup_Click_Error:
    
    Unload Frm_PleaseWait
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Backing Up Data - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_ShpBttnBackup_Click
    
End Sub

Private Sub ShpBttnCancel_Click()
    Unload Me 'Unload the form from memory
End Sub

Private Sub ShpBttnRefresh_Click()
    iStatus = "Refreshing..": Call Lv_DblClick
End Sub
