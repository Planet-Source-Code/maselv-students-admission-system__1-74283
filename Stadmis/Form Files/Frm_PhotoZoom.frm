VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_PhotoZoom 
   BackColor       =   &H00CFE1E2&
   Caption         =   "App Title : Popup Menus"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4095
   Icon            =   "Frm_PhotoZoom.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Stadmis.AutoSizer AutoSizer 
      Left            =   1800
      Top             =   3960
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   -600
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   -1080
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Fra_PopupMenu 
      BackColor       =   &H00CFE1E2&
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Tag             =   "AutoSizer:HW"
      Top             =   360
      Width           =   3855
      Begin VB.Image ImgDBPhoto 
         Height          =   3135
         Left            =   0
         Stretch         =   -1  'True
         Tag             =   "AutoSizer:C"
         Top             =   120
         Width           =   3855
      End
      Begin VB.Image ImgVirtualPhoto 
         Height          =   135
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Info:"
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
      TabIndex        =   0
      Tag             =   "AutoSizer:Y"
      Top             =   3960
      Width           =   990
   End
   Begin VB.Image ImgHeader 
      Height          =   360
      Left            =   0
      Picture         =   "Frm_PhotoZoom.frx":038A
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:W"
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image ImgFooter 
      Height          =   735
      Left            =   0
      Picture         =   "Frm_PhotoZoom.frx":0C2A
      Stretch         =   -1  'True
      Tag             =   "AutoSizer:WY"
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Menu MnuPhotos 
      Caption         =   "Photo Menus"
      Begin VB.Menu MnuPhoto 
         Caption         =   "&Attach Photo"
         Index           =   0
      End
      Begin VB.Menu MnuPhoto 
         Caption         =   "&Remove Photo"
         Index           =   1
      End
      Begin VB.Menu MnuPhoto 
         Caption         =   "&Maximize Photo"
         Index           =   2
      End
      Begin VB.Menu MnuPhoto 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuPhoto 
         Caption         =   "Save Image As..."
         Index           =   4
      End
      Begin VB.Menu MnuPhoto 
         Caption         =   "Copy Image to Clipboard"
         Index           =   5
      End
      Begin VB.Menu MnuPhoto 
         Caption         =   "Set Image as Desktop Wallpaper"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Frm_PhotoZoom"
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

Public DestImg As Image
Public SourceImg As Image
Public DestImgOutline As Object
Public PhotoIndex&, ImgHeight&, ImgWidth&

Private FrmHeight%, FrmWidth%
Private Resize, LoadingImage As Boolean

Private Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long

Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPI_SETDESKWALLPAPER = &H14
Private Const SPIF_SENDWININICHANGE = &H2

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    Dim PercZoom&
    
    'Get Zoom percentage
    PercZoom = (ImgDBPhoto.Height * 100) / DestImg.Height
    
    'Display the image name and dimensions
    lblInfo.Caption = "Name: " & SourceImg.ToolTipText & VBA.vbCrLf & _
                      "Dimensions: " & ImgVirtualPhoto.Width & " x " & ImgVirtualPhoto.Height & " {" & PercZoom & "%}"
    
    AutoSizer.Enabled = True
    AutoSizer.GetInitialPositions
    Call AutoSizer.AutoResize
    
End Sub

Private Sub Form_Load()
    If AutoSizer.Enabled Then vMsgBox "Please set the enabled property of the Autosizer control to False", vbExclamation, App.Title & " : Invalid Initial Setting"
End Sub

Private Sub Form_Resize()
    
    Dim PercZoom&
    
    'Get Zoom percentage
    PercZoom = (ImgDBPhoto.Height * 100) / DestImg.Height
    
    'Display the image name and dimensions
    lblInfo.Caption = "Name: " & SourceImg.ToolTipText & VBA.vbCrLf & _
                      "Dimensions: " & ImgVirtualPhoto.Width & " x " & ImgVirtualPhoto.Height & " {" & PercZoom & "%}"
    
End Sub

Private Sub ImgDBPhoto_DblClick()
    
    'Toggle window between Maximized and Normal
    Me.WindowState = VBA.IIf(Me.WindowState = &H0, &H2, &H0)
    
End Sub

Private Sub ImgDBPhoto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Public Sub MnuPhoto_Click(Index As Integer)
On Local Error GoTo Handle_MnuPhoto_Click_Error
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        
        Case &H0: 'Attach/Replace specified control's Photo
            
            'Call Function to enbale User select an Image file from the Computer
            AttachPhoto DestImg.Parent, SourceImg, Me.Dlg, , ImgHeight, ImgWidth, PhotoIndex
            
            'Call Procedure to Fit image to image holder
            If SourceImg.Picture <> &H0 Then If Not Nothing Is DestImgOutline Then Call FitPicTo(SourceImg, DestImg, DestImgOutline) Else DestImg.Picture = SourceImg.Picture
            
        Case &H1: 'Remove displayed Photo
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Confirm if User wants to remove the displayed Photo, if No then quit this procedure
            If vMsgBox("Are you sure you want to remove the displayed Photo?", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " : Confirming Deletion", SourceImg.Parent) = vbNo Then Exit Sub
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
            'Clear Photo
            SourceImg.Picture = Nothing: DestImg.Picture = Nothing
            
            On Local Error Resume Next
            SourceImg.Move DestImgOutline.Left, DestImgOutline.Top, DestImgOutline.Width, DestImgOutline.Height
            
        Case &H2: 'Maximize displayed Photo
            
            Dim FrmHeading As String 'Holds the initial Form's Caption
            
            FrmHeading = Me.Caption 'Assign initial Form's Caption
            
            'Change caption to display name of the selected Picture
            Me.Caption = App.Title & " : " & SourceImg.ToolTipText
            
            With ImgDBPhoto 'Execute a series of statements for ImgMaximized control
                
                Me.MnuPhotos.Visible = False 'Hide Photo Menu
                
                .Picture = SourceImg.Picture 'Assign the Source Picture to this Image control
                ImgVirtualPhoto.Picture = .Picture
                
                Dim PercZoom&
                
                .Top = 200
                .Left = .Top - 50
                AutoSizer.Enabled = False
                
                'Zoom 150% into the Picture
                .Height = SourceImg.Height * 1.5
                PercZoom = (.Height * 100) / DestImg.Height
                .Width = SourceImg.Width * 1.5
                
                Fra_PopupMenu.Height = .Height + (.Top * &H2) - 50
                Fra_PopupMenu.Width = .Width + (.Left * &H2)
                
                ImgHeader.Width = Fra_PopupMenu.Left + Fra_PopupMenu.Width + 100
                ImgFooter.Width = ImgHeader.Width
                ImgFooter.Top = Fra_PopupMenu.Top + Fra_PopupMenu.Height + 100
                
                Me.Height = ImgFooter.Top + ImgFooter.Height + 800
                Me.Width = ImgFooter.Width + 100
                
                'If the Form height exceeds the Screen then...'Fit the image height to the screen and..
                If Me.Height > Screen.Height Then
                    
                    Dim MyH
                    
                    MyH = Me.Height
                    
                    Me.Height = Screen.Height - 5000
                    Me.Width = (Me.Height * Me.Width) / MyH
                    
                    ImgFooter.Top = Me.Height - (ImgFooter.Height + 800)
                    ImgFooter.Width = Me.Width - 100
                    ImgHeader.Width = ImgFooter.Width
                    Fra_PopupMenu.Width = Me.Width - Fra_PopupMenu.Left - 300
                    Fra_PopupMenu.Height = ImgFooter.Top - (Fra_PopupMenu.Top + 200)
                    
                    .Width = Fra_PopupMenu.Width - .Left
                    
                    Call FitPicTo(SourceImg, ImgDBPhoto, Fra_PopupMenu)
                    
                End If
                
                'Display the image name and dimensions
                lblInfo.Caption = "Name: " & SourceImg.ToolTipText & VBA.vbCrLf & _
                                  "Dimensions: " & ImgVirtualPhoto.Width & " x " & ImgVirtualPhoto.Height & " {" & PercZoom & "%}"
                
                lblInfo.Width = ImgFooter.Width
                lblInfo.Top = ImgFooter.Top + ((ImgFooter.Height / &H2) - (lblInfo.Height / &H2))
                
'                Call AutoSizer.GetInitialPositions
                AutoSizer.Enabled = True
                
                'Get current Form's dimensions
                FrmWidth = Me.Width: FrmHeight = Me.Height
                Call AutoSizer.AutoResize
                
                Resize = True 'Allow Form to resizing control when it's being resized
                
                'Hide Picture manipulation Menus
                MnuPhoto(&H0).Visible = False: MnuPhoto(&H1).Visible = False
                MnuPhoto(&H2).Visible = False: MnuPhoto(&H3).Visible = False
                MnuPhotos.Visible = True
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                CenterForm Me, SourceImg.Parent  'Display the Image to the User
                
                'Indicate that a process or operation is in progress.
                Screen.MousePointer = vbHourglass
                
                lblInfo.Caption = VBA.vbNullString 'Clear image name
                
            End With 'End WITH statement
            
            Me.Caption = FrmHeading 'Restore Form's caption
            
        Case &H4: 'Save Image As...
SaveAsLocation:
            Load Frm_DataEntry
            
            vBuffer(&H0) = VBA.vbNullString
            Frm_DataEntry.Caption = App.Title & " : Save Image As.."
            Frm_DataEntry.LblInput.Caption = "File Location:"
            
            Frm_DataEntry.mIsBrowser = True
            Frm_DataEntry.mDialogAction = &H1 'Save As Dialog
            Frm_DataEntry.txtEntry.Tag = SourceImg.ToolTipText
            
            '1. The specified Path MUST exist
            '2. Generate a message box if the selected file already exists.
            '    The user must confirm whether to overwrite the file
            '3. The returned file shouldn't have Read Only attribute set and won't be in a write-protected Folder.
            
            'Set the filters that are displayed in the Type list box of the CommonDialog control
            Frm_DataEntry.Dlg.Filter = "JPEG (*.JPG,*.JPEG,*.JPE,*.JFIF)|*.JPG;*.JPEG;*.JPE;*.JFIF|Bitmap (*.bmp,*.dib)|*.bmp;*.dib|GIF (*.GIF)|*.gif|PNG (*.PNG)|*.png|TIFF (*.TIF,*.TIFF)|*.tif;*.tiff"
            
            'cdlOFNPathMustExist, cdlOFNOverwritePrompt, CdlOFNNoReadOnlyReturn and cdlOFNHelpButton
            Frm_DataEntry.Dlg.FLAGS = &H800 + &H2 + &H8000 + &H10
            
            Call CenterForm(Frm_DataEntry)
            
            'If the Image has been saved then...
            If vBuffer(&H0) <> VBA.vbNullString And vBuffer(&H0) <> "Cancelled" Then
                
                'If an image with the same filename already exists then...
                If vFso.FileExists(vBuffer(&H0) & ".jpg") Then
                    
                    'If User does not want to replace it then...
                    vIndex(&H0) = vMsgBox("Would you like to replace the existing file?", vbQuestion + vbCustomButtons + vbDefaultButton2, App.Title & " : ", SourceImg.Parent, , , , , , "&Yes|&Rename|&Cancel")
                    
                    Select Case vIndex(&H0)
                        
                        Case &H0: GoTo Exit_MnuPhoto_Click 'Cancelled
                        Case &H1: GoTo SaveAsLocation 'Rename...
                        Case &H2: vFso.DeleteFile vBuffer(&H0) & ".jpg" 'Replace
                        
                    End Select 'Close SELECT..CASE block statement
                    
                End If 'Close respective IF..THEN block statement
                
                'Save the Image in the specified Location
                VB.SavePicture SourceImg.Picture, vBuffer(&H0) & ".jpg"
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "The Image has successfully been saved.", vbInformation, App.Title & " : Photo Zoom", SourceImg.Parent
                
                'Indicate that a process or operation is in progress.
                Screen.MousePointer = vbHourglass
                
            End If 'Close respective IF..THEN block statement
            
            vBuffer(&H0) = VBA.vbNullString
            
        Case &H5: 'Copy Image to Clipboard...
            
            'Discard all initial entries in the Clipboard
            VB.Clipboard.Clear
            
            'Put the Image on the Clipboard object
            VB.Clipboard.SetData ImgDBPhoto.Picture
            
            'Indicate that a process or operation is complete.
            Screen.MousePointer = vbDefault
            
            'Inform User
            vMsgBox "The Image has successfully been copied to ClipBoard.", vbInformation, App.Title & " : Photo Zoom", SourceImg.Parent
            
            'Indicate that a process or operation is in progress.
            Screen.MousePointer = vbHourglass
            
        Case &H6: 'Set Image as Desktop Wallpaper...
            
            'Save the Image in the specified Location
            VB.SavePicture ImgDBPhoto.Picture, App.Path & "\Wallpaper.jpg"
            
            'If the Image has successfully been saved at the Application's default location then...
            If vFso.FileExists(App.Path & "\Wallpaper.jpg") Then
                
                'Set the Image as the Desktop's Wallpaper
                vIndex(&H0) = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, App.Path & "\Wallpaper.jpg", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
                
                'Delete the Image from the saved location
                vFso.DeleteFile App.Path & "\Wallpaper.jpg", True
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Inform User
                vMsgBox "The Image has successfully been set as the Desktop's Wallpaper.", vbInformation, App.Title & " : Photo Zoom", SourceImg.Parent
                
                'Indicate that a process or operation is in progress.
                Screen.MousePointer = vbHourglass
                
                'Re-initialize elements of the fixed-size array and release dynamic-array storage space.
                Erase vIndex
                
            Else 'If the Image has not been saved at the Application's default location then...
                
                'Indicate that a process or operation is complete.
                Screen.MousePointer = vbDefault
                
                'Warn User
                vMsgBox "Unable to set the Image as the Desktop's Wallpaper.", vbExclamation, App.Title & " : Photo Zoom", SourceImg.Parent
                
                'Indicate that a process or operation is in progress.
                Screen.MousePointer = vbHourglass
                
            End If 'Close respective IF..THEN block statement
            
    End Select 'Close SELECT..CASE block statement 'End Select 'Close SELECT..CASE block statement..CASE statement
    
    'If this Form is not visible to the User then unload it from the Memory
    If Not Me.Visible Then Unload Me
    
Exit_MnuPhoto_Click:
    
    'Change Mouse Pointer to its initial state
    Screen.MousePointer = MousePointerState
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_MnuPhoto_Click_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error Displaying Photo - " & Err.Number, SourceImg.Parent
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Resume execution at the specified Label
    Resume Exit_MnuPhoto_Click
    
End Sub

Private Sub Resizer_AfterResize()
On Local Error GoTo Handle_Resizer_AfterResize_Error
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    'Call Procedure to Fit image to image holder
    Call FitPicTo(ImgVirtualPhoto, ImgDBPhoto, Fra_PopupMenu)
    
Exit_Resizer_AfterResize:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    Exit Sub 'Quit this Sub-Procedure
    
Handle_Resizer_AfterResize_Error:
    
    'Indicate that a process or operation is complete.
    Screen.MousePointer = vbDefault
    
    'Warn User
    vMsgBox Err.Description, vbExclamation, App.Title & " : Error AfterResize - " & Err.Number, Me
    
    'Resume execution at the specified Label
    Resume Exit_Resizer_AfterResize
    
End Sub
