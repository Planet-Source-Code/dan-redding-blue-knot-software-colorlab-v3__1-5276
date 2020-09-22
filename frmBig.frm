VERSION 5.00
Begin VB.Form frmBig 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "The Big Picture"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picImg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   60
      Picture         =   "frmBig.frx":0000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   60
      Width           =   2925
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   1800
      Width           =   555
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help!"
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   1800
      Width           =   555
   End
End
Attribute VB_Name = "frmBig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strLastPicDir As String, lWid As Long, lHei As Long, lGap As Long

Private Sub cmdHelp_Click()
    MsgBox "Double-click image to load a new .JPG or .GIF" & vbNewLine & _
        vbTab & "for viewing against background color." & vbNewLine & _
        vbTab & "(supports transparent GIFFs but only" & vbNewLine & _
        vbTab & " first frame of animated GIFFs.)" & vbNewLine & _
        "Right-click image to set background color to" & vbNewLine & _
        vbTab & "average color of the image.", vbInformation + vbOKOnly, _
        "What does this window do?"
End Sub

Private Sub cmdHide_Click()
    Me.Visible = False
    frmColorRef.chkBig = False
End Sub

Private Sub Form_DblClick()
    picImg_DblClick
End Sub

Public Sub Pos()
    Me.top = frmColorRef.top
    If Screen.Width - (frmColorRef.left + frmColorRef.Width) <= (Me.Width + 4) Then
        Me.left = frmColorRef.left - Me.Width - 4 * Screen.TwipsPerPixelX
    Else
        Me.left = frmColorRef.left + frmColorRef.Width + 4 * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub Form_Load()
    lWid = picImg.left * 2 + picImg.Width
    lHei = picImg.top * 2 + picImg.Height
    lGap = Me.Height / Screen.TwipsPerPixelX - Me.ScaleHeight + cmdHide.Height
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width / Screen.TwipsPerPixelX < lWid Then Me.Width = lWid * Screen.TwipsPerPixelX
    If Me.Height / Screen.TwipsPerPixelY < lHei + lGap Then Me.Height = (lHei + lGap) * Screen.TwipsPerPixelY
    cmdHide.top = (Me.ScaleHeight) - cmdHide.Height
    cmdHide.left = (Me.ScaleWidth) - (cmdHide.Width)
    cmdHelp.top = cmdHide.top
    cmdHelp.left = cmdHide.left - cmdHelp.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    frmColorRef.chkBig = False
    If Not blnEnd Then Cancel = True
End Sub


Private Sub picImg_DblClick()
Dim sFile As SelectedFile, strFavs As String
    FileDialog.sDlgTitle = "Choose Big Picture Image"
    FileDialog.flags = OFS_FILE_OPEN_FLAGS
    FileDialog.sFilter = "GIF Files(*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "JPEG Files (*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & Chr$(0)
    FileDialog.sDefFileExt = "*.gif"
    sFile = ShowOpen(frmColorRef.hwnd, True, "", strLastPicDir)
    If sFile.bCanceled Then Exit Sub
    strLastPicDir = sFile.sLastDirectory
    On Error GoTo BadPic
    picImg.AutoSize = True
    picImg.Picture = LoadPicture(strLastPicDir & sFile.sFiles(1))
    picImg.AutoSize = False
    Debug.Print picImg.Width; " "; picImg.ScaleWidth + 1
    picImg.Width = picImg.ScaleWidth + 1
    picImg.Height = picImg.ScaleHeight + 1
    lWid = picImg.left * 2 + picImg.Width
    lHei = picImg.top * 2 + picImg.Height
    If Me.Width < lWid Then Me.Width = lWid
    If Me.Height / Screen.TwipsPerPixelY < lHei + lGap Then Me.Height = (lHei + lGap) * Screen.TwipsPerPixelY
    Exit Sub
BadPic:
    MsgBox "Can't Load Picture." & vbNewLine & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub picImg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lX As Long, lY As Long, lColor As Long, lTran As Long
Dim dVR As Double, dVG As Double, dVB As Double, lVCount As Long
'Dim dHR As Double, dHG As Double, dHB As Double, lHCount As Long
Dim lFinalColor As Long
    If Button = vbRightButton Or Shift = vbShiftMask Then
        Me.Caption = "Averaging Color from Picture..."
        lTran = picImg.Point(picImg.Width - 1, picImg.Height - 1)
        For lX = 0 To picImg.Width - 1
            For lY = 0 To picImg.Height - 1
                lColor = picImg.Point(lX, lY)
                If lColor <> -1 And lColor <> lTran Then
                    dVR = dVR + RGBRed(lColor)
                    dVG = dVG + RGBGreen(lColor)
                    dVB = dVB + RGBBlue(lColor)
                    lVCount = lVCount + 1
                End If
            Next lY
        Next lX
        lFinalColor = RGB(CLng(dVR / lVCount), CLng(dVG / lVCount), CLng(dVB / lVCount))
        frmColorRef.ChangeColor lFinalColor
        Me.Caption = "The Big Picture"
    End If
End Sub

