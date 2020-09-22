VERSION 5.00
Begin VB.Form frmFav 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ColorLab Favorites"
   ClientHeight    =   4470
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeFav 
      Caption         =   "Change"
      Height          =   315
      Left            =   1215
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Change the color for the selected favorite"
      Top             =   3900
      Width           =   1140
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   315
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Rename the selected favorite"
      Top             =   3900
      Width           =   1140
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   2100
      Picture         =   "frmFav.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Send to HTML Chart"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   1800
      Picture         =   "frmFav.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Send to Printer"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   1500
      Picture         =   "frmFav.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Send List to Clipboard"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   1200
      Picture         =   "frmFav.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Clear List"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   900
      Picture         =   "frmFav.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Merge"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   600
      Picture         =   "frmFav.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Save As"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   300
      Picture         =   "frmFav.frx":07BC
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Save"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdTool 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "frmFav.frx":0906
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Open"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Delete the selected favorite"
      Top             =   3600
      Width           =   765
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "Use"
      Height          =   315
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Make the selected favorite the current color"
      Top             =   3600
      Width           =   765
   End
   Begin VB.ListBox lstFav 
      Height          =   3180
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdAddFav 
      Caption         =   "Add"
      Height          =   315
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Add the current color to the favorites list"
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label lblHex 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   4260
      UseMnemonic     =   0   'False
      Width           =   2355
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuAs 
         Caption         =   "Save &As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuMerge 
         Caption         =   "&Merge..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear List"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide List"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuSend 
      Caption         =   "&Send"
      Begin VB.Menu mnuCopy 
         Caption         =   "To Clip&board..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "To Printer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuChart 
         Caption         =   "To HTML Chart..."
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuCurrent 
      Caption         =   "Current"
      Begin VB.Menu mnuCurFile 
         Caption         =   "(default)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTitle As String, strLastFav As String, strLastHTML As String, strLastHTMLDir As String, strLastDir As String

Private Sub cmdAddFav_Click()
Dim strDesc As String
    strDesc = InputBox("Please enter a short description" & _
        vbNewLine & "to help remember this color by:", "Add to Favorites", Hex$(frmColorRef.picColor.BackColor))
    If strDesc = "" Then Exit Sub
    lstFav.AddItem strDesc
    lstFav.ItemData(lstFav.NewIndex) = frmColorRef.picColor.BackColor
    lstFav.ListIndex = lstFav.NewIndex
    lstFav_Click
End Sub
Private Sub cmdChangeFav_Click()
Dim choice As Integer, lFav As Long
    If lstFav.ListIndex < 0 Then
        Beep
    Else
        choice = MsgBox("Do you want to change " & lstFav.List(lstFav.ListIndex) & vbNewLine & _
            "to represent the current color?", vbQuestion + vbYesNo, "Change Meaning of Favorite")
        If choice = vbYes Then lstFav.ItemData(lstFav.ListIndex) = frmColorRef.picColor.BackColor
    End If
    lFav = lstFav.ItemData(lstFav.ListIndex)
    lblHex.Caption = "H:""" & _
        MakeHex(RGBRed(lFav), RGBGreen(lFav), RGBBlue(lFav)) & """ V:&H" & Hex$(lFav)
End Sub


Private Sub cmdDelete_Click()
    If lstFav.ListIndex < 0 Then
        Beep
    Else
        lstFav.RemoveItem lstFav.ListIndex
        lblHex.Caption = ""
    End If
End Sub

Private Sub cmdRename_Click()
Dim strDesc As String
    If lstFav.ListIndex < 0 Then
        Beep
    Else
        strDesc = InputBox("Please enter a new name for" & _
            vbNewLine & lstFav.List(lstFav.ListIndex) & ":", "Rename Favorite", lstFav.List(lstFav.ListIndex))
        If strDesc = "" Then Exit Sub
        lstFav.List(lstFav.ListIndex) = strDesc
    End If
End Sub


Private Sub cmdTool_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuOpen_Click
        Case 1
            mnuSave_Click
        Case 2
            mnuAs_Click
        Case 3
            mnuMerge_Click
        Case 4
            mnuClear_Click
        Case 5
            mnuCopy_Click
        Case 6
            mnuPrint_Click
        Case 7
            mnuChart_Click
    End Select
End Sub

Private Sub cmdUse_Click()
    lstFav_DblClick
End Sub


Private Sub Form_Load()
    FileDialog.hwndOwner = frmColorRef.hwnd
    FileDialog.sInitDir = App.Path
    strTitle = GetINIString("ChartExport", "Title", "Color Chart")
    strLastHTML = GetINIString("ChartExport", "File", "ColorLab.html")
    strLastHTMLDir = GetINIString("ChartExport", "Dir", App.Path)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Visible = False
    frmColorRef.chkFav = False
    If blnEnd Then
        WriteINI "ChartExport", "Title", strTitle
        WriteINI "ChartExport", "File", strLastHTML
        WriteINI "ChartExport", "Dir", strLastHTMLDir
    Else
        Cancel = True
    End If
End Sub

Private Sub lstFav_Click()
Dim lFav As Long
    lFav = lstFav.ItemData(lstFav.ListIndex)
    lblHex.Caption = "H:""" & _
        MakeHex(RGBRed(lFav), RGBGreen(lFav), RGBBlue(lFav)) & """ V:&H" & Hex$(lFav)
End Sub

Private Sub lstFav_DblClick()
    frmColorRef.SetUndo
    frmColorRef.ChangeColor lstFav.ItemData(lstFav.ListIndex)
End Sub

Public Sub Pos()
    Me.top = frmColorRef.top
    If frmColorRef.left <= Me.Width + 4 * Screen.TwipsPerPixelX Then
        Me.left = frmColorRef.left + frmColorRef.Width + 4 * Screen.TwipsPerPixelX
    Else
        Me.left = frmColorRef.left - Me.Width - 4 * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub mnuAs_Click()
Dim sFile As SelectedFile
    If lstFav.ListCount = 0 Then Exit Sub
    FileDialog.sFilter = "ColorLab Favorites (*.clf)" & Chr$(0) & "*.clf" & Chr$(0) & Chr$(0)
    FileDialog.sDefFileExt = "*.clf"
    FileDialog.sDlgTitle = "Save ColorLab Favorites"
    FileDialog.flags = OFS_FILE_SAVE_FLAGS
    sFile = ShowSave(frmColorRef.hwnd, True, strLastFav, strLastDir)
    If sFile.bCanceled Then Exit Sub
    strLastFav = sFile.sFiles(1)
    strLastDir = sFile.sLastDirectory
    SaveFavs strLastDir & strLastFav
    mnuCurFile.Caption = strLastDir & strLastFav
End Sub

Private Sub mnuChart_Click()
Dim sFile As SelectedFile, strFavs As String, i As Integer, _
    strLine As String, lColor As Long, choice As Integer, _
    lReturn As Long
    
    If lstFav.ListCount = 0 Then Exit Sub
    frmExpo.Show vbModal, frmColorRef
    If Not blnCancel Then
        
        FileDialog.sFilter = "HTML Files (*.html)" & Chr$(0) & "*.html" & Chr$(0) & Chr$(0)
        FileDialog.sDefFileExt = "*.html"
        FileDialog.sDlgTitle = "Export ColorLab Favorites to HTML Chart"
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
        ReDim Preserve sFile.sFiles(1) As String
        sFile = ShowSave(frmColorRef.hwnd, True, strLastHTML, strLastHTMLDir)
        If sFile.bCanceled Then Exit Sub
        strLastHTML = sFile.sFiles(1)
        strLastHTMLDir = sFile.sLastDirectory
        strTitle = InputBox("Title for web page:", "HTML Export", strTitle)
        Open strLastHTMLDir & strLastHTML For Output As #1
        
        Print #1, "<HTML>"
        Print #1, "<HEAD>"
        Print #1, "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=iso-8859-1"">"
        Print #1, "<TITLE>" & strTitle & " - Generated by ColorLab </TITLE>"
        Print #1, "</HEAD>"
        Print #1, "<BODY>"
        Print #1, "<CENTER><H2>ColorLab Reference Chart</H2></CENTER><BR>"
        For i = 0 To lstFav.ListCount - 1
            strLine = ""
            lColor = lstFav.ItemData(i)
            If iExpo > 1 Then
                strLine = "&nbsp;&nbsp;&nbsp;&nbsp;HTML: """ & _
                    MakeHex(RGBRed(lColor), RGBGreen(lColor), RGBBlue(lColor)) & """"
            End If
            If iExpo Mod 2 = 1 Then
                strLine = strLine & "&nbsp;&nbsp;&nbsp;&nbsp;VB: &H" & Hex$(lColor)
            End If
            Print #1, "<B>" & lstFav.List(i) & "</B><TT>" & strLine & "</TT><BR>"
            Print #1, "<TABLE BORDER=""1"" CELLPADDING=""5"" BGCOLOR=""#" & _
                MakeHex(RGBRed(lColor), RGBGreen(lColor), RGBBlue(lColor)) & """>"
            Print #1, "<TR><TD ALIGN = ""CENTER"">"
            Print #1, "<FONT COLOR = ""#000000"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#FFFFFF"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#0000FF"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#FF00FF"">text <B>bold </B></FONT><BR>"
            Print #1, "<FONT COLOR = ""#808080"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#C0C0C0"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#000080"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#800080"">text <B>bold </B></FONT><BR>"
            Print #1, "<FONT COLOR = ""#FF0000"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#FFFF00"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#00FF00"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#00FFFF"">text <B>bold </B></FONT><BR>"
            Print #1, "<FONT COLOR = ""#800000"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#808000"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#008000"">text <B>bold </B></FONT>"
            Print #1, "<FONT COLOR = ""#008080"">text <B>bold </B></FONT></TD></TR>"
            Print #1, "</TABLE>"
        Next i
        Print #1, "</BODY>"
        Print #1, "</HTML>"
        Close #1
        choice = MsgBox("Color Chart Generation Complete." & vbNewLine & _
            "Would you like to view the chart in" & vbNewLine & _
            "your default web browser?", vbQuestion + vbYesNo, "ColorLab")
        If choice = vbYes Then
            lReturn = ShellExecute(0&, vbNullString, strLastHTMLDir & strLastHTML, _
                vbNullString, vbNullString, vbNormalFocus)
            If lReturn = -1 Then MsgBox "Unable to launch default browser", vbCritical + vbOKOnly, "Error"
        End If
    End If
End Sub

Private Sub mnuClear_Click()
Dim choice As Integer
    choice = MsgBox("Clear Favorites List?", vbQuestion + vbYesNo, "Clear")
    If choice = vbYes Then
        lstFav.Clear
        lblHex.Caption = ""
    End If
End Sub

Private Sub mnuCopy_Click()
Dim i As Integer, strLine As String, lColor As Long, strClip As String
    If lstFav.ListCount = 0 Then Exit Sub
    frmExpo.Show vbModal, frmColorRef
    If Not blnCancel Then
        strClip = ""
        strClip = strClip & "ColorLab Reference Chart" & vbNewLine
        For i = 0 To lstFav.ListCount - 1
            strLine = ""
            lColor = lstFav.ItemData(i)
            If iExpo > 1 Then
                strLine = vbTab & "HTML: """ & _
                    MakeHex(RGBRed(lColor), RGBGreen(lColor), RGBBlue(lColor)) & _
                    """"
            End If
            If iExpo Mod 2 = 1 Then
                strLine = strLine & vbTab & "VB: &H" & Hex$(lColor)
            End If
            strClip = strClip & lstFav.List(i) & vbNewLine & strLine & vbNewLine
        Next i
        Clipboard.SetText strClip
    End If
End Sub

Private Sub mnuCurFile_Click()
    MsgBox "This item is only for viewing the" & vbNewLine & _
        "name and location of the current" & vbNewLine & _
        "favorites file.  It doesn't do anything else!", _
        vbInformation + vbOKOnly, "Hello!"
End Sub

Private Sub mnuHelp_Click()
    MsgBox "+ Open: Opens a saved favorites list" & vbNewLine & _
        "+ Save: Save the current list to a file" & vbNewLine & _
        "+ Save As...: Save the current list under a different name" & vbNewLine & _
        "+ Merge: Combine current list with a saved list" & vbNewLine & _
        "+ Clear List: Delete all favorites from current list" & vbNewLine & _
        "+ Hide List: Close this window" & vbNewLine & _
        "+ Send to Clipboard: Puts a list of all colors and codes in the clipboard" & vbNewLine & _
        "+ Send to Printer: Prints a list of all colors and codes" & vbNewLine & _
        "+ Send to HTML Chart: Generates a web page with a color reference chart" & vbNewLine & _
        "+ Current: Shows the name of the currently loaded favorites file" & vbNewLine & _
        "+ Add Button: Adds Current Color to Favorites" & vbNewLine & _
        "+ Use Button: Make Selected Favorite Current Color" & vbNewLine & _
        "+ Delete Button: Remove Selected Favorite from list" & vbNewLine & _
        "+ Rename Button: Change description of selected favorite" & vbNewLine & _
        "+ Change Button: Use current color for selected favorite", _
        vbInformation + vbOKOnly, "How do I use favorites?"
End Sub

Private Sub mnuHide_Click()
    Me.Visible = False
    frmColorRef.chkFav = False
End Sub

Private Sub mnuMerge_Click()
Dim strFavs As String, strOldFavs As String, sFile As SelectedFile
    FileDialog.sDlgTitle = "Merge ColorLab Favorites"
    FileDialog.flags = OFS_FILE_OPEN_FLAGS
    FileDialog.sFilter = "ColorLab Favorites (*.clf)" & Chr$(0) & "*.clf" & Chr$(0) & Chr$(0)
    FileDialog.sDefFileExt = "*.clf"
    sFile = ShowOpen(frmColorRef.hwnd, True, "", strLastDir)
    If sFile.bCanceled Then Exit Sub
    strLastFav = sFile.sFiles(1)
    strLastDir = sFile.sLastDirectory
    strFavs = GetINISection("Favorites", strLastDir & strLastFav)
    LoadFavs strFavs & lstFav
    lblHex.Caption = ""
    mnuCurFile.Caption = strLastDir & strLastFav
End Sub

Private Sub mnuOpen_Click()
Dim sFile As SelectedFile, strFavs As String
    FileDialog.sDlgTitle = "Load ColorLab Favorites"
    FileDialog.flags = OFS_FILE_OPEN_FLAGS
    FileDialog.sFilter = "ColorLab Favorites (*.clf)" & Chr$(0) & "*.clf" & Chr$(0) & Chr$(0)
    FileDialog.sDefFileExt = "*.clf"
    sFile = ShowOpen(frmColorRef.hwnd, True, strLastFav, strLastDir)
    If sFile.bCanceled Then Exit Sub
    strLastFav = sFile.sFiles(1)
    strLastDir = sFile.sLastDirectory
    lstFav.Clear
    LoadFavs strLastDir & strLastFav
    mnuCurFile.Caption = strLastDir & strLastFav
End Sub

Private Sub mnuPrint_Click()
Dim i As Integer, strLine As String, lColor As Long
    If lstFav.ListCount = 0 Then Exit Sub
    frmExpo.Show vbModal, frmColorRef
    If Not blnCancel Then
        Printer.FontSize = 10
        Printer.Print "ColorLab Reference Chart" & vbNewLine
        For i = 0 To lstFav.ListCount - 1
            strLine = ""
            lColor = lstFav.ItemData(i)
            If iExpo > 1 Then
                strLine = vbTab & "HTML: """ & _
                    MakeHex(RGBRed(lColor), RGBGreen(lColor), RGBBlue(lColor)) & _
                    """"
            End If
            If iExpo Mod 2 = 1 Then
                strLine = strLine & vbTab & "VB: &H" & Hex$(lColor)
            End If
            Printer.FontBold = True
            Printer.Print lstFav.List(i)
            Printer.FontBold = False
            Printer.Print strLine
        Next i
        Printer.EndDoc
    End If
End Sub

Private Sub mnuSave_Click()
Dim strFavs As String
    If strLastFav = "" Then
        mnuAs_Click
    Else
        SaveFavs strLastDir & strLastFav
    End If
End Sub

Public Sub LoadFavs(Optional strPath As String)
Dim strItem As String
    If strPath = "" Then strPath = getFullPath(App.Path, "current.clf")
On Error GoTo NoFile
    Open strPath For Input As #1
On Error GoTo 0
    Do While Not EOF(1)
        Line Input #1, strItem
        lstFav.AddItem Mid$(strItem, 8)
        lstFav.ItemData(lstFav.NewIndex) = CLng("&H00" & left$(strItem, 6))
    Loop
    Close #1
    lblHex.Caption = ""
NoFile:
End Sub

Public Sub SaveFavs(Optional strPath As String)
Dim strItem As String, lList As Long
    If strPath = "" Then strPath = getFullPath(App.Path, "current.clf")

    Open strPath For Output As #1
    
    For lList = 0 To lstFav.ListCount - 1
        strItem = ZHex(lstFav.ItemData(lList), 6) & "=" & lstFav.List(lList)
        Print #1, strItem
    Next lList
    Close #1
End Sub
