VERSION 5.00
Begin VB.Form frmColorRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "{ C o l o r L a b }"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4845
   Icon            =   "frmColorRef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   321
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4260
      TabIndex        =   110
      Top             =   4380
      Width           =   555
   End
   Begin VB.CommandButton cmdMainHelp 
      Caption         =   "Help!"
      Height          =   375
      Left            =   4260
      TabIndex        =   109
      Top             =   4020
      Width           =   555
   End
   Begin VB.CheckBox chkBig 
      Caption         =   "B&ig Picture"
      Height          =   375
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Show/Hide the Big Picture Window"
      Top             =   4380
      Width           =   975
   End
   Begin VB.CheckBox chkFav 
      Caption         =   "&Favorites"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Show/Hide the Favorite Colors list"
      Top             =   4380
      Width           =   915
   End
   Begin VB.CommandButton cmdVBClip 
      Height          =   375
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Copy VB Hex code to clipboard"
      Top             =   2580
      Width           =   375
   End
   Begin VB.CommandButton cmdHTMLClip 
      Height          =   375
      Left            =   3810
      Picture         =   "frmColorRef.frx":548A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copy HTML Hex code to clipboard"
      Top             =   2100
      Width           =   375
   End
   Begin VB.PictureBox picAnchor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   1
      Left            =   4245
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   86
      TabStop         =   0   'False
      ToolTipText     =   "Click to make current color an anchor color for the blend bar"
      Top             =   3405
      Width           =   525
   End
   Begin VB.PictureBox picAnchor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   0
      Left            =   4245
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   85
      TabStop         =   0   'False
      ToolTipText     =   "Click to make current color an anchor color for the blend bar"
      Top             =   105
      Width           =   525
   End
   Begin VB.VScrollBar vL 
      Height          =   375
      LargeChange     =   4
      Left            =   3720
      Max             =   240
      SmallChange     =   2
      TabIndex        =   26
      Top             =   3960
      Width           =   195
   End
   Begin VB.VScrollBar vS 
      Height          =   375
      LargeChange     =   4
      Left            =   2580
      Max             =   240
      TabIndex        =   25
      Top             =   3960
      Width           =   195
   End
   Begin VB.VScrollBar vH 
      Height          =   375
      LargeChange     =   4
      Left            =   1440
      Max             =   239
      TabIndex        =   24
      Top             =   3960
      Width           =   195
   End
   Begin VB.VScrollBar vB 
      Height          =   375
      LargeChange     =   4
      Left            =   3720
      Max             =   255
      TabIndex        =   23
      Top             =   3300
      Width           =   195
   End
   Begin VB.VScrollBar vG 
      Height          =   375
      LargeChange     =   4
      Left            =   2580
      Max             =   255
      TabIndex        =   22
      Top             =   3300
      Width           =   195
   End
   Begin VB.VScrollBar vR 
      Height          =   375
      LargeChange     =   4
      Left            =   1440
      Max             =   255
      TabIndex        =   21
      Top             =   3300
      Width           =   195
   End
   Begin VB.Timer tmr5by5 
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton cmd5by5 
      Caption         =   "&5 × 5"
      Height          =   555
      Left            =   120
      Picture         =   "frmColorRef.frx":55D4
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Take a 5x5 average sample from the screen"
      Top             =   1380
      Width           =   675
   End
   Begin VB.PictureBox pic5x5 
      Height          =   1485
      Left            =   2580
      ScaleHeight     =   1425
      ScaleWidth      =   1425
      TabIndex        =   50
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   75
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   74
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   73
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   72
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   71
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   70
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   69
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   68
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   67
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   66
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   65
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   64
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   63
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   62
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   61
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   60
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   59
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   58
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   57
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   56
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   55
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   54
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   53
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   52
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picPoint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   51
         ToolTipText     =   "Enlargement of the 5x5 area under the pointer.  If you are reading this, you are getting color feedback!"
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.TextBox txtH 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   900
      TabIndex        =   14
      ToolTipText     =   "Hue (Tint) value for above color"
      Top             =   3960
      Width           =   555
   End
   Begin VB.TextBox txtL 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3180
      TabIndex        =   18
      ToolTipText     =   "Luminence (Brightness) value for above color"
      Top             =   3960
      Width           =   555
   End
   Begin VB.TextBox txtS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      ToolTipText     =   "SAturation (Richness) value for above color"
      Top             =   3960
      Width           =   555
   End
   Begin VB.Timer tmrPick 
      Left            =   1440
      Top             =   0
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "Sa&mple"
      Height          =   555
      Left            =   120
      Picture         =   "frmColorRef.frx":571E
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Pick the Background Color from Screen"
      Top             =   840
      Width           =   675
   End
   Begin VB.TextBox txtVB 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   5
      ToolTipText     =   "VB Hex Code - 0 to FFFFFF"
      Top             =   2580
      Width           =   1275
   End
   Begin VB.TextBox txtHTML 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   3
      ToolTipText     =   "HTML Hex code - 000000 to FFFFFF"
      Top             =   2040
      Width           =   1275
   End
   Begin VB.TextBox txtG 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      ToolTipText     =   "Green value for above color"
      Top             =   3300
      Width           =   555
   End
   Begin VB.TextBox txtB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3180
      TabIndex        =   12
      ToolTipText     =   "Blue value for above color"
      Top             =   3300
      Width           =   555
   End
   Begin VB.TextBox txtR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   900
      TabIndex        =   8
      ToolTipText     =   "Red value for above color"
      Top             =   3300
      Width           =   555
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Pick"
      Height          =   795
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Pick the Background Color"
      Top             =   60
      Width           =   675
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   60
      ScaleHeight     =   1635
      ScaleWidth      =   4035
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Current color"
      Top             =   300
      Width           =   4095
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   7
         Left            =   3420
         TabIndex        =   83
         ToolTipText     =   "Click to adjust brightness"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   6
         Left            =   3000
         TabIndex        =   82
         ToolTipText     =   "Click to adjust brightness"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   5
         Left            =   2580
         TabIndex        =   81
         ToolTipText     =   "Click to adjust brightness"
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   4
         Left            =   3420
         TabIndex        =   80
         ToolTipText     =   "Click to adjust brightness"
         Top             =   615
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   3
         Left            =   2580
         TabIndex        =   79
         ToolTipText     =   "Click to adjust brightness"
         Top             =   615
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   2
         Left            =   3420
         TabIndex        =   78
         ToolTipText     =   "Click to adjust brightness"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   1
         Left            =   3000
         TabIndex        =   77
         ToolTipText     =   "Click to adjust brightness"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "-20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   2580
         TabIndex        =   76
         ToolTipText     =   "Click to adjust brightness"
         Top             =   180
         Width           =   435
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gray"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Index           =   15
         Left            =   1620
         TabIndex        =   45
         ToolTipText     =   "Standard text colors"
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Silver"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   14
         Left            =   840
         TabIndex        =   44
         ToolTipText     =   "Standard text colors"
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Navy"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   13
         Left            =   1620
         TabIndex        =   43
         ToolTipText     =   "Standard text colors"
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fuchsia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   225
         Index           =   12
         Left            =   840
         TabIndex        =   42
         ToolTipText     =   "Standard text colors"
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   225
         Index           =   11
         Left            =   1620
         TabIndex        =   41
         ToolTipText     =   "Standard text colors"
         Top             =   960
         Width           =   345
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aqua"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Index           =   10
         Left            =   840
         TabIndex        =   40
         ToolTipText     =   "Standard text colors"
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Olive"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   225
         Index           =   9
         Left            =   1620
         TabIndex        =   39
         ToolTipText     =   "Standard text colors"
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yellow"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   840
         TabIndex        =   38
         ToolTipText     =   "Standard text colors"
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   7
         Left            =   1620
         TabIndex        =   37
         ToolTipText     =   "Standard text colors"
         Top             =   780
         Width           =   465
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lime"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   6
         Left            =   840
         TabIndex        =   36
         ToolTipText     =   "Standard text colors"
         Top             =   780
         Width           =   375
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maroon"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   5
         Left            =   1620
         TabIndex        =   35
         ToolTipText     =   "Standard text colors"
         Top             =   600
         Width           =   645
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   4
         Left            =   840
         TabIndex        =   34
         ToolTipText     =   "Standard text colors"
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Purple"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Index           =   3
         Left            =   1620
         TabIndex        =   33
         ToolTipText     =   "Standard text colors"
         Top             =   420
         Width           =   645
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   32
         ToolTipText     =   "Standard text colors"
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   31
         ToolTipText     =   "Standard text colors"
         Top             =   60
         Width           =   495
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Black"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   1620
         TabIndex        =   30
         ToolTipText     =   "Standard text colors"
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.Label lblHTML2 
      AutoSize        =   -1  'True
      Caption         =   """"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   108
      Top             =   2070
      Width           =   195
   End
   Begin VB.Label lblUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   107
      ToolTipText     =   "Click to revert to this Color"
      Top             =   4440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1740
      TabIndex        =   106
      ToolTipText     =   "Click to revert to this Color"
      Top             =   4440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   105
      ToolTipText     =   "Click to revert to this Color"
      Top             =   4440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   104
      ToolTipText     =   "Click to revert to this Color"
      Top             =   4440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   103
      ToolTipText     =   "Click to revert to this Color"
      Top             =   4440
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label lblPast 
      Caption         =   "Recent:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   102
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   4260
      TabIndex        =   101
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3180
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   4260
      TabIndex        =   100
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   4260
      TabIndex        =   99
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2820
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   4260
      TabIndex        =   98
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   4260
      TabIndex        =   97
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   4260
      TabIndex        =   96
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   4260
      TabIndex        =   95
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   4260
      TabIndex        =   94
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   4260
      TabIndex        =   93
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   4260
      TabIndex        =   92
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4260
      TabIndex        =   91
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1380
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4260
      TabIndex        =   90
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4260
      TabIndex        =   89
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4260
      TabIndex        =   88
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblBlend 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4260
      TabIndex        =   87
      ToolTipText     =   "BlendBar: Click to make blended color the current color"
      Top             =   660
      Width           =   495
   End
   Begin VB.Label lblInstruct 
      Alignment       =   2  'Center
      Caption         =   "Hold mouse over item for description"
      Height          =   255
      Left            =   840
      TabIndex        =   84
      Top             =   60
      Width           =   2835
   End
   Begin VB.Image imgHappy 
      Height          =   240
      Left            =   3840
      Picture         =   "frmColorRef.frx":5868
      ToolTipText     =   "Happy, happy little program!"
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblH 
      Caption         =   "&H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   660
      TabIndex        =   13
      Top             =   4020
      Width           =   210
   End
   Begin VB.Label lblL 
      Caption         =   "&L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2940
      TabIndex        =   17
      Top             =   4020
      Width           =   195
   End
   Begin VB.Label lblS 
      Caption         =   "&S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   15
      Top             =   4020
      Width           =   225
   End
   Begin VB.Label lblHSL 
      AutoSize        =   -1  'True
      Caption         =   "Hue / Saturation / Luminence values:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   49
      Top             =   3720
      Width           =   3120
   End
   Begin VB.Label lblRGB 
      AutoSize        =   -1  'True
      Caption         =   "Red / Green / Blue values for the above color:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   48
      Top             =   3060
      Width           =   3840
   End
   Begin VB.Label lblVB 
      Caption         =   "&&H"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   47
      Top             =   2610
      Width           =   1755
   End
   Begin VB.Label lblHTML 
      AutoSize        =   -1  'True
      Caption         =   """# "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   46
      Top             =   2070
      Width           =   585
   End
   Begin VB.Label lblG 
      Caption         =   "&G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   9
      Top             =   3360
      Width           =   225
   End
   Begin VB.Label lblB 
      Caption         =   "&B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2940
      TabIndex        =   11
      Top             =   3360
      Width           =   195
   End
   Begin VB.Label lblR 
      Caption         =   "&R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   660
      TabIndex        =   7
      Top             =   3360
      Width           =   210
   End
   Begin VB.Label lblV 
      Alignment       =   1  'Right Justify
      Caption         =   "The VB hex code for the above color is:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   29
      Top             =   2520
      Width           =   1830
   End
   Begin VB.Label lblHT 
      Alignment       =   1  'Right Justify
      Caption         =   "The HTML hex code for the above color is:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   28
      Top             =   1980
      Width           =   1875
   End
End
Attribute VB_Name = "frmColorRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sColor As SelectedColor
Dim iHpos As Integer, iVpos As Integer
Dim blnUpdate As Boolean, blnHue As Boolean, blnSat As Boolean, blnLum As Boolean, blnBuddy As Boolean, blnAdjust As Boolean

'APIs for color-sampling routines

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RGBType
    r As Byte
    G As Byte
    B As Byte
    Filler As Byte
End Type

Private Type RGBLongType
    clr As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub chkBig_Click()
    frmBig.Pos
    frmBig.Visible = chkBig
End Sub

Private Sub chkFav_Click()
    frmFav.Pos
    frmFav.Visible = chkFav
End Sub

Private Sub cmd5by5_Click()
'Start a 5x5 average sample
Dim lReturn As Long
    SetUndo
    lReturn = SetCapture(picColor.hwnd)
    tmr5by5.Interval = 100
    pic5x5.Visible = True
End Sub



Private Sub cmdChange_Click()
'Call common color dialog (no .dll)
'This routine by Paul Mather, with minor modifications
    On Error GoTo e_Trap
    sColor.oSelectedColor = picColor.BackColor
    sColor = ShowColor(Me.hwnd, True, sColor.oSelectedColor)
    If Not sColor.bCanceled Then
        SetUndo
        ChangeColor sColor.oSelectedColor
    End If
    Exit Sub
e_Trap:
    Exit Sub
End Sub


Private Sub cmdHTMLClip_Click()
    Clipboard.SetText (txtHTML.Text)
End Sub

Private Sub cmdMainHelp_Click()

    MsgBox "+ Current Color Window: Shows any color as background to" & vbNewLine & _
    vbTab & "the 16 standard text colors" & vbNewLine & _
    "+ -20...+20 Boxes: Darken or Brighten the color by percentage" & vbNewLine & _
    "+ Pick Button: Pick color from standard color dialog" & vbNewLine & _
    "+ Sample Button: Pick color from a point on-screen" & vbNewLine & _
    "+ 5 x 5 Button: Pick a color by averaging a square of 5x5" & vbNewLine & _
    vbTab & "points on screen" & vbNewLine & _
    "+ BlendBar (at right): Click box on either end to set current" & vbNewLine & _
    vbTab & "color as 'anchor color'.  Click anywhere in the 15-shade" & vbNewLine & _
    vbTab & "blend area to select that color as current color." & vbNewLine & _
    vbTab & "Right-click either end to reselect that color" & vbNewLine & _
    vbTab & "as current color." & vbNewLine & _
    "+ Color Clipboard Buttons: Copy that hex code to clipboard" & vbNewLine & _
    "+ RGB/HSL Values: Tweak colors numerically" & vbNewLine & _
    "+ Recent Colors: Click to bring back one of the last five colors" & vbNewLine & _
    "+ Favorites Button: Load & Save your best colors, create HTML" & vbNewLine & _
    vbTab & "color charts, and more!" & vbNewLine & _
    "+ Big Picture Button: See the color in a large window and" & vbNewLine & _
    vbTab & "test your web graphics against it" & vbNewLine & _
    "+ Quit Button: Buh-Bye!", vbInformation + vbOKOnly, "So much stuff in such a little space!"
End Sub

Private Sub cmdPick_Click()
'start a single point sampling
Dim lReturn As Long
    SetUndo
    lReturn = SetCapture(picColor.hwnd)
    tmrPick.Interval = 50
End Sub


Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdVBClip_Click()
    Clipboard.SetText (txtVB.Text)
End Sub

Private Sub Form_Load()
Dim strFavs As String
'Startup
    cmdChange.Picture = Me.Icon
    cmdVBClip.Picture = cmdHTMLClip.Picture
    'get saved custom colors
    cc = GetINIString("Settings", "CustomColors", "")
    iExpo = GetININumber("Settings", "Export", 1)
    HtoU cc
    'get saved blendbar settings
    picAnchor(0).BackColor = GetININumber("BlendBar", "Top", 65535)
    picAnchor(1).BackColor = GetININumber("BlendBar", "Bottom", 255)
    'get favorites
    strFavs = GetINISection("Favorites")
    Load frmFav
    frmFav.LoadFavs
    Load frmBig
    'setup screen w/ default gray
    ChangeColor GetININumber("Settings", "Current", 12632256)
    BlendEm
    SetHook Me.hwnd, True
End Sub

Private Sub updateCodes()
Dim r As Long, G As Long, B As Long, HSLV As HSLCol, lColor As Long, i As Integer
    
    blnUpdate = True 'keeps the routine from being called again everytime it sets a value
    lColor = picColor.BackColor
    'the VB one is easy! ;)
    txtVB.Text = Hex$(lColor)
        
    'calculate & set the individual RGB values & scrolls
    
    r = RGBRed(lColor)
    G = RGBGreen(lColor)
    B = RGBBlue(lColor)
    txtR.Text = r
    vR = 255 - r
    txtR.BackColor = RGB(Val(txtR.Text), 0, 0)
    txtG.Text = G
    vG = 255 - G
    txtG.BackColor = RGB(0, Val(txtG.Text), 0)
    If Val(txtG.Text) > 172 Then
        txtG.ForeColor = vbBlack
    Else
        txtG.ForeColor = vbWhite
    End If
    txtB.Text = B
    VB = 255 - B
    txtB.BackColor = RGB(0, 0, Val(txtB.Text))
    
    'put together the HTML code
    
    txtHTML.Text = MakeHex(r, G, B)
    
    'Calculate & set HSL boxes & scrolls
    HSLV = RGBtoHSL(lColor)
    
    If Not blnHue Then
        txtH.Text = HSLV.Hue
        vH = 239 - IIf(HSLV.Hue > 239, 0, HSLV.Hue)
    End If
    
    If Not blnSat Then
        txtS.Text = HSLV.Sat
        vS = 240 - HSLV.Sat
    End If
    
    If Not blnLum Then
        txtL.Text = HSLV.Lum
        vL = 240 - HSLV.Lum
    End If
    
    'set the adjust brightness boxes if they're visible (not hidden by
    'pic5x5, which is acting as a container only).  If they're not visible,
    'an average sample is going on, and why add extra processing to an
    'already complicated task?
    
    If Not pic5x5.Visible Then
        For i = 0 To 7
            'the first function produces the values .2, .15, .1, .05
            'for the darken routine, the second function reverses the
            'order for brighten
            Select Case i
                Case 0 To 3
                    lblAdj(i).BackColor = Darken(lColor, (0.05 * (4 - i)))
                Case 4 To 7
                    lblAdj(i).BackColor = Brighten(lColor, (0.05 * (i - 3)))
            End Select
            lblAdj(i).ForeColor = ContrastingColor(lblAdj(i).BackColor)
        Next i
    End If
    blnUpdate = False 'now changes to the text boxes with trigger this routine
End Sub

Private Sub Form_Paint()
Dim RC As RECT, lReturn As Long, i As Integer
        RC.top = lblBlend(0).top - 2
        RC.left = lblBlend(0).left - 2
        RC.Bottom = lblBlend(14).top + lblBlend(14).Height + 2
        RC.Right = lblBlend(14).left + lblBlend(14).Width + 2
        lReturn = DrawEdge(hdc, RC, EDGE_BUMP, BF_RECT) ' Or BF_SOFT)
    For i = 0 To 1
        lReturn = GetClientRect(picAnchor(i).hwnd, RC)
        lReturn = InflateRect(RC, 2, 2)
        RC.top = RC.top + picAnchor(i).top
        RC.left = RC.left + picAnchor(i).left
        RC.Bottom = RC.Bottom + picAnchor(i).top
        RC.Right = RC.Right + picAnchor(i).left
        lReturn = DrawEdge(hdc, RC, EDGE_BUMP, BF_RECT) ' Or BF_SOFT)
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetHook Me.hwnd, False
End Sub

Private Sub Form_Resize()
    If chkFav Then frmFav.Visible = (Me.WindowState = vbNormal)
    If chkBig Then frmBig.Visible = (Me.WindowState = vbNormal)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strFavs As String, lReturn As Long
'Save settings
    WriteINI "Settings", "CustomColors", Chr$(34) & cc & Chr$(34) 'save custom colors
    WriteINI "Settings", "Export", CStr(iExpo)
    WriteINI "Settings", "Current", CStr(picColor.BackColor)
    
    WriteINI "BlendBar", "Top", CStr(CLng(picAnchor(0).BackColor))
    WriteINI "BlendBar", "Bottom", CStr(CLng(picAnchor(1).BackColor))
    frmFav.SaveFavs
    blnEnd = True
    Unload frmFav
    Unload frmBig
    On Error Resume Next
    lReturn = ReleaseCapture()
    
End Sub

Private Sub imgHappy_Click()
' "About"
    MsgBox "ColorLab  v 3.2" & vbNewLine & _
        "{ f r e e w a r e }" & vbNewLine & _
        "©1999 B/W Software" & vbNewLine & _
        "Dan Redding" & vbNewLine & vbNewLine & _
        "Screen color pick routines and 'Form_Move' event" & vbNewLine & _
        "adapted from sample code by Matt Hart" & vbNewLine & _
        "[ http://www.matthart.com ]" & vbNewLine & _
        "Common Dialog routines based on routines by Paul Mather" & vbNewLine & _
        "[available at http://www.planet-source-code.com/vb]" & vbNewLine & vbNewLine & _
        "For René, who had way too many links to HTML color tables...", _
        vbInformation + vbOKOnly, "About ColorLab"
End Sub

Private Sub lblAdj_Click(Index As Integer)
'Brightness adjustment boxes
    SetUndo
    ChangeColor lblAdj(Index).BackColor
End Sub

Private Sub lblBlend_Click(Index As Integer)
    SetUndo
    ChangeColor lblBlend(Index).BackColor
End Sub

Private Sub lblUndo_Click(Index As Integer)
Dim lColor As Long
    lColor = lblUndo(Index).BackColor
    SetUndo
    ChangeColor lColor
End Sub


Private Sub picAnchor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        picAnchor(Index).BackColor = picColor.BackColor
        BlendEm
    ElseIf Button = 2 Then
        SetUndo
        ChangeColor picAnchor(Index).BackColor
    End If
End Sub

Private Sub picColor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lReturn As Long
    'check that we are actually sampling
    If tmrPick.Interval > 0 Or tmr5by5.Interval > 0 Then
        tmrPick.Interval = 0
        tmr5by5.Interval = 0
        pic5x5.Visible = False
        updateCodes
        lReturn = ReleaseCapture
    End If
End Sub

Private Sub tmr5by5_Timer()
'adapted from the routine in tmrPick, this samples 25 different
'pixels centering on the cursor, setting 25 small picture boxes
'to produce the 'enlarged' view.  The routine then runs through
'all 25 calculates the average color by averaging the seperate
'red, green, and blue values to set the main color window.

Static lX As Long, lY As Long
On Local Error Resume Next
Dim P As POINTAPI, H As Long, hD As Long, r As Long
Dim i As Integer, Red As Long, Blue As Long, Green As Long
Dim X1 As Long, Y1 As Long
Static ScrX As Long, ScrY As Long
    If ScrX = 0 Then
        ScrX = Screen.Width / Screen.TwipsPerPixelX
        ScrY = Screen.Height / Screen.TwipsPerPixelY
    End If
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    For i = 0 To 24
        '5x5 position relative to cursor (x & y = -2 to 2)
        X1 = (lX + (i Mod 5) - 2)
        Y1 = (lY + (i \ 5) - 2)
        P.x = X1
        P.y = Y1
        
        If X1 < 0 Or Y1 < 0 Or X1 > ScrX Or Y1 > ScrY Then
            r = 0
        Else
            'this information needs to be recalcualted
            'for each point; after all, the 5x5 square
            'could overlap 2 or more windows
            
            'which window?
            H = WindowFromPoint(X1, Y1)
            
            'get device context for that window
            hD = GetDC(H)
            
            'convert screen coordinates to local window
            ScreenToClient H, P
            
            'get color
            r = GetPixel(hD, P.x, P.y)
            If r = -1 Then
                'titlebar or other special area
                'get color by copying the point to picturebox, then checking that
                BitBlt picPoint(i).hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
                r = picPoint(i).Point(0, 0)
            Else
                'R is the color
                picPoint(i).PSet (0, 0), r
            End If
            'Must do to prevent memory leaks
            ReleaseDC H, hD
        End If
        'set backcolor of whole picturebox to R
        picPoint(i).BackColor = r
    Next i
    
    'averaging
    
    For i = 0 To 24
        Red = Red + RGBRed(picPoint(i).BackColor)
        Blue = Blue + RGBBlue(picPoint(i).BackColor)
        Green = Green + RGBGreen(picPoint(i).BackColor)
    Next i
    
    'set main picturebox w/ average color
    ChangeColor RGB(CInt(Red / 25), CInt(Green / 25), CInt(Blue / 25))

End Sub

Private Sub tmrPick_Timer()
'This routine adapted from a project by Matt Hart

'Matt's comments follow:
' Getpixel sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This sample shows how to get the pixel color of any point
' on the screen. The GetPixel API requires CLIENT coordinates,
' so you must first get the window handle and hDC where the
' cursor is. Once you get that, you can get the pixel.
'
' However, there's one "gotcha" I found while writing this.
' Window titlebars return a "-1" for the pixel color, which
' is invalid! So, what I did to get around that was use
' BitBlt to copy a pixel from that device to the PictureBox
' control I'm using to show the colors, then use the Point
' method to check the color.

'for detailed comments, see corresponding function in tmr5x5
Static lX As Long, lY As Long
On Local Error Resume Next
Dim P As POINTAPI, H As Long, hD As Long, r As Long
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    H = WindowFromPoint(lX, lY)
    hD = GetDC(H)
    ScreenToClient H, P
    r = GetPixel(hD, P.x, P.y)
    If r = -1 Then
        BitBlt picColor.hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
        r = picColor.Point(0, 0)
    Else
        picColor.PSet (0, 0), r
    End If
    ReleaseDC H, hD
    ChangeColor r
End Sub

Private Sub txtB_Change()
    'change blue value
    If blnUpdate Then Exit Sub 'updating, don't need to adjust anything else here
    If Not blnAdjust Then
        SetUndo
        blnAdjust = True
    End If
    UpdateBlue
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UpdateBlue
        blnAdjust = False
        txtB.SelStart = 0
        txtB.SelLength = Len(txtB.Text)
        KeyAscii = 0
    End If
End Sub

Private Sub txtB_LostFocus()
    UpdateBlue
    blnAdjust = False
End Sub

Private Sub UpdateBlue()
    'too high?
    If Val(txtB.Text) > 255 Then
        txtB.Text = "255"
    Else
        txtB.Text = Val(txtB.Text)
    End If
    'set new color
    ChangeColor RGB(Val(txtR.Text), Val(txtG.Text), Val(txtB.Text))
    'blnbuddy if txtB was changed by vB scroller.
    'without this, the two routines would trigger each other until overflow
    If Not blnBuddy Then
        blnBuddy = True
        VB.Value = 255 - Val(txtB.Text)
        blnBuddy = False
    End If
    'select if one of these two values (easy to overtype)
    If txtB.Text = "0" Or txtB.Text = "255" Then txtB_GotFocus
End Sub
Private Sub txtB_GotFocus()
'select all when get focus
    txtB.SelStart = 0
    txtB.SelLength = Len(txtB.Text)
End Sub

'For txtG routine comments, see txtB
Private Sub txtG_Change()
    If blnUpdate Then Exit Sub
    If Not blnAdjust Then
        SetUndo
        blnAdjust = True
    End If
    UpdateGreen
End Sub

Private Sub UpdateGreen()
    If Val(txtR.Text) > 255 Then
        txtG.Text = "255"
    Else
        txtG.Text = Val(txtG.Text)
    End If
    ChangeColor RGB(Val(txtR.Text), Val(txtG.Text), Val(txtB.Text))
    If Not blnBuddy Then
        blnBuddy = True
        vR.Value = 255 - Val(txtG.Text)
        blnBuddy = False
    End If
    If txtG.Text = "0" Or txtG.Text = "255" Then txtG_GotFocus
End Sub
Private Sub txtG_GotFocus()
    txtG.SelStart = 0
    txtG.SelLength = Len(txtG.Text)
End Sub
Private Sub txtG_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UpdateGreen
        txtG.SelStart = 0
        txtG.SelLength = Len(txtG.Text)
        blnAdjust = False
        KeyAscii = 0
    End If
End Sub

Private Sub txtG_LostFocus()
    UpdateGreen
    blnAdjust = False
End Sub

Private Sub txtH_Change()
    If blnUpdate Then Exit Sub 'updating, don''t need to change here
    If Not blnAdjust Then
        SetUndo
        blnAdjust = True
    End If
    UpdateHue
End Sub

Private Sub UpdateHue()
Dim HSLV As HSLCol
    'too high?
    If Val(txtH.Text) >= HSLMAX Then
        txtH.Text = HSLMAX - 1
    Else
        txtH.Text = Val(txtH.Text)
    End If
    'calc & set new rgb color
    HSLV.Hue = Val(txtH.Text)
    HSLV.Sat = Val(txtS.Text)
    HSLV.Lum = Val(txtL.Text)
    'protect from another loop (HSL->RGB->HSL sometimes changes HSL due to rounding errors)
    blnHue = True
    ChangeColor HSLtoRGB(HSLV)
    blnHue = False
    'Protect from infinite loop adjusting vH scroller
    If Not blnBuddy Then
        blnBuddy = True
        vH.Value = 239 - Val(txtH.Text)
        blnBuddy = False
    End If
    'select for overtyping if high or low val
    If txtH.Text = "0" Or txtH.Text = "239" Then txtH_GotFocus

End Sub
Private Sub txtH_GotFocus()
    txtH.SelStart = 0
    txtH.SelLength = Len(txtH.Text)
End Sub
Private Sub txtH_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UpdateHue
        txtH.SelStart = 0
        txtH.SelLength = Len(txtH.Text)
        blnAdjust = False
        KeyAscii = 0
    End If
End Sub

Private Sub txtH_LostFocus()
    UpdateHue
    blnAdjust = False
End Sub

Private Sub txtHTML_Change()
Dim r As Long, G As Long, B As Long
    txtHTML.Text = UCase$(txtHTML.Text) 'uppercase it
    txtHTML.SelStart = iHpos 'keep cursor where it was after uppercase
    If Len(txtHTML.Text) = 6 Then 'full code; change color
        If Not isHex(txtHTML.Text) Then
            Beep 'not valid!
            Exit Sub
        End If
        'get RGB values from hex string, the easy way
        r = Val("&H" & Mid$(txtHTML.Text, 1, 2))
        G = Val("&H" & Mid$(txtHTML.Text, 3, 2))
        B = Val("&H" & Right$(txtHTML.Text, 2))
        
        'set color and update codes
        '(unless txtHTML was changed BY the updateCodes rotuine)
        If Not blnUpdate Then SetUndo
        picColor.BackColor = RGB(r, G, B)
        If Not blnUpdate Then
            updateCodes
        End If
    End If
End Sub

Private Sub txtHTML_GotFocus()
    txtHTML.SelStart = 0
    txtHTML.SelLength = Len(txtHTML.Text)
    iHpos = 0 'save cursor position
End Sub

Private Sub txtHTML_KeyDown(KeyCode As Integer, Shift As Integer)
'save new cursor position
    If KeyCode = vbKeyBack Then
        iHpos = txtHTML.SelStart - 1
    Else
        iHpos = txtHTML.SelStart + 1
    End If
End Sub

'for txtL comments, see corresponding in txtH
Private Sub txtL_Change()
    If blnUpdate Then Exit Sub
    If Not blnAdjust Then
        SetUndo
        blnAdjust = True
    End If
    UpdateLumin
End Sub

Private Sub UpdateLumin()
Dim HSLV As HSLCol
    
    If Val(txtL.Text) > HSLMAX Then
        txtL.Text = HSLMAX
    Else
        txtL.Text = Val(txtL.Text)
    End If
    HSLV.Hue = Val(txtH.Text)
    HSLV.Sat = Val(txtS.Text)
    HSLV.Lum = Val(txtL.Text)
    blnHue = True
    ChangeColor HSLtoRGB(HSLV)
    blnHue = False
    If Not blnBuddy Then
        blnBuddy = True
        vL.Value = 240 - Val(txtL.Text)
        blnBuddy = False
    End If
    If txtL.Text = "0" Or txtL.Text = "240" Then txtL_GotFocus

End Sub
Private Sub txtL_GotFocus()
    txtL.SelStart = 0
    txtL.SelLength = Len(txtL.Text)
End Sub

Private Sub txtL_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UpdateLumin
        txtL.SelStart = 0
        txtL.SelLength = Len(txtL.Text)
        blnAdjust = False
        KeyAscii = 0
    End If
End Sub

Private Sub txtL_LostFocus()
    UpdateLumin
    blnAdjust = False
End Sub

'for txtR comments, see corresponding in txtB
Private Sub txtR_Change()
    If blnUpdate Then Exit Sub
    If Not blnAdjust Then
        SetUndo
        blnAdjust = True
    End If
    UpdateRed
End Sub

Private Sub UpdateRed()
    If Val(txtR.Text) > 255 Then
        txtR.Text = "255"
    Else
        txtR.Text = Val(txtR.Text)
    End If
    ChangeColor RGB(Val(txtR.Text), Val(txtG.Text), Val(txtB.Text))
    blnBuddy = True
    vR.Value = 255 - Val(txtR.Text)
    blnBuddy = False
    If Not blnBuddy Then
        blnBuddy = True
        vR.Value = 255 - Val(txtR.Text)
        blnBuddy = False
    End If
    If txtR.Text = "0" Or txtR.Text = "255" Then txtR_GotFocus
End Sub
Private Sub txtR_GotFocus()
    txtR.SelStart = 0
    txtR.SelLength = Len(txtR.Text)
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UpdateRed
        txtR.SelStart = 0
        txtR.SelLength = Len(txtR.Text)
        blnAdjust = False
        KeyAscii = 0
    End If
End Sub

Private Sub txtR_LostFocus()
    UpdateRed
    blnAdjust = False
End Sub

'for txtS comments, see corresponding in txtH
Private Sub txtS_Change()
    If blnUpdate Then Exit Sub
    If Not blnAdjust Then
        SetUndo
        blnAdjust = True
    End If
    UpdateSat
End Sub

Private Sub UpdateSat()
Dim HSLV As HSLCol
    If Val(txtS.Text) > HSLMAX Then
        txtS.Text = HSLMAX
    Else
        txtS.Text = Val(txtS.Text)
    End If
    HSLV.Hue = Val(txtH.Text)
    HSLV.Sat = Val(txtS.Text)
    HSLV.Lum = Val(txtL.Text)
    blnSat = True
    ChangeColor HSLtoRGB(HSLV)
    blnSat = False
    If Not blnBuddy Then
        blnBuddy = True
        vS.Value = 240 - Val(txtS.Text)
        blnBuddy = False
    End If
    If txtS.Text = "0" Or txtS.Text = "240" Then txtS_GotFocus
End Sub

Private Sub txtS_GotFocus()
    txtS.SelStart = 0
    txtS.SelLength = Len(txtS.Text)
End Sub

Private Sub txtS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        UpdateSat
        txtS.SelStart = 0
        txtS.SelLength = Len(txtS.Text)
        blnAdjust = False
        KeyAscii = 0
    End If
End Sub

Private Sub txtS_LostFocus()
    UpdateSat
    blnAdjust = False
End Sub

Private Sub txtVB_Change()
Dim VBV As Long
    txtVB.Text = UCase$(txtVB.Text)
    If Not isHex(txtVB.Text) Then
        Beep 'invalid hex code
        Exit Sub
    End If
    'adjust selection
    txtVB.SelStart = Len(txtVB.Text)
    'change color
    VBV = CLng("&H1" & txtVB.Text)
    'this avoids the negative if the 16th bit is set
    If VBV > 0 Then
        VBV = VBV - CLng("&H1" & String$(Len(txtVB.Text), "0"))
    Else
        VBV = 0
    End If
    
    'set color & update the codes
    If Not blnUpdate Then SetUndo
    picColor.BackColor = VBV
    If Not blnUpdate Then
        updateCodes
    End If
    If VBV = 0 Then txtVB_GotFocus
End Sub

Private Sub txtVB_GotFocus()
    txtVB.SelStart = 0
    txtVB.SelLength = Len(txtVB.Text)
    iVpos = 0
End Sub

Private Function isHex(strHex As String) As Boolean
'check that a string contains only 0-9 and A-F
Dim blnHex As Boolean, i As Integer, strChar As String * 1
    If Len(strHex) = 0 Then Exit Function
    blnHex = True
    For i = 1 To Len(strHex)
        strChar = Mid$(strHex, i, 1)
        blnHex = blnHex And ((strChar >= "0" And strChar <= "9") Or (strChar >= "A" And strChar <= "F"))
    Next i
    isHex = blnHex
End Function

'Scroll bars imitating spin buttons

Private Sub vB_Change()
    'blnBuddy keeps this event and the txt?_change events from
    'calling each other
    If Not blnBuddy Then
        blnBuddy = True
        'up is down and down is up!
        txtB.Text = 255 - VB.Value
        blnBuddy = False
    End If
End Sub

Private Sub vB_LostFocus()
    blnAdjust = False
End Sub

Private Sub vG_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtG.Text = 255 - vG.Value
        blnBuddy = False
    End If
End Sub

Private Sub vG_LostFocus()
    blnAdjust = False
End Sub

Private Sub vH_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtH.Text = 239 - vH.Value
        blnBuddy = False
    End If
End Sub

Private Sub vH_LostFocus()
    blnAdjust = False
End Sub

Private Sub vL_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtL.Text = 240 - vL.Value
        blnBuddy = False
    End If
End Sub

Private Sub vL_LostFocus()
    blnAdjust = False
End Sub

Private Sub vR_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtR.Text = 255 - vR.Value
        blnBuddy = False
    End If
End Sub

Private Sub vR_LostFocus()
    blnAdjust = False
End Sub

Private Sub vS_Change()
    If Not blnBuddy Then
        blnBuddy = True
        txtS.Text = 240 - vS.Value
        blnBuddy = False
    End If
End Sub

Private Sub BlendEm()
Dim i As Integer
    For i = 0 To 14
        lblBlend(i).BackColor = Blend(picAnchor(0).BackColor, picAnchor(1).BackColor, (i + 1) / 16)
    Next i
End Sub

Public Sub ChangeColor(lColor As Long)
    picColor.BackColor = lColor
    frmBig.BackColor = lColor
    frmBig.picImg.BackColor = lColor
    updateCodes
End Sub

Public Sub SetUndo()
Dim i As Integer, c As Integer
    c = 4
    For i = 0 To 4
        If lblUndo(i).BackColor = picColor.BackColor Then
            c = i
            Exit For
        End If
    Next i
    For i = c To 1 Step -1
        lblUndo(i).BackColor = lblUndo(i - 1).BackColor
        If lblUndo(i).BackColor <> &H8000000F Then lblUndo(i).Visible = True

    Next i
    lblUndo(0).BackColor = picColor.BackColor
    lblUndo(0).Visible = True
    lblPast.Visible = True
End Sub

Private Sub vS_LostFocus()
    blnAdjust = False
End Sub
