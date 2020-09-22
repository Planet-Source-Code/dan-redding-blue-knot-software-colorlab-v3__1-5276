VERSION 5.00
Begin VB.Form frmExpo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export Codes"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   1875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Frame fraEx 
      Caption         =   "Export..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1635
      Begin VB.OptionButton optHTML 
         Caption         =   "HTML Codes"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.OptionButton optBoth 
         Caption         =   "Both"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   1275
      End
      Begin VB.OptionButton optVB 
         Caption         =   "VB Codes"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    If optVB Then
        iExpo = 1
    ElseIf optHTML Then
        iExpo = 2
    Else
        iExpo = 3
    End If
    blnCancel = False
    Unload Me
End Sub

Private Sub Form_Load()
    If iExpo = 3 Then
        optBoth.Value = True
    ElseIf iExpo = 2 Then
        optHTML.Value = True
    Else
        optVB.Value = True
    End If
    blnCancel = True
End Sub

