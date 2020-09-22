VERSION 5.00
Begin VB.Form Rem 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Reminder : "
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   2430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   2430
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Enter Text to be remembered"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu Main 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu Opt 
         Caption         =   "New Reminder"
         Index           =   0
      End
      Begin VB.Menu Opt 
         Caption         =   "Save"
         Index           =   1
      End
      Begin VB.Menu Opt 
         Caption         =   "Delete"
         Index           =   2
      End
      Begin VB.Menu Opt 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Opt 
         Caption         =   "List All"
         Index           =   4
      End
      Begin VB.Menu Opt 
         Caption         =   "Change Colour"
         Index           =   5
      End
   End
End
Attribute VB_Name = "Rem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        PopupMenu Main
End Sub

Private Sub Form_Resize()
   Text.Width = Me.Width
   Text.Height = Me.Height
End Sub

