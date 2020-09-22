VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RNote 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Reminder Note :  "
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_date 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer remover 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   0
      Top             =   1560
   End
   Begin VB.CommandButton cmd_mnu 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1680
      MaskColor       =   &H000000FF&
      Picture         =   "RNote.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   450
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Note 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbl_date 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date : "
      BeginProperty Font 
         Name            =   "Times"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1890
   End
   Begin VB.Label lbl_id 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuOpt 
         Caption         =   "New Note  (Ctrl + Ins)"
         Index           =   0
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Delete Note  (Ctrl + Del)"
         Index           =   1
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Hide Note  (Esc)"
         Index           =   2
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Note Manager (F2)"
         Index           =   4
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Help (f1)"
         Index           =   5
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "About Us"
         Index           =   7
      End
      Begin VB.Menu mnuOpt 
         Caption         =   "Exit (F5)"
         Index           =   8
      End
   End
End
Attribute VB_Name = "RNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim timer_count As Integer

Private Sub cmd_mnu_Click()
    PopupMenu mnuMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
   If KeyCode = 27 Then
        If Trim(Note.Text) <> "" Then
            Save Me
        End If
        unload_note Me
    
    ElseIf KeyCode = 45 And Shift = 2 Then
        load_newNote
    
    ElseIf KeyCode = 46 And Shift = 2 Then
        Me.remover.Enabled = True
        Delete Me
    
    ElseIf KeyCode = 112 Then
         frm_help.Show
    
    ElseIf KeyCode = 113 Then
         frm_all.Show
    
    ElseIf KeyCode = 114 Then
         Cdlg.ShowColor
         Note.BackColor = Cdlg.Color
    
    ElseIf KeyCode = 115 Then
         Cdlg.ShowColor
         Note.ForeColor = Cdlg.Color
         
    ElseIf KeyCode = 116 Then
        End
    ElseIf KeyCode = 117 Then
        frmAbout.Show
    ElseIf KeyCode = 83 And Shift = 2 Then
        Save Me
    Else
      ' MsgBox KeyCode & "  " & Shift
    End If
    
End Sub


Private Sub Form_Load()
    KeyPreview = True
    timer_count = 0
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox Button
End Sub

Private Sub Form_Resize()
    Note.Height = Me.Height
    Note.Width = Me.Width
    cmd_mnu.Left = Me.Width - 450
End Sub



Private Sub mnuOpt_Click(Index As Integer)
    Select Case Index
         Case 0: Call Form_KeyDown(45, 2)
         Case 1: Call Form_KeyDown(46, 2)
         Case 2: Call Form_KeyDown(27, 0)
         Case 4: Call Form_KeyDown(113, 0)
         Case 5: Call Form_KeyDown(112, 0)
         Case 7: Call Form_KeyDown(117, 0)
         Case 8: Call Form_KeyDown(116, 0)
    End Select
End Sub

Private Sub remover_Timer()
    Dim complete As Boolean
    complete = False
    
    If Me.Width > 180 Then
           Me.Width = Me.Width - 100
    Else
           complete = True
    End If
    
    If Me.Height > 405 Then
           Me.Height = Me.Height - 100
    ElseIf complete = True Then
           remover.Enabled = False
           Unload Me
    End If
End Sub

Private Sub Timer_Timer()
    timer_count = timer_count + 1
    If timer_count > 25 Then
       Timer.Enabled = False
       txt_date.ForeColor = vbBlack
       Exit Sub
    End If
    
    txt_date.ForeColor = RGB(Int(Rnd(1) * 200), Int(Rnd(1) * 500), Int(Rnd(1) * 400))
    
End Sub


