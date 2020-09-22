VERSION 5.00
Begin VB.Form frm_all 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "All Notes"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_opt 
      Caption         =   "Pack"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmd_opt 
      Caption         =   "Save"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   13
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   12
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmd_move 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txt_note 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox chk_del 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txt_note 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txt_note 
      Height          =   1605
      Index           =   1
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox txt_note 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   615
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   495
      Index           =   1
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2895
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   4
      Left            =   1680
      TabIndex        =   8
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deleted"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frm_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim All_Rs As New ADODB.Recordset

Private Sub cmd_move_Click(Index As Integer)
 On Error GoTo Err_Occured
  Select Case Index
     Case 0:    All_Rs.MoveFirst
     
     Case 1:    All_Rs.MovePrevious
                If All_Rs.BOF = True Then
                    All_Rs.MoveFirst
                End If
                    
     Case 2:    All_Rs.MoveNext
                If All_Rs.EOF = True Then
                    All_Rs.MoveLast
                End If
                
     Case 3:    All_Rs.MoveLast
  End Select
Err_Occured:
End Sub

Private Sub cmd_opt_Click(Index As Integer)
  Select Case Index
    Case 0:   All_Rs.Update
    Case 1:
              MsgBox "This will Permanently delete " & vbCrLf & _
                     "all marked notes..." & vbCrLf & vbCrLf & _
                     "Are you Sure...?", vbYesNo + vbQuestion, "Pack"
              
              If vbYes Then
                  Start.DB_Conn.Execute "Delete from rem_notes where del = true"
                  All_Rs.Requery
              End If
  End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
   KeyPreview = True
   All_Rs.Open "select id,text,cdate,del,sdate from rem_notes order by id", Start.DB_Conn, adOpenDynamic, adLockOptimistic
   
   Set txt_note(0).DataSource = All_Rs
   txt_note(0).DataField = "id"
   Set txt_note(1).DataSource = All_Rs
   txt_note(1).DataField = "text"
   Set txt_note(2).DataSource = All_Rs
   txt_note(2).DataField = "cdate"
   Set txt_note(3).DataSource = All_Rs
   txt_note(3).DataField = "sdate"
   Set chk_del.DataSource = All_Rs
   chk_del.DataField = "del"
   
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   All_Rs.Close
End Sub
