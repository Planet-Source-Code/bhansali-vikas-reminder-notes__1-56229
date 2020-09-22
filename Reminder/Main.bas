Attribute VB_Name = "Start"
Option Explicit

Public DB_Conn As New ADODB.Connection

Dim Rs_Notes As New ADODB.Recordset
Dim Rs_Max As New ADODB.Recordset

Dim Last_Id As Integer
Dim Last_top As Integer
Dim Last_left As Integer

Dim frm As RNote

Sub Main()
   Dim date_diff As Integer
   Dim Note_Counter As Integer
   

   
   frmAbout.Height = frmAbout.Height - 500
   frmAbout.BorderStyle = 0
   frmAbout.Show
   
   DB_Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBReminder.mdb;Persist Security Info=False"
   DB_Conn.Open
   
   Last_Id = 0
   Note_Counter = 0
   
   Rs_Notes.Open "select * from Rem_Notes where del = false", DB_Conn, adOpenDynamic, adLockOptimistic
   While Rs_Notes.EOF <> True
     Set frm = New RNote
     
     date_diff = DateDiff("d", VBA.Format$(Date, "dd-mmm-yy"), VBA.Format$(Rs_Notes.Fields(3), "dd-mmm-yy"))
     
     If date_diff = 0 Then
         frm.Timer.Enabled = True
     End If
     
     If date_diff <= 0 Then
          Set_Data frm
          Note_Counter = Note_Counter + 1
     End If
     
     Rs_Notes.MoveNext
   Wend
   
   Rs_Notes.Close
    
    Last_top = 5500
    Last_left = 9500
    
    Last_Id = Find_Last_Id
    
    If Note_Counter = 0 Then
        load_newNote
    End If
 End Sub

Sub Set_Data(frm_note As RNote)
          frm.lbl_id.Caption = Rs_Notes.Fields(0)
          frm.Caption = frm.Caption & Rs_Notes.Fields(0)
          frm.Note.Text = Rs_Notes.Fields(1)
          frm.txt_date = VBA.Format$(Rs_Notes.Fields(2), "dd-mmm-yy")
          
          frm.Left = Rs_Notes.Fields(5)
          frm.Top = Rs_Notes.Fields(6)
          
          frm.Note.BackColor = Rs_Notes.Fields(7)
          frm.Note.ForeColor = Rs_Notes.Fields(8)
          
          frm.Width = Rs_Notes.Fields(9)
          frm.Height = Rs_Notes.Fields(10)
          
          frm.Show
End Sub


Sub load_newNote()
   Set frm = New RNote
   Last_Id = Last_Id + 1
   frm.lbl_id.Caption = Last_Id
   frm.Caption = frm.Caption & Last_Id
   frm.Top = Last_top
   frm.Left = Last_left
   frm.txt_date = Format(Date, "dd-mmm-yy")
   frm.Show
   
   Last_top = Last_top - 500
   If (Last_top < 500) Then
       Last_top = 5500
   End If
   
   Last_left = Last_left - 500
   If (Last_left < 500) Then
       Last_left = 9500
   End If
   
End Sub


Sub unload_note(frm_note As RNote)
    Unload frm_note
End Sub


Sub Save(frm_note As RNote)
     
     Rs_Notes.Open "select * from rem_notes where id = " & Val(frm_note.lbl_id.Caption), DB_Conn, adOpenDynamic, adLockOptimistic
     
     If Rs_Notes.EOF = True Then
        Rs_Notes.AddNew
        frm_note.lbl_id.Caption = Find_Last_Id + 1
     End If
            
     Rs_Notes.Fields(0) = Val(frm_note.lbl_id.Caption)
     Rs_Notes.Fields(1) = frm_note.Note
     Rs_Notes.Fields(2) = frm_note.txt_date
     Rs_Notes.Fields(4) = False
     Rs_Notes.Fields(5) = frm_note.Left
     Rs_Notes.Fields(6) = frm_note.Top
     Rs_Notes.Fields(7) = frm_note.Note.BackColor
     Rs_Notes.Fields(8) = frm_note.Note.ForeColor
     Rs_Notes.Fields(9) = frm_note.Width
     Rs_Notes.Fields(10) = frm_note.Height
     
     Rs_Notes.Update
    
     Rs_Notes.Close
End Sub



Sub Delete(frm_note As RNote)
     Rs_Notes.Open "select * from rem_notes where id = " & Val(frm_note.lbl_id.Caption), DB_Conn, adOpenDynamic, adLockOptimistic
     
     If Rs_Notes.EOF = True Then
        Unload frm_note
        Rs_Notes.Close
        Exit Sub
     End If
     
    Rs_Notes.Fields(4) = True
    Rs_Notes.Update
    Rs_Notes.Close
End Sub


Function Find_Last_Id() As Integer
   
   Rs_Max.Open "select * from max_id", DB_Conn, adOpenDynamic, adLockOptimistic
     
         If IsNull(Rs_Max.Fields(0)) Then
             Find_Last_Id = 0
         Else
             Find_Last_Id = Val(Rs_Max.Fields(0))
         End If
     
   Rs_Max.Close
   
End Function
