Attribute VB_Name = "Module1"
Sub sample_form()
   UserForm1.Show
End Sub


Sub sample_cell()
   Call CalendarForm.setDate(Range("B2"))
   Call CalendarForm.setCallBackCell("B2")
   CalendarForm.Show
End Sub
