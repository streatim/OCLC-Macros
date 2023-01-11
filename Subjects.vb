'MacroName:Subjects
'MacroDescription:Adds subject headings to Nancy Drew books.

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddField 1, "650 0Detective and mystery stories."
   CS.AddField 1, "650 0Drew, Nancy (Fictitious character) ßv Juvenile fiction."
   CS.AddField 1, "650 0Women detectives ßv Juvenile fiction."
   CS.AddField 1, "650 0Adventure stories ßv Juvenile fiction."

End Sub
