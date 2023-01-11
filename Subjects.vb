'MacroName:Subjects
'MacroDescription:Subject headings

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddFieldLine 17,""

   CS.InsertText "650"

   CS.InsertText " "

   CS.InsertText "0"

   CS.InsertText "Detective and mystery storires"

   CS.BackSpace

   CS.BackSpace

   CS.BackSpace

   CS.BackSpace

   CS.InsertText "ies."

   CS.AddFieldLine 18,""

   CS.InsertText "650"

   CS.InsertText " "

   CS.InsertText "0"

   CS.InsertText "Drew, Nancy (Fictitious character) ßv Juvenile fiction."

   CS.AddFieldLine 19,""

   CS.InsertText "650"

   CS.InsertText " "

   CS.InsertText "0"

   CS.InsertText "Women detectives ßv Juvenile fiction."

   CS.AddFieldLine 20,""

   CS.InsertText "650"

   CS.InsertText " "

   CS.InsertText "0"

   CS.InsertText "Adventure stories ßv Juvenile fiction."

End Sub
