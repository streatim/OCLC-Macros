'MacroName:Bravender
'MacroDescription:Bravender

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddFieldLine 2,""

   CS.InsertText "599"

   CS.InsertText " "

   CS.InsertText " "

   CS.InsertText "Donated by Patricia Bravendar"

   CS.BackSpace

   CS.BackSpace

   CS.InsertText "er."

End Sub
