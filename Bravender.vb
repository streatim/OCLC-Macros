'MacroName:Bravender
'MacroDescription:Adds Patricia Bravender donation statement.

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddField 1, "599 Donated by Patricia Bravender"

End Sub
