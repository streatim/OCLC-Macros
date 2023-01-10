'MacroName:atab
'MacroDescription:atab

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddFieldLine 2,""

   CS.InsertText "949"

   CS.InsertText " "

   CS.InsertText " "

   CS.InsertText "*atab=asub"

End Sub
