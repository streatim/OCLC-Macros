'MacroName:atab
'MacroDescription:Adds a command to invoke the subject authority load table.

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddField 1, "949  *atab=asub"

End Sub
