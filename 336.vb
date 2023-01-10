'MacroName:336
'MacroDescription:Add fields 336 - 338 for text, unmediated, volume.

Sub Main
   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   
   CS.AddField 1, "336  text ßb txt ß2 rdacontent"
   CS.AddField 1, "337  unmediated ßb n ß2 rdamedia"
   CS.AddField 1, "338  volume ßb nc ß2 rdacarrier"

End Sub
