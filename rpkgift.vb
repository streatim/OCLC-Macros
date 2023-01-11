'MacroName:rpkgift
'MacroDescription:Adds Katrina Baetz-Matthews donation statement.

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddField 1, "599  Gift from Katrina Baetz-Matthews, 2010"
   CS.AddField 1, "7102 Library donation, 2010"

End Sub

