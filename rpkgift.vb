'MacroName:rpkgift
'MacroDescription:gift lines

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.CursorRow = 18

   CS.CursorPosition = 27

   CS.CursorRow = 19

   CS.CursorPosition = 124

   CS.AddFieldLine 20,""

   CS.InsertText "599"

   CS.InsertText " "

   CS.InsertText " "

   CS.InsertText "Gift from Katrina Baetz-Matthews, 10"

   CS.BackSpace

   CS.BackSpace

   CS.InsertText "2010."

   CS.AddFieldLine 21,""

   CS.InsertText "710"

   CS.InsertText "2"

   CS.InsertText " "

   CS.InsertText "Library donation, 2010."

End Sub

