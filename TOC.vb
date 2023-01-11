'MacroName:TOC
'MacroDescription:Adds Table of Contents statement where the cursor is.

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.InsertText "  ßz Access table of contents"

End Sub
