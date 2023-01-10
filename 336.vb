'MacroName:336
'MacroDescription:336-338

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddFieldLine 24,""

   CS.InsertText "336"

   CS.TabKey True

   CS.TabKey True

   CS.InsertText "test"

   CS.BackSpace

   CS.BackSpace

   CS.InsertText "xt ßb txt ß2 rdacontent"

   CS.AddFieldLine 25,""

   CS.InsertText "337"

   CS.TabKey True

   CS.TabKey True

   CS.InsertText "unmediated ßb n ß2 rdamedia"

   CS.AddFieldLine 26,""

   CS.InsertText "338"

   CS.TabKey True

   CS.TabKey True

   CS.InsertText "volume ßb nc ß2 rdacarrier"

   CS.AddFieldLine 27,""

End Sub
