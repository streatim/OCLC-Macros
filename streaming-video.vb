'MacroName:streaming-video
'MacroDescription:streaming video url

Sub Main

   Dim CS As Object
   Set CS = CreateObject("Connex.Client")

   CS.AddFieldLine 22,""

   CS.InsertText "856"

   CS.InsertText "4"

   CS.InsertText "0"

   CS.InsertText "ßu http://library.umd.umich.edu/research/fwd.php?http://digital.films.com/PortalPlaylists.aspx?aid=1213&xtid=     ßz Access web version."

End Sub
