'MacroName:Integrate949
'MacroDescription:Create an Integrated 949 Function.
Declare Sub Overlay(user)
Declare Function AllFormats()
Declare Function TypeDropdown()
Declare Function YNBox(title, qtext)


Sub Main
   'Set Variables
   Dim CS As Object
   Set CS = CreateObject("Connex.Client")
   Dim user as String
   'Set the initials used in the 949 Field.
   user = "mtist"

End Sub

Sub Overlay(user)
    msgbox(user)
End Sub

Function AllFormats()
    Dim formatList As String
    formatList = "a    book"&chr$(10)&"0    periodical"&chr$(10)&"@    ebook"&chr$(10)&"g    DVD"&chr$(10)&"2    videotape"&chr$(10)&"3    laserdisc"&chr$(10)&"j    CD, music"&chr$(10)&"I    CD, spoken"&chr$(10)&"8    tape, music"&chr$(10)&"6    tape, spoken"&chr$(10)&"m    software"&chr$(10)&"9    website"&chr$(10)&"e    map"&chr$(10)&"c    music score"&chr$(10)&"t    thesis"&chr$(10)&"o    kit"&chr$(10)&"r    3D object"&chr$(10)&"4    slide"&chr$(10)&"h    record album"&chr$(10)
    msgbox(formatList)
End Function

Function TypeDropdown()
   Dim TypeValues$(3)
   TypeValues(0) = "Default"
   TypeValues(1) = "Overlay"
   TypeValues(2) = "Volumes"
   TypeValues(3) = "Multi"

   
   Begin Dialog TypeDrop 120, 50, "Test Dropdown"
      Text 4, 4, 100, 40, "Type of 949"
      OKButton 90, 20, 25, 15
      DropListBox 10,20,50,100,TypeValues$(),.949Type
   End Dialog           
   Dim DialogBox AS TypeDrop
   Dialog DialogBox
   msgbox(DialogBox.949Type)
End Function

Function YNBox(title, qtext)
    Begin Dialog YNBox 120, 50, title
        Text 4, 4, 100, 40, qtext
        OKButton 90, 20, 25, 15
        OptionGroup .ListYN
            OptionButton 10, 15, 25, 12, "Yes"
            OptionButton 40, 15, 25, 12, "No"
    End Dialog           
    Dim DialogBox AS YNBox
    Dialog DialogBox
    IF DialogBox.ListYN = 0 THEN
        OutputValue = "True"        
    ELSE
        OutputValue = "False"
    END IF
    YNBox = OutputValue
End Function 
