'MacroName:Integrate949
'MacroDescription:Create an Integrated 949 Function.
Declare Sub AllFormats()

Declare Function Default949(user, callNum, formatType, formatCode)
Declare Function GetBarcode()
Declare Function GetFormatCode(formatType, formatCode)
Declare Function GetFormatInfo(jamesBond, marcType, marcBlvl, returnValue)
Declare Function GetICode()
Declare Function GetIStatus()
Declare Function GetIType()
Declare Function GetLocation(callNum)
Declare Function getOverlayBNum()
Declare Function MacroPick()
Declare Function Overlay949(user, callNum, formatType, formatCode)
Declare Function DropdownBox(valueArray$(), title, qtext)
Declare Function YNBox(title, qtext)

'Main Function. Sets username variable and handles anything that requires writing or reading the record.
Sub Main
    'Set the initials used in the 949 Field.
    Dim user as String
    user = "mtist"    
    
    'Set Connexion Client Variable and confirm we are logged in.
    Dim CS As Object
    Set CS = CreateObject("Connex.Client")
    IF CS.IsOnline = False Then
        CS.Logon "","",""
    End If

    'Determine which Macro we are running.
    Dim selectedMacro as String
    selectedMacro = MacroPick()

    'Get the Type, blvl, and 007 Field.
    Dim jamesBond as String
    Dim marcType as String
    Dim marcBlvl as String
    CS.GetField "007", 1, jamesBond
    CS.GetFixedField "Type", marcType
    CS.GetFixedField "blvl", marcBlvl
    'Determine default format Types and formatCodes based on the gathered values.
    Dim formatType as String
    Dim formatCode as String
    formatType = GetFormatInfo(jamesBond, marcType, marcBlvl, "type")
    formatCode = GetFormatInfo(jamesBond, marcType, marcBlvl, "code")
    'Get Call Number, if it exists.
    Dim callNum as String
    CS.GetField "050", 1, callNum

    'Check the selected Macro, then set the output line based on that.
    Dim outputLine as String
    IF selectedMacro = "Default 949" THEN outputLine = Default949(user, callNum, formatType, formatCode)
    IF selectedMacro = "Overlay 949" THEN outputLine = Overlay949(user, callNum, formatType, formatCode)

    CS.AddField 1, outputLine
End Sub

'Default 949 Process
Function Default949(user, callNum, formatType, formatCode)
    Dim outputString as String
    'Get Call number information and estimate location.
    Dim callLoc as String
    callLoc = getLocation(callNum)
    'Get Barcode
    Dim barcode as String
    barcode = GetBarcode()
    'Get IType
    Dim iTypeCode as String
    iTypeCode = GetIType()
    'Get Format (b2) Code
    Dim b2Code as String
    b2Code = GetFormatCode(formatType, formatCode)
    'Get ICode2 Value
    Dim i2Code as String
    i2Code = GetICode()
    'Get iStatus Value
    Dim iStatus as String
    iStatus = GetIStatus()
    'Create the output 949 string.
    outputString = "949  *recs=b;bn=" & callLoc & ";ins=" & user & ";i=" & barcode & "/sta=" & iStatus & "/loc=" & callLoc & "/ty=" & iTypeCode & "/i2=" & i2Code & "/b2=" & b2Code & ";"
    Default949 = outputString
End Function

'Multiple 949 Process
Function Multi949(user, callNum, formatType, formatCode)
    Dim outputString as String
    'In this there are two items, with separate barcodes, locations, item types (but the same i2, b2, call number) Will confirm with Heidi what this is used for.
    'Set up item 1

    Dim itemInfo$(2, 3)
    For i = 1 to 2
        msgBox("You will now insert information for Item #"&i)
        'Get Call number information and estimate location.
        itemInfo$(i, 0) = getLocation(callNum)
        'Get Barcode
        itemInfo$(i, 1) = getBarcode()
        'Get IType
        itemInfo$(i, 2) = getIType()
        'Get Copy Information
        itemInfo$(i, 3) = GetCopyInfo
    Next i
    'Get Format (b2) Code
    Dim b2Code as String
    b2Code = GetFormatCode(formatType, formatCode)
    'Get ICode2 Value
    Dim i2Code as String
    i2Code = GetICode()
    'Get iStatus Value
    Dim iStatus as String
    iStatus = GetIStatus()
    'Create the output 949 string. 
    outputString = "949  *recs=b;ins=" & user & ";i=" & itemInfo$(1, 1) & "/sta=" & iStatus & "/loc=" & itemInfo$(1, 0) & "/ty=" & itemInfo$(1, 2) & "/cop=" & itemInfo$(1, 3) & "/i2=" & i2Code & "/b2=" & b2Code & ";i=" & itemInfo$(2, 1) & "/sta=" & iStatus & "/loc=" & itemInfo$(2, 0) & "/ty=" & itemInfo$(2, 2) & "/cop=" & itemInfo$(2, 3) & ";"

    Default949 = outputString
End Function

'Overlay 949 Process
Function Overlay949(user, callNum, formatType, formatCode)
    Dim outputString as String
    'Get Call number information and estimate location.
    Dim callLoc as String
    callLoc = getLocation(callNum)
    'Get overlay BNumber
    Dim bNumValue as String
    bNumValue = getOverlayBNum()
    'Get Barcode
    Dim barcode as String
    barcode = GetBarcode()
    'Get IType
    Dim iTypeCode as String
    iTypeCode = GetIType()
    'Get Format (b2) Code
    Dim b2Code as String
    b2Code = GetFormatCode(formatType, formatCode)
    'Get ICode2 Value
    Dim i2Code as String
    i2Code = GetICode()
    'Get iStatus Value
    Dim iStatus as String
    iStatus = GetIStatus()
    'Create the output 949 string.
    outputString =  "949  *recs=b;bn=" & callLoc & ";ov=" & bNumValue & "ins=" & user & ";i=" & barcode & "/sta=" & iStatus & "/loc=" & callLoc & "/ty=" & iTypeCode & "/i2=" & i2Code & "/b2=" & b2Code & ";"
    Overlay949 = outputString
End Function

'Display a list of all format types
Sub AllFormats()
    Dim formatList As String
    Dim formats$(19)
    formats(0)  = "a" & chr$(9) & "book"
    formats(1)  = "0" & chr$(9) & "periodical"
    formats(2)  = "@" & chr$(9) & "ebook"
    formats(3)  = "g" & chr$(9) & "DVD"
    formats(4)  = "2" & chr$(9) & "videotape"
    formats(5)  = "3" & chr$(9) & "laserdisc"
    formats(6)  = "j" & chr$(9) & "CD, music"
    formats(7)  = "I" & chr$(9) & "CD, spoken"
    formats(8)  = "8" & chr$(9) & "tape, music"
    formats(9)  = "6" & chr$(9) & "tape, spoken"
    formats(10) = "m" & chr$(9) & "software"
    formats(11) = "9" & chr$(9) & "website"
    formats(12) = "e" & chr$(9) & "map"
    formats(13) = "c" & chr$(9) & "music score"
    formats(14) = "t" & chr$(9) & "thesis"
    formats(15) = "o" & chr$(9) & "kit"
    formats(16) = "r" & chr$(9) & "3D object"
    formats(17) = "4" & chr$(9) & "slide"
    formats(18) = "h" & chr$(9) & "record album"
    For i = 0 to 18
        formatList = formatList & formats(i) & chr$(10)
    Next i
    msgbox(formatList)
End Sub

'Prompts the user to provide a barcode.
Function GetBarcode()
    Dim barcodeNum as String
    barcodeNum = InputBox$("Scan Barcode:", "Barcode")
    IF len(barcodeNum) <> 14 THEN
        msgbox("A 14 character barcode must be scanned")
        barcodeNum = GetBarcode()
    END IF 
    GetBarcode = barcodeNum
End Function

'Prompts the user to provide copy information.
Function GetCopyInfo()
    Dim copyInput as String
    copyInput = InputBox$("Copy:", "Copy Information")
    GetCopyInfo = copyInput
End Function

'Prompt User for Format Code using default values.
Function GetFormatCode(formatType, formatCode)
    Dim formatInput as String
    formatInput = InputBox$("Format (b2):"&chr$(10)&chr$(10)&"Suggested code: "& formatType & chr$(10) & chr$(10) & "Enter 'list' to see available codes.","949",formatCode)
    IF formatInput = "list" THEN 
        Call AllFormats()
        formatInput = GetFormatCode(formatType, formatCode)
    END IF
    IF len(formatInput) <> 1 THEN
        msgbox("Format (b2) Code should be a single character.")
        formatInput = GetFormatCode(formatType, formatCode)
    End If
    GetFormatCode = formatInput
End Function

'Take in 007 info, Type, and blvl and determine the Format Type and Code.
Function GetFormatInfo(jamesBond, marcType, marcBlvl, returnValue)
    Dim outputArray$(1)
    '0 is Type, 1 is Code. Set the defaults.
    outputArray$(0) = "Unknown, please select"
    outputArray$(1) = "-"
    Dim returnString as String
    'Check for an 007 field and set the m7 fields. This is stolen shamelessly from Patrick's original function.
    Dim m7a as String
    Dim m7b as String
    Dim m7e as String
    Dim m7g as String
    Dim jamesBond2 as String
    jamesBond2=chr$(223)+"a "+mid$(jamesBond,6)+" "
    for x=2 to (len(jamesBond2)/5)+1
        if mid$(getfield (jamesBond2,x,chr$(223)),1,1)="a" then m7a=mid$(getfield (jamesBond2,x,chr$(223)),3,1)
        if mid$(getfield (jamesBond2,x,chr$(223)),1,1)="b" then m7b=mid$(getfield (jamesBond2,x,chr$(223)),3,1)
        if mid$(getfield (jamesBond2,x,chr$(223)),1,1)="e" then m7e=mid$(getfield (jamesBond2,x,chr$(223)),3,1)
        if mid$(getfield (jamesBond2,x,chr$(223)),1,1)="g" then m7g=mid$(getfield (jamesBond2,x,chr$(223)),3,1)
    next x
    'Match the combination of values to determine the type and code.
    if marcType="a" and marcBlvl="m" then outputArray$(0)="Book":outputArray$(1)="a"
    if marcType="a" and marcBlvl="s" then outputArray$(0)="Serial" & chr$(10) & chr$(10) & "Please choose whether item is a book (a) or periodical (0)":outputArray$(1)="a or 0"
    if marcType="m" and marcBlvl="m" then outputArray$(0)="Software":outputArray$(1)="m"
    if marcType="e" then outputArray$(0)="Map":outputArray$(1)="e"
    if marcType="c" then outputArray$(0)="Music Score":outputArray$(1)="c"
    if marcType="t" then outputArray$(0)="Thesis":outputArray$(1)="t"
    if marcType="o" then outputArray$(0)="Kit":outputArray$(1)="o"
    if marcType="r" then outputArray$(0)="3D Object":outputArray$(1)="r"
    if marcType="g" and m7a="v" and m7e="v" then outputArray$(0)="DVD":outputArray$(1)="g"
    if marcType="g" and m7a="v" and m7e="g" then outputArray$(0)="Laserdisc":outputArray$(1)="3"
    if marcType="g" and m7a="v" and m7b="f" then outputArray$(0)="Videocassette":outputArray$(1)="2"
    if marcType="g" and m7a="g" and m7e="s" then outputArray$(0)="Slide":outputArray$(1)="4"
    if marcType="a" and marcBlvl="m" and m7a="c" and m7b="r" then outputArray$(0)="ebook or website":outputArray$(1)="@ or 9"
    if marcType="m" and marcBlvl="m" and m7a="c" and m7b="r" then outputArray$(0)="ebook or website":outputArray$(1)="@ or 9" 
    if marcType="j" and m7a="s" and m7b="d" and m7g="g" then outputArray$(0)="CD, Music":outputArray$(1)="j"
    if marcType="i" and m7a="s" and m7b="d" and m7g="g" then outputArray$(0)="CD, Spoken":outputArray$(1)="i"
    if marcType="j" and m7a="s" and m7b="s" then outputArray$(0)="tape, Music":outputArray$(1)="8"
    if marcType="i" and m7a="s" and m7b="s" then outputArray$(0)="tape, spoken":outputArray$(1)="6"
    if marcType="i" and m7a="s" and m7b="d" and m7g<>"g" then outputArray$(0)="Record Album":outputArray$(1)="h"
    if marcType="j" and m7a="s" and m7b="d" and m7g<>"g" then outputArray$(0)="Record Album":outputArray$(1)="h"

    IF returnValue = "type" THEN
        returnString = outputArray$(0)
    END IF
    IF returnValue = "code" THEN
        returnString = outputArray$(1)
    END IF
    GetFormatInfo = returnString
End Function

'This function prompts the user to provide an I Code.
Function GetICode()
    Dim ICodeInput as String
    ICodeInput = InputBox$("Icode2:" & chr$(10) & chr$(10) & "Choose from:" & chr$(10) & "-    None" & chr$(10) & "a    CONTENT ADDED" & chr$(10) & "b    SUBJECT ADDED" & chr$(10)& "c    NOTE/SUB ADDED","ICode2","-")
    IF len(ICodeInput) <> 1 THEN
        msgbox("ICode2 must be a single character value.")
        ICodeInput = GetICode()
    END IF
    GetICode = ICodeInput
End Function

'This function prompts the user to provide an I Status Code.
Function GetIStatus()
    Dim IStatusInput as String
    IStatusInput = InputBox$("Item Status:" & chr$(10) & chr$(10) & "Choose from:" & chr$(10) & "-   Available" & chr$(10) & "p   In Process","IStatus","p")
    IF len(IStatusInput) <> 1 THEN
        msgbox("I Status Code must be a single character value.")
        IStatusInput = GetIStatus()
    END IF
    GetIStatus = IStatusInput
End Function

'This function prompts the user to provide an Type Code value.
Function GetIType()
    Dim iTypeInput as String
    iTypeInput = InputBox$("Enter IType:", "IType", "0")
    IF iTypeInput = "" THEN
        msgBox("Input an IType value.")
        iTypeInput = GetIType()
    End If
    GetIType = iTypeInput
End Function

'Prompts the user to confirm the location. Defaults to c3rd or c4th based on the call number.
Function GetLocation(callNum)
    dim callLocation as String 
    dim callDefault as String
    'Set default Call # Location based on Call Number, if it exists.
    callDefault = ""
    If callNum <> "" THEN
        if asc(mid$(callNum,6,1))<=75 then callDefault="c3rd"
        if asc(mid$(callNum,6,1))>=76 then callDefault="c4th"
    END IF
    'Prompt the user to correct the location, if necessary.
    callLocation = InputBox$("Enter location:","Location",callDefault)
    IF callLocation = "" THEN 
        msgbox("Please provide a location.")
        callLocation = GetLocation(callNum)
    End If
    getLocation = callLocation
End Function

'Prompts the user for an overlay bib number.
Function getOverlayBNum()
    Dim overlayBib as String
    overlayBib = InputBox$("Enter a Bib record number to overlay. (.b is not necessary)", "Bib Overlay")
    If len(overlayBib) = 8 THEN overlayBib = ".b"&overlayBib
    IF len(overlayBib) = 9 AND Left(overlayBib, 1) = "b" THEN overlayBib = "."&overlayBib
    IF len(overlayBib) <> 10 THEN
        msgBox("Please enter a valid bib record number")
        overlayBib = getOverlayBNum()
    END IF
    getOverlayBNum = overlayBib
End Function

'List all Macros in a dropdown and return the selected result.
Function MacroPick()
    Dim MacroValues$(3)
    MacroValues$(0) = "Default 949"
    MacroValues$(1) = "Overlay 949"
    MacroValues$(2) = "Volumes 949"
    MacroValues$(3) = "Multi 949"

    MacroPick = DropdownBox(MacroValues$(), "Select a Macro", "Select a Macro")
End Function

'Displays a dropdown and returns the result.
Function DropdownBox(valueArray$(), title, qtext)
   Begin Dialog TypeDrop 120, 50, title
      Text 4, 4, 100, 40, qtext
      OKButton 90, 20, 25, 15
      DropListBox 10,20,70,100,valueArray$(),.returnValue
   End Dialog           
   Dim DialogBox AS TypeDrop
   Dialog DialogBox
   DropdownBox = valueArray(DialogBox.returnValue)
End Function

'Displays a Yes/No box and returns True/False
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
