'MacroName:Macro
'MacroDescription:Overarching Macro for the Mardigian Library. Contains Integrated Resource and 949 Macros at present.
'MacroDescription:A breakdown of specific macros can be found https://github.com/streatim/OCLC-Macros

'Overarching Functions with customizable values
Declare Function MacroPick() 'Order of Macros in dropdown set here.

'Macro Functions
Declare Sub Run949Macro(user, selectedMacro)
Declare Function Default949(user, callNum, formatType, formatCode)
Declare Function Multi949(user, callNum, formatType, formatCode)
Declare Function Overlay949(user, callNum, formatType, formatCode)
Declare Function Volume949(user, callNum, formatType, formatCode)
Declare Sub RunIntegratedResource(user)

'Specific Sub/Macros for Fixed and Variable Field Changes
Declare Sub IntegratedFixedFieldChanges()
Declare Function IntegratedVariableFieldChanges(user)

'Subs with Acceptable values (for list commands)
Declare Sub AllFormats()
Declare Sub AllICodes()
Declare Sub AllIStatus()

'Action Functions
Declare Function CheckVariableField(marc, checkString)
Declare Function GetBarcode()
Declare Function GetCopyInfo()
Declare Function GetDBID()
Declare Function GetFormatCode(formatType, formatCode)
Declare Function GetFormatInfo(jamesBond, marcType, marcBlvl, returnValue)
Declare Function GetICode()
Declare Function GetIStatus()
Declare Function GetIType()
Declare Function GetLocation(callNum)
Declare Function GetOverlayBNum()
Declare Function GetVolume()
Declare Function ProxyCheck(url)

'Prompt Box Functions
Declare Function DropdownBox(valueArray$(), title, qtext)
Declare Function TextBox(qtext, title, optional defaultText)
Declare Function YNBox(qtext, title)

'Main Function. Sets username variable and handles anything that requires writing or reading the record.
Sub Main
    'Set the initials used in the 949 Field.
    Dim user as String
    user = "mtist"    
    
    'Determine which Macro we are running.
    Dim selectedMacro as String
    selectedMacro = MacroPick()

    'Check the selected Macro, then run the Macro based on that.
    Dim outputLine as String
    IF selectedMacro = "Default 949" THEN Call Run949Macro(user, selectedMacro)
    IF selectedMacro = "Overlay 949" THEN Call Run949Macro(user, selectedMacro)
    IF selectedMacro = "Volumes 949" THEN Call Run949Macro(user, selectedMacro)
    IF selectedMacro = "Multicopy 949" THEN Call Run949Macro(user, selectedMacro)
    IF selectedMacro = "Integrated Resource" THEN Call RunIntegratedResource(user)

End Sub

'List all Macros in a dropdown and return the selected result.
Function MacroPick()
    Dim MacroValues$(4)
    MacroValues$(0) = "Default 949"
    MacroValues$(1) = "Overlay 949"
    MacroValues$(2) = "Volumes 949"
    MacroValues$(3) = "Multicopy 949"
    MacroValues$(4) = "Integrated Resource"

    MacroPick = DropdownBox(MacroValues$(), "Select a Macro", "Select a Macro")
End Function

'***MACRO FUNCTIONS/SUBS***

Sub Run949Macro(user, selectedMacro)
    'Set Connexion Client Variable and confirm we are logged in.
    Dim CS As Object
    Set CS = CreateObject("Connex.Client")
    IF CS.IsOnline = False Then
        CS.Logon "","",""
    End If

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

    IF selectedMacro = "Default 949" THEN outputLine = Default949(user, callNum, formatType, formatCode)
    IF selectedMacro = "Overlay 949" THEN outputLine = Overlay949(user, callNum, formatType, formatCode)
    IF selectedMacro = "Volumes 949" THEN outputLine = Volume949(user, callNum, formatType, formatCode)
    IF selectedMacro = "Multicopy 949" THEN outputLine = Multi949(user, callNum, formatType, formatCode)
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
        itemInfo$(i, 3) = GetCopyInfo()
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

    Multi949 = outputString
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
    outputString =  "949  *recs=b;bn=" & callLoc & ";ov=" & bNumValue & ";ins=" & user & ";i=" & barcode & "/sta=" & iStatus & "/loc=" & callLoc & "/ty=" & iTypeCode & "/i2=" & i2Code & "/b2=" & b2Code & ";"
    Overlay949 = outputString
End Function

'Volume 949 Process
Function Volume949(user, callNum, formatType, formatCode)
    Dim outputString as String
    'This applies different metadata for up to 4 volumes of a record.
    'Get Format (b2) Code
    Dim b2Code as String
    b2Code = GetFormatCode(formatType, formatCode)
    'Get ICode2 Value
    Dim i2Code as String
    i2Code = GetICode()
    'Get iStatus Value
    Dim iStatus as String
    iStatus = GetIStatus()
    Dim itemInfo$(4, 3)
    For i = 1 to 4
        msgBox("You will now insert information for Volume #"&i)
        'Get Call number information and estimate location.
        itemInfo(i, 0) = getLocation(callNum)
        'Get Barcode
        itemInfo(i, 1) = GetBarcode()
        'Get IType
        itemInfo(i, 2) = GetIType()
        'Get Volume info
        itemInfo(i, 3) = GetVolume()
    Next i
    'Create the output 949 string.
    outputString = "949  *recs=b;bn=" & itemInfo(1, 0) & ";ins=" & user &";/i2=" & i2Code & "/b2=" & b2Code
    For i = 1 to 4
        outputString = outputString & "i=" & itemInfo(i, 1) & "/sta=" & iStatus & "/loc=" & itemInfo(i, 0) & "/ty=" & itemInfo(i, 2) & "/v=" & itemInfo(i, 3) & ";"
    Next i

    Volume949 = outputString
End Function

'Integrated Resource Process
Sub RunIntegratedResource(user)
    'Set variables.
    Dim CS As Object
    Dim blvlCheck as String
    blvlCheck = "bis"
    'Set the CS object for Connexion macros. Make sure they're logged in.
    Set CS = CreateObject("Connex.Client")
    IF CS.IsOnline = False Then
        CS.Logon "","",""
    End If

    'At present I don't know how to confirm the record type. I am going to backend it by checking Type and BLvl.
    CS.GetFixedField "Type", typeField
    CS.GetFixedField "BLvl", blvlField
    IF (typeField = "a") AND (inStr(blvlCheck, blvlField)>0) THEN 
        MsgBox "This is should already be a Continuing Resource."
    ELSE
        Call IntegratedFixedFieldChanges()
    END IF 

    Dim outputString as String
    outputString = IntegratedVariableFieldChanges(user)
    CS.AddField 1, outputString
End Sub


'***Macro Subs/Actions***

'Fixed Field Changes for Macro: integratedResource()
Sub IntegratedFixedFieldChanges()
    Dim CS as Object
    Set CS = CreateObject("Connex.Client")
    'Set the format type to Continuing Resource.
    CS.ChangeRecordType 2

    'Set the type to A and Blvl to i. (that's the default setting)
    CS.SetFixedField "Type", "a"
    CS.SetFixedField "BLvl", "i"

    'Set S/L to 2 and Form to o. Blank out Orig.
    CS.SetFixedField "S/L", "2"
    CS.SetFixedField "Form", "o"
    CS.SetFixedField "Orig", " "

    'SrTp can be either d or w, defaulting to d.
    Dim SrTpField as String
    Dim SrTpCheck as String
    SrTpCheck = "dw"
    CS.GetFixedField "SrTp", SrTpField
    IF inStr(SrTpCheck, SrTpField) = 0 THEN
        CS.SetFixedField "SrTp", "d"
    END IF

    'Freq and Regl need to agree, so we default to 'u' for both.
    CS.SetFixedField "Freq", "u"
    CS.SetFixedField "Regl", "u"

    'DtSt can be 'c', 'd', or 'u'. We default to 'u' locally (though 'c' is the technical default)
    Dim DtStField as String
    Dim DtStCheck as String
    DtStCheck = "cdu"
    CS.GetFixedField "DtSt", DtStField
    IF inStr(DtStCheck, DtStField) = 0 THEN
        CS.SetFixedField "DtSt", "u"
        ' "," is the second Date field in Dates.
        CS.SetFixedField ",", "uuuu"
    END IF 
End Sub

'Variable Field Changes for Macro: integratedResource()
Function IntegratedVariableFieldChanges(user) 
    Dim CS as Object
    Set CS = CreateObject("Connex.Client")
    Dim publicDB as String
    publicDB = "False"
    'Set the 049
    CS.SetField 1, "049  EYDX [WEB]"
    'Loop through 500 Fields to see if any of them exist or include a "Title from homepage" string in their field.
    IF (CheckVariableField("500", "Title from homepage") = "False") THEN
        Dim msg500 as String
        CS.AddField 1, "500  Title from homepage (viewed "&Format(Now, "MMMM d, yyyy")&")"
    END IF
    'Prompt to see if this is a restricted database; if so, add a 506.
    IF (YNBox("Restricted Resource", "Is this a Restricted Resource?") = "True") THEN
        CS.AddField 1, "506  Access restricted to University of Michigan-Dearborn affiliates."
    ELSE 
        publicDB = "True"
    END IF
    'Loop through the 538 fields and make sure there is a mode of access statement.
    IF (CheckVariableField("538", "Mode of access") = "False") THEN
        CS.AddField 1, "538  Mode of access: World Wide Web."
    END IF
    'Now to update the 856. Check first to see if it's an A-Z list entry.
    IF (YNBox("In A-Z List", "Is this in the A-Z List?") = "True") THEN
        'It is an A-Z List entry. Take in the  #, delete the 856 fields, and insert the correct 856.
        Dim DBID As String
        DBID = GetDBID()
        'Confirm whether there are 856 fields or not.
        bool$ = CS.GetField("856", 1, noVal$)
        IF bool$ = "-1" THEN
            'Loop and delete 856 4 0 fields.
            Dim x as Integer
            x = 1
            
            Do
                CS.GetField "856", x, testVal$
                IF (Left(testVal$, 5) = "85640") THEN
                    CS.DeleteField "856", x
                ELSE
                    x = x+1
                END IF
                bool$ = CS.GetField("856", x, testVal$)
            Loop While bool$ <> "0"
        END IF
        'Insert the correct 856.
        CS.AddField 1, "85640ßu https://library.umd.umich.edu/verify/redirect.php?ID="&cStr(DBID)&" ßz Access Web version"
    ELSE
        'It is not an A-Z List entry. Check to see if it is a public database.
        IF (publicDB = "False") THEN
            'This is not a public database. Let us fuss with the 
            'Cycle through the 856 4 0 fields until you hit a database URL you want to use and then put the forwarding script in front of it.
            Dim DBURL As String
            bool$ = CS.GetField("856", 1, noVal$)
            IF bool$ = "-1" THEN
                X = 1
                Do
                    CS.GetField "856", x, testString$
                    IF (Left(testString$, 5) = "85640") THEN
                        'Check to see if this is the 856 they want to use.
                        Dim urlToProxy as String
                        urlToProxy = proxyCheck(testString$)
                        IF(urlToProxy <> "False") THEN
                            CS.DeleteField "856", x
                            CS.AddField x, "85640ßu https://library.umd.umich.edu/verify/fwd.php?"&cStr(urlToProxy)&" ßz Access Web version" 
                            x = x+1
                        ELSE
                            CS.DeleteField "856", x
                        END IF
                    ELSE
                        x = x+1
                    END IF
                    bool$ = CS.GetField(marc, x, testString$)
                Loop While bool$ <> "0"    
            ELSE
                'There is not a preexisting URL. Just put the forwarding script in there with a blank value.
                CS.AddField 1, "85640ßu https://library.umd.umich.edu/verify/fwd.php? ßz Access Web version"           
            END IF
        END IF
    END IF
    Dim loadTable as String
    loadTable = "b"
    'If this is a free resource from Ann Arbor, we need to add a 960 and change the loadTable.
    IF (YNBox("From Ann Arbor", "Is this a free resource from Ann Arbor?") = "True") THEN
        CS.AddField 1, "960  *recs=bio;ins="&cStr(user)&";ßnFree through UM-AA" 
        loadTable = "bio"
    END IF
    'Pass back the remaining 949 to submit.
    Dim outputLine as String
    outputLine = "949  *recs="&cStr(loadTable)&";bn=mweb;b2=9;b3=t;ins="&cStr(user)&";i=/loc=mweb/sta=w/ty=11/i2=-;" 
    IntegratedVariableFieldChanges = outputLine
End Function

'***Subs with Acceptable values (for list commands)***

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

'Display a List of all ICodes
Sub AllICodes()
    Dim ICodeList As String
    Dim ICodes$(3)
    ICodes(0)  = "-" & chr$(9) & "None"
    ICodes(1)  = "a" & chr$(9) & "CONTENT ADDED"
    ICodes(2)  = "b" & chr$(9) & "SUBJECT ADDED"
    ICodes(3)  = "c" & chr$(9) & "NOTE/SUB ADDED"
    For i = 0 to 3
        ICodeList = ICodeList & ICodes(i) & chr$(10)
    Next i
    msgbox(ICodeList)
End Sub

Sub AllIStatus()
    Dim IStatusList As String
    Dim IStatuses$(1)
    IStatuses(0)  = "-" & chr$(9) & "Available"
    IStatuses(1)  = "p" & chr$(9) & "In Process"
    For i = 0 to 1
        IStatusList = IStatusList & IStatuses(i) & chr$(10)
    Next i
    msgbox(IStatusList)
End Sub

'***Action Functions***'

'This function checks all instances of a variable MARC field to see if a string exists and lets the calling program know if it does or not.
Function CheckVariableField(marc, checkString)
    Dim CS as Object
    Set CS = CreateObject("Connex.Client")

    Dim x as Integer
    x = 1
    Dim sField as String
    Dim exists as String
    exists = "False"
    Do
        CS.GetField marc, x, sField
        IF inStr(sField, checkString) > 0 THEN
            exists = "True"
            EXIT DO        
        END IF
        x = x+1
        bool$ = CS.GetField(marc, x, sField)
    Loop While bool$ <> "0"
    CheckVariableField = exists
End Function

'Prompts the user to provide a barcode.
Function GetBarcode()
    Dim barcodeNum as String
    barcodeNum = TextBox("Scan Barcode:", "Barcode")
    IF len(barcodeNum) <> 14 THEN
        msgbox("A 14 character barcode must be scanned")
        barcodeNum = GetBarcode()
    END IF 
    GetBarcode = barcodeNum
End Function

'Prompts the user to provide copy information.
Function GetCopyInfo()
    Dim copyInput as String
    copyInput = TextBox("Copy:", "Copy Information")
    GetCopyInfo = copyInput
End Function

'Prompts the user to provide a Database ID #.
Function GetDBID()
    Dim DBIDInput as String
    DBIDInput = TextBox("Please Type in the Database ID #:","Database ID#")
    GetDBID = DBIDInput
End Function

'Prompt User for Format Code using default values.
Function GetFormatCode(formatType, formatCode)
    Dim formatInput as String
    formatInput = TextBox("Format (b2):"&chr$(10)&chr$(10)&"Suggested code: "& formatType & chr$(10) & chr$(10) & "Enter 'list' to see available codes.","949",formatCode)
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
    ICodeInput = TextBox("Icode2:" & chr$(10) & chr$(10) & "Enter 'list' to see available codes.", "Icode2","-")
    IF ICodeInput = "list" THEN
        Call AllICodes()
        ICodeInput = GetICode()
    END IF
    IF len(ICodeInput) <> 1 THEN
        msgbox("ICode2 must be a single character value.")
        ICodeInput = GetICode()
    END IF
    GetICode = ICodeInput
End Function

'This function prompts the user to provide an I Status Code.
Function GetIStatus()
    Dim IStatusInput as String
    IStatusInput = TextBox("Item Status:" & chr$(10) & chr$(10) & "Enter 'list' to see available codes.","IStatus","p")
    IF IStatusInput = "list" THEN
        Call AllIStatus() 
        IStatusInput = GetIStatus()
    END IF
    IF len(IStatusInput) <> 1 THEN
        msgbox("I Status Code must be a single character value.")
        IStatusInput = GetIStatus()
    END IF
    GetIStatus = IStatusInput
End Function

'This function prompts the user to provide an Type Code value.
Function GetIType()
    Dim iTypeInput as String
    iTypeInput = TextBox("Enter IType:", "IType", "0")
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
    callLocation = TextBox("Enter location:","Location",callDefault)
    GetLocation = callLocation
End Function

'Prompts the user for an overlay bib number.
Function GetOverlayBNum()
    Dim overlayBib as String
    overlayBib = TextBox("Enter a Bib record number to overlay. (.b is not necessary)", "Bib Overlay")
    If len(overlayBib) = 8 THEN overlayBib = ".b"&overlayBib
    IF len(overlayBib) = 9 AND Left(overlayBib, 1) = "b" THEN overlayBib = "."&overlayBib
    IF len(overlayBib) <> 10 THEN
        msgBox("Please enter a valid bib record number")
        overlayBib = GetOverlayBNum()
    END IF
    GetOverlayBNum = overlayBib
End Function

'Prompts the user for a volume.
Function GetVolume()
    Dim VolumeInput as String
    VolumeInput = TextBox("Enter Volume:", "Volume")
    GetVolume = VolumeInput
End Function

'Takes a URL and asks the user is they would like to add the library proxy to it.
Function ProxyCheck(url)
   Dim tstVal as String
   Dim fullurl as String
   tstVal = url
   strt$ = inStr(tstVal, "ßu")
   leng& = Len(tstVal)
   stp1$ = MID(tstVal, cInt(strt$)+2, leng&)
   'Check to see if there's another delimter after the URL.
    IF (inStr(stp1$, "ß") = 0) THEN
        'There is no other delimiter. All we have left is the URL.
        fullurl = TRIM(stp1$)
    ELSE
        'There is a delimiter.
        fullurl = TRIM(LEFT(stp1$, cInt(inStr(stp1$, "ß"))-1))
    END IF
   'fullurl is the full URL. That's what we return based on the result of this dialog box. But first we need to format it for the MsgBox 
    '6 is the width of a full character. 12 is the height of any character. 3 lines of 15 characters would be 45 characters. 108
    Dim msgURL1 as String
    Dim msgURL2 as String
    Dim msgURL3 as String
    msgURL1 = mid(fullurl, 1, 35)
    msgURL2 = mid(fullurl, 36, 70)
    msgURL3 = mid(fullurl, 71, 105)
    IF(len(fullurl)>90) THEN
        msgURL3 = msgURL3 & "..."
    END IF
    
    Begin Dialog YNBox 220, 100, "Proxy URL"
        Text 4, 4, 100, 20, "Do you want to proxy the following URL"
        Text 4, 24, 215, 10, msgURL1 
        Text 4, 34, 215, 10, msgURL2 
        Text 4, 44, 215, 10, msgURL3 
        OKButton 90, 70, 25, 15
        OptionGroup .ListYN
            OptionButton 10, 60, 25, 12, "Yes"
            OptionButton 40, 60, 25, 12, "No"
    End Dialog           
    Dim DialogBox AS YNBox
    Dialog DialogBox
    IF DialogBox.ListYN = 0 THEN
        OutputValue = fullurl        
    ELSE
        OutputValue = "False"
    END IF
    ProxyCheck = OutputValue
End Function

'***Prompt Box Functions***'

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

'Displays a Text Box and confirms that a result has been entered. No other validation. Will need to rebuild it so the cancel button is recognized.
Function TextBox(qtext, title, optional defaultText)
    Dim TextInput as String
    TextInput = InputBox$(qtext, title, defaultText)
    IF TextInput = "" THEN
        MsgBox("Please put a value into the text box.")
        TextInput = TextBox(qtext, title, defaultText)
    END IF
    TextBox = TextInput
End Function

'Displays a Yes/No box and returns True/False
Function YNBox(title, qtext)
    Begin Dialog YNBox 120, 65, title
        Text 4, 4, 100, 40, qtext
        OKButton 90, 20, 25, 15
        OptionGroup .ListYN
            OptionButton 10, 50, 25, 12, "Yes"
            OptionButton 40, 50, 25, 12, "No"
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