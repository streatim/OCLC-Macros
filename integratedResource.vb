'MacroName:intres
'MacroDescription:Create an Integrated Resource. No longer drafted.

Declare Sub FixedFieldChanges()
Declare Sub VariableFieldChanges()

Declare Function AZList()
Declare Function CheckVariableField(marc, checkString)
Declare Function ProxyCheck(url)
Declare Function YNBox(title, qtext)

Sub Main
    'Set variables.
    Dim CS As Object
    Dim typeField as String
    Dim blvlField as String
    Dim blvlCheck as String
    blvlCheck = "bis"
    'Set the CS object for Connexion macros.
    Set CS = CreateObject("Connex.Client")
    'Ensure you are logged on.
    IF CS.IsOnline = False Then
        CS.Logon "","",""
    End If
    'At present I don't know how to confirm the record type. I am going to backend it by checking Type and BLvl.
    CS.GetFixedField "Type", typeField
    CS.GetFixedField "BLvl", blvlField
    IF (typeField = "a") AND (inStr(blvlCheck, blvlField)>0) THEN 
        MsgBox "This is should already be a Continuing Resource."
    ELSE
        Call FixedFieldChanges()
    END IF 
    Call VariableFieldChanges()
End Sub

Sub FixedFieldChanges()
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

Sub VariableFieldChanges() 
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
        DBID = AZList()
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
    'Finally, add in the 949 Field for export.
    CS.AddField x, "949  *recs=b;bn=mweb;b2=9;b3=t;ins=mtist;i=/loc=mweb/sta=w/ty=11/i2=-;" 
End Sub

Function AZList()
    Begin Dialog DBIDPrompt 130, 60, "Database ID #"
      Text 4, 4, 120, 12, "Please Type in the Database ID #:"
      TextBox 4, 20, 120, 12, .DBNum
      OKButton 4, 34, 40, 20
    End Dialog
    Dim DBID AS DBIDPrompt
    Dialog DBID  
    IF DBID.DBNum = "" THEN
      MsgBox "Please put a value into the text box."
      OutputValue = AZList()
    ELSE
      OutputValue = DBID.DBNum
    END IF
    AZList = OutputValue
End Function

Function CheckVariableField(marc, checkString)
    'This function checks all potential variable fields to see if a string exists and lets the calling program know if it does or not.
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