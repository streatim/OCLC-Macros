'MacroName:intres
'MacroDescription:Create an Integrated Resource

Declare Sub FixedFieldChanges()
Declare Sub VariableFieldChanges()
Declare Function CheckVariableField(marc, checkString)
Declare Function RestrictedPrompt()

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
    'Set the 049
    CS.SetField 1, "049  EYDX [WEB]"
    'Add a title from the Homepage if not already there.
    Dim x as Integer
    Dim sField as String    
    x = 1
    'Loop through 500 Fields to see if any of them exist or include a "Title from homepage" string in their field.
    IF (CheckVariableField("500", "Title from homepage") = "False") THEN
        Dim msg500 as String
        CS.AddField 1, "500  Title from homepage (viewed "&Format(Now, "MMMM d, yyyy")&")"
    END IF
    'Prompt to see if this is a restricted database; if so, add a 506.
    IF (RestrictedPrompt() = "True") THEN
        CS.AddField 1, "506  Access restricted to University of Michigan-Dearborn affiliates."
    END IF
    'Loop through the 538 fields and make sure there is a mode of access statement.
    IF (CheckVariableField("538", "Mode of access") = "False") THEN
        CS.AddField 1, "538  Mode of access: World Wide Web."
    END IF
    '
End Sub

Function RestrictedPrompt()
    Begin Dialog IsRestricted 120, 50, "Restricted Resource"
        Text 4, 4, 100, 40, "Is this a Restricted Resource?"
        OKButton 90, 20, 25, 15
      OptionGroup .Restricted
        OptionButton 10, 15, 25, 12, "Yes"
        OptionButton 40, 15, 25, 12, "No"
    End Dialog
    Dim Restricted As IsRestricted
    Dialog Restricted
    Dim OutputValue As String
    IF Restricted.Restricted = 0 THEN
        OutputValue = "True"
    ELSE
        OutputValue = "False"
    END IF
    RestrictedPrompt = OutputValue
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
