Attribute VB_Name = "basTESTvbajson"
Option Explicit
Option Compare Text
Option Private Module

'"ID","Type","Status","Priority","Milestone","Owner","Summary","AllLabels","Link"
'"vbajson1","Defect","FIXED","Medium","","","outcome","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=1
    ' How can i read a parsed JSON string as an array?
'"vbajson2","Defect","New","High","","","parseString bug","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=2
    ' I found an issue that crashes the parseString function where data delimited
    ' with a single quote and containing encoded single quotes.
    ' It causes a freeze. This can be fixed by adding a single quote to the case statement:
    '        Select Case (char)
    '           Case """", "\\", "/", "'"
    '              SB.Append char
    '              index = index + 1
    '           Case "b"
'"vbajson3","Defect","FIXED","Medium","","","Incorrect CrLf encoding?","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=3
    ' Reported by webmas...@ediy.co.nz, Mar 24, 2009
    ' Some data seemed to have double the enters in text every time it was saved,
    ' it seems to be because
    '               Case "n"
    '                 SB.Append vbNewLine
    '                 index = index + 1
    ' should be:
    '               Case "n"
    '                 SB.Append vbLf
    '                 index = index + 1
    ' in the parseString function.
'"vbajson4","Defect","FIXED","Medium","","","improve parseNumber() for other decimal settings","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=4
    ' Reported by telmo.ca...@gmail.com, Jun 12, 2009
    ' I have added to parseNumber():
    '        If InStr(Value, ".") Or InStr(Value, "e") Or InStr(Value, "E") Then
    '            ' for PT Local Settings where decimal is ","
    '            If CStr(1.2) = "1,2" Then value = Replace(value, ".", ",", 1, -1, 1)
    '            parseNumber = CDbl(Value)
    '        Else
    '            parseNumber = CInt(Value)
'"vbajson5","Defect","New","Medium","","","Added suport for JSON-RPC 2.0 in jsonlib","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=5
'"vbajson6","Defect","New","Medium","","","Enter one-line summary","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=6
'"vbajson7","Defect","New","Medium","","","Cannot parse a JSON string containing an array...","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=7
'"vbajson8","Defect","New","Medium","","","Cannot convert a 2-d array to JSON","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=8
'"vbajson9","Defect","New","Medium","","","Thank you for this code!","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=9
'"vbajson10","Defect","New","Medium","","","improve parseNumber() with Long number","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=10
'"vbajson11","Defect","New","Medium","","","double backslash parse problem","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=11
'"vbajson12","Defect","New","Medium","","","String wont parse","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=12
'"vbajson13","Defect","New","Medium","","","here's an update for office 64-bit support","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=13
'"vbajson14","Defect","New","Medium","","","Unable to parse strings containing colons - Infinite loop","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=14
'"vbajson15","Defect","New","Medium","","","Unable to handle multi-dimensional arrays","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=15
'"vbajson16","Defect","New","Medium","","","parseNumber and regional settings","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=16
'"vbajson17","Defect","New","Medium","","","http: 85","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=17
'"vbajson18","Defect","New","Medium","","","Redundant vbCrLf","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=18
'"vbajson19","Defect","New","Medium","","","Spaces improperly removed from object keys","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=19
'"vbajson20","Defect","New","Medium","","","cint in parse number is issue (wont deal with big numbers!)","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=20
'"vbajson21","Defect","New","Medium","","","parseKey with Key containing "":""","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=21
'"vbajson22","Defect","New","Medium","","","Bug: Case statement comparing one character ""\"" to ""\\""","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=22
'"vbajson23","Defect","New","Medium","","","Need a new owner","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=23
'
' Change integration:
' Add the file JSON.bas from Michael Glaser, review and integrate changes as appropriate
' VBJSON (http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html)
' is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed
'
'=============================================================================================================================
' Tasks:
' %005 -
' %004 -
' %003 -
' %002 -
' %001 - Get test result "VALIDATED" be verified automatically with online parser
' Issues:
' #006 -
' #005 -
' #004 -
' #002 - vbatest2 still kills Excel
' #001 - Run-time error '424' Object required in test vbajson1
'=============================================================================================================================

' 20141124 - v011 - FIXED vbajson3 - s/vbNewLine/vbLf
    ' FIXED vbajson4
' 20141121 - v011 - FIXED #003 - parse_test3 breaks RunAllTests

' http://stackoverflow.com/questions/244777/can-i-comment-a-json-file
' The answer is no for strict JSON interchange.
' The correct approach is here: http://blog.getify.com/json-comments/
'
' *** Online JSON Validators
' *** http://www.jsonlint.com/
' *** http://jsonformatter.curiousconcept.com/
'

Public Sub RunAllvbajsonTests()

'    vbajson1
'    Debug.Print "=> vbajson1 Finished!" & vbCrLf
    vbajson2
    Debug.Print "=> vbajson2 Finished!" & vbCrLf
Exit Sub
    vbajson3
    Debug.Print "=> vbajson3 Finished!" & vbCrLf
    vbajson4
    Debug.Print "=> vbajson4 Finished!" & vbCrLf
    vbajson5
    Debug.Print "=> vbajson5 Finished!" & vbCrLf
    vbajson6
    Debug.Print "=> vbajson6 Finished!" & vbCrLf
    vbajson7
    Debug.Print "=> vbajson7 Finished!" & vbCrLf

End Sub

Private Sub vbajson1()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strJson As String

    Debug.Print "=> vbajson1"

    ' read the JSON into an object:
    strJson = "{ bla:""hi"", ""items"": [{""it"":1,""itx"":2},{""i3"":""x""}] }"
    Debug.Print , "strJson=" & strJson
    Set o = lib.parse(strJson)

' Use Online JSON Validator to get the following validated:
'{
'    "bla": "hi",
'    "items": [
'        {
'            "it": 1,
'            "itx": 2
'        },
'        {
'            "i3": "x"
'        }
'    ]
'}

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    ' get data from arrays etc.:
    Debug.Print , "Bla: " & o.Item("bla") & " - Items of itx: " & _
        o.Item("items").Item(1).Item("itx")

    Debug.Print , "VALIDATED"

    Set lib = Nothing
    Set o = Nothing

End Sub

Private Sub vbajson2()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson2"

    Debug.Print , "vbajson2: This test will kill Excel!"
    Debug.Print , "NEEDS ERROR HANDLING"
Exit Sub

    ' read the JSON into an object:
    Set o = lib.parse("{bla:'hi I'm a single quote!'"", items: [{it:1,itx:2},{i3:'x'}] }")
   
    ' get the parsed text back:
    Debug.Print lib.toString(o)

    ' get data from arrays etc.:
    Debug.Print "Bla: " & o.Item("bla") & " - Items: " & _
        o.Item("items").Item(1).Item("itx")

    Set lib = Nothing
    Set o = Nothing

End Sub

Private Sub vbajson3()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson3"

    Debug.Print , "vbajson3: FIXED."

    Set lib = Nothing
    Set o = Nothing

End Sub

Private Sub vbajson4()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson4"

    Debug.Print , "vbajson4: FIXED. Testing needed for other locale."

    Set lib = Nothing
    Set o = Nothing

End Sub

Private Sub vbajson5()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson5"

    Debug.Print , "vbajson5: Test case needed."

    Set lib = Nothing
    Set o = Nothing

End Sub

Private Sub vbajson6()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson6"

    Debug.Print , "vbajson6: Test case needed."

    Set lib = Nothing
    Set o = Nothing

End Sub

Private Sub vbajson7()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson7"

    Debug.Print , "vbajson7: Test case needed."

    Set lib = Nothing
    Set o = Nothing

End Sub

