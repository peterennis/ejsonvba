Attribute VB_Name = "basTESTvbajson"
Option Explicit

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
'"vbajson3","Defect","New","Medium","","","Incorrect CrLf encoding?","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=3
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
'"vbajson4","Defect","New","Medium","","","improve parseNumber() for other decimal settings","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=4
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
' *** Online JSON Validator
' *** http://www.jsonlint.com/
'
Public Sub vbajson1()

    Dim S As New jsonlib
    Dim o As Object
    Dim strJson As String

    ' read the JSON into an object:
    strJson = "{ bla:""hi"", ""items"": [{""it"":1,""itx"":2},{""i3"":""x""}] }"
    Debug.Print "strJson=" & strJson
    Set o = S.parse(strJson)

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
    Debug.Print S.toString(o)

    ' get data from arrays etc.:
    Debug.Print "Bla: " & o.Item("bla") & " - Items of itx: " & _
        o.Item("items").Item(1).Item("itx")

End Sub

Public Sub vbajson2()

    Dim S As New jsonlib
    Dim o As Object

    Debug.Print "vbajson2: This test will kill Excel!" & vbCrLf & _
        "    NEEDS ERROR HANDLING"

Exit Sub

    ' read the JSON into an object:
    Set o = S.parse("{bla:'hi I'm a single quote!'"", items: [{it:1,itx:2},{i3:'x'}] }")
   
    ' get the parsed text back:
    Debug.Print S.toString(o)

    ' get data from arrays etc.:
    Debug.Print "Bla: " & o.Item("bla") & " - Items: " & _
        o.Item("items").Item(1).Item("itx")

End Sub

Public Sub vbajson3()

    Dim S As New jsonlib
    Dim o As Object

    Debug.Print "vbajson3: Test case needed."

End Sub

