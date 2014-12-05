Attribute VB_Name = "basTESTvbajson"
Option Explicit
Option Compare Text
Option Private Module

' http://stackoverflow.com/questions/244777/can-i-comment-a-json-file
' The answer is no for strict JSON interchange.
' The correct approach is here: http://blog.getify.com/json-comments/
'
' *** Online JSON Validators
' *** http://www.jsonlint.com/
' *** http://jsonformatter.curiousconcept.com/
' *** http://www.freeformatter.com/json-formatter.html
' *** http://www.jsoneditoronline.org/ (validated by jsonlint)
' *** http://json.parser.online.fr/
' *** http://www.jsontest.com/ (JSONTest.com is a testing platform for services utilizing JSON)
' *** http://www.ist.rit.edu/~jxs/services/JSON/ (JSON Explorer)
' *** http://json-ld.org/ ***

Public Sub RunAllvbajsonTests()

    vbajson1
    Debug.Print "=> vbajson1 Finished!" & vbCrLf
    vbajson1a
    Debug.Print "=> vbajson1a Finished!" & vbCrLf
Exit Sub
'    vbajson2
'    Debug.Print "=> vbajson2 Finished!" & vbCrLf
'    vbajson3
'    Debug.Print "=> vbajson3 Finished!" & vbCrLf
'    vbajson4
'    Debug.Print "=> vbajson4 Finished!" & vbCrLf
'    vbajson5
'    Debug.Print "=> vbajson5 Finished!" & vbCrLf
    vbajson6
    Debug.Print "=> vbajson6 Finished!" & vbCrLf
'    vbajson7
'    Debug.Print "=> vbajson7 Finished!" & vbCrLf
'    vbajson7b
'    Debug.Print "=> vbajson7b Finished!" & vbCrLf
'    vbajson8
'    Debug.Print "=> vbajson8 Finished!" & vbCrLf
'    vbajson8b
'    Debug.Print "=> vbajson8b Finished!" & vbCrLf
'    vbajson8c
'    Debug.Print "=> vbajson8c Finished!" & vbCrLf
'    vbajson9
'    Debug.Print "=> vbajson9 Finished!" & vbCrLf
'    vbajson10
'    Debug.Print "=> vbajson10 Finished!" & vbCrLf
'    vbajson10a
'    Debug.Print "=> vbajson10a Finished!" & vbCrLf
'    vbajson11
'    Debug.Print "=> vbajson11 Finished!" & vbCrLf
'    vbajson12
'    Debug.Print "=> vbajson12 Finished!" & vbCrLf
'    vbajson13
'    Debug.Print "=> vbajson13 Finished!" & vbCrLf
    vbajson14
    Debug.Print "=> vbajson14 Finished!" & vbCrLf
    vbajson15
    Debug.Print "=> vbajson15 Finished!" & vbCrLf
    vbajson16
    Debug.Print "=> vbajson16 Finished!" & vbCrLf
    vbajson16a
    Debug.Print "=> vbajson16a Finished!" & vbCrLf
Exit Sub
    vbajson17
    Debug.Print "=> vbajson17 Finished!" & vbCrLf
    vbajson18
    Debug.Print "=> vbajson18 Finished!" & vbCrLf
    vbajson19
    Debug.Print "=> vbajson19 Finished!" & vbCrLf
    vbajson20
    Debug.Print "=> vbajson20 Finished!" & vbCrLf

End Sub

Private Sub vbajson1()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strJson As String

    Debug.Print "=> vbajson1"

    ' read the JSON into an object:
    strJson = "{ bla:""hi"", ""items"": [{""it"":1,""itx"":2},{""i3"":""x""}] }"
    Debug.Print , "strJson=" & strJson & " DOES NOT VALIDATE AT jsonlint.com"
    Debug.Print , "EXPECTING STRING"
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

    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.GetParseError
        Debug.Print , "FAILED"
        GoTo PROC_EXIT
    End If

    ' get data from arrays etc.:
    Debug.Print , "Bla: " & o.Item("bla") & " - Items of itx: " & _
        o.Item("items").Item(1).Item("itx")

PROC_EXIT:
    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson1a()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strJson As String

    Debug.Print "=> vbajson1a"

    ' read the JSON into an object:
    strJson = "{ bla:""hi"", ""here are some items"": [{""it"":1,""itx"":2},{""i3"":""x""}] }"
    Debug.Print , "strJson=" & strJson & " DOES NOT VALIDATE AT jsonlint.com"
    Debug.Print , "EXPECTING STRING"
    Set o = lib.parse(strJson)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.GetParseError
        Debug.Print , "FAILED"
        GoTo PROC_EXIT
    End If

    ' get data from arrays etc.:
    Debug.Print , "The Blah: " & o.Item("bla")
    Debug.Print , "The Item of itx: " & o.Item("here are some items").Item(1).Item("itx")

PROC_EXIT:
    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson2()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson2"

    Debug.Print , "vbajson2: This test will kill Excel!"
    Debug.Print , "NEEDS DEBUGGING AND ERROR HANDLING"
Exit Sub

    ' read the JSON into an object:
    Set o = lib.parse("{bla:'hi I'm a single quote!'"", items: [{it:1,itx:2},{i3:'x'}] }")
   
    ' get the parsed text back:
    Debug.Print lib.toString(o)

    ' get data from arrays etc.:
    Debug.Print "Bla: " & o.Item("bla") & " - Items: " & _
        o.Item("items").Item(1).Item("itx")

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson3()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson3"

    Debug.Print , "vbajson3: FIXED."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson4()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson4"

    Debug.Print , "vbajson4: FIXED. Testing needed for other locale."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson5()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson5"

    Debug.Print , "vbajson5: FIXED. Test case needed."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson6()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson6"

    strTest = "{""Cus:ip"":[123,456,789],""Da:te"":[1,2,3],""Close:Type"":[""stock"",""bo::nd"",""sto:::ck""]}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)
   
    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson7()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson7"

    strTest = "{""total_rows"":36778,""offset"":26220,""rows"":[" & _
                "{""id"":""6b80c0b76"",""key"":""a@bbb.net"",""value"":{""entryid"":""81151F241C2500"",""subject"":""test subject"",""senton"":""2009-7-09 22:03:43""}}," & _
                "{""id"":""b10ed9bee"",""key"":""b@bbb.net"",""value"":{""entryid"":A7C3CF74EA95C9F"",""subject"":""test subject2"",""senton"":""2009-4-21 10:18:26""}}]}"
    Debug.Print "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)
   
    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "FAILED"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson7b()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson7b"

    strTest = "{""Cusip"":[123,456,789],""Date"":[1,2,3],""CloseType"":[""stock"",""bond"",""stock""]}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)
   
    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson8()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson8"

    ' Create a 2-d array, such as:
    Dim arr(0 To 1, 0 To 1) As String
    arr(0, 0) = "a"
    arr(0, 1) = "b"
    arr(1, 0) = "c"
    arr(1, 1) = "d"

    ' Try to convert to JSON with
    ' Debug.Print lib.toString(arr)
    ' Type Mismatch ERROR raised here: toString = Replace(obj, ",", ".")
    
    Debug.Print , "vbajson8: FAILED. - Not supported in this version of VBA-JSON"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson8b()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson8b"

    Dim arr(0 To 3) As Variant
    arr(0) = "a"
    arr(1) = "b"
    arr(2) = "c"
    arr(3) = "d"

    Debug.Print , "lib.toString(arr)=" & lib.toString(arr)

    Debug.Print , "VALIDATED"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson8c()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson8c"

    Dim arr(1 To 2) As Variant
    arr(1) = Array("a", "b")
    arr(2) = Array("c", "d")

    Debug.Print , "lib.toString(arr)=" & lib.toString(arr)

    Debug.Print , "VALIDATED"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson9()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson9"

    Debug.Print , "vbajson9: CLOSED."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson10()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson10"

    strTest = "{""BigNumber1"":32769}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED"

    strTest = "{""BigNumber2"":1234567890}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED"

    strTest = "{""BigNumber3"":123456789012345678901}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED WITH ROUNDING"

    strTest = "{""BigNumber4"":1234567890123456789012345678901234567890}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED WITH e+39"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson10a()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson10a"

    strTest = "{""RealNumber1"":32.769}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED"

    strTest = "{""RealNumber2"":0.1234567890}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED"

    strTest = "{""RealNumber3"":1.23456789012345678901}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED WITH ROUNDING TO 16 DECIMAL PLACES"

    strTest = "{""RealNumber4"":-12345.67890123456789012345678901234567890}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    Debug.Print , "VALIDATED WITH ROUNDING TO 12 DECIMAL PLACES"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson11()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson11"

    strTest = "{""Path"":""C:\sample\sample.jpg""}"
    Set o = lib.parse(strTest)
    Debug.Print , "1. strTest=" & strTest
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED"
    Debug.Print

    strTest = "{""Path"":""C:\\sample\\sample.jpg""}"
    Debug.Print , "2. strTest=" & strTest
    Set o = lib.parse(strTest)
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED"
    Debug.Print

    strTest = "{""Path"":""C:\\\sample\\\sample.jpg""}"
    Debug.Print , "3. strTest=" & strTest
    Set o = lib.parse(strTest)
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED"
    Debug.Print

    strTest = "{""Path"":""C:\\\\sample\\\\sample.jpg""}"
    Debug.Print , "4. strTest=" & strTest
    Set o = lib.parse(strTest)
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED"
    Debug.Print

    strTest = "{""Path"":""C:\\\\\sample\\\\\sample.jpg""}"
    Debug.Print , "5. strTest=" & strTest
    Set o = lib.parse(strTest)
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED"
    Debug.Print

    strTest = "{""Path"":""C:\\\\\\sample\\\\\\sample.jpg""}"
    Debug.Print , "6. strTest=" & strTest
    Set o = lib.parse(strTest)
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson12()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson12"

    strTest = "{""ListsState"":{""MenuLocation"":[""Kelim"",""ChecklistTools""],""CurentLoadedChecklist"":""ToolsConfig"",""InnerDoc"":{""DapiotRegel"":{""ClassName"":""White"",""CHLTitle"":""???? ???"",""Fields"":{}},""ToolsConfig"":{""ClassName"":""White"",""CHLTitle"":""??????"",""Fields"":{""ToolsConfigHeliID"":""036"",""ToolsConfigCrewSize"":""3"",""ToolsConfigOperativeWgt"":""1,500"",""ToolsConfigNumOf669"":""0"",""ToolsConfigNumOf669Doc"":""0"",""ToolsConfigNumOf669Med"":""0"",""ToolsConfigNumOf669Equip"":""0"",""ToolsConfigNumOfSol"":""0"",""ToolsConfigNumOfPax"":""0"",""ToolsConfigCargo"":""0"",""ToolsConfigCar"":""0"",""ToolsConfigFuelExtTanks"":""0"",""ToolsConfigFuelTotal"":""0"",""ToolsConfigCarUnits_Save"":""?\\?""}}}}}"
    Set o = lib.parse(strTest)
    Debug.Print , "strTest=" & strTest
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED BUT INTERNATIONAL CHARACTERS NOT DISPLAYED - NEED APPROPRIATE LOCALE SETUP"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson13()

    Dim lib As New jsonlib
    Dim o As Object
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim str4 As String
    Dim str5 As String
    Dim str6 As String
    Dim str7 As String
    Dim str8 As String
    Dim strTest As String

    Debug.Print "=> vbajson13"

    str1 = "{""schedules"":[{""summary"":""Sign in"",""executedOn"":""10/Oct/12 1:50 PM"",""cycleName"":""asdf"",""cycleID"":15,""label"":""1, 2, 3, 4, 5"",""issueId"":123,""versionName"":""asdf"",""issueID"":123,""defects"":["
    Debug.Print , "str1=" & str1
    str2 = "{""key"":""124"",""status"":""Closed"",""summary"":""Title""},{""key"":""asdf"",""status"":""Closed"",""summary"":""asdfasdf""}],""executedByDisplay"":""Name of person"",""executionStatus"":""2"",""htmlComment"":""asdfasd"",""projectID"":""asdf"",""executedBy"":""asdasg"",""component"":"""",""versionID"":""adasd"",""issueKey"":""asdf"",""scheduleID"":73,""comment"":""adsfasdf""},"
    Debug.Print , "str2=" & str2
    str3 = "{""summary"":""asdf"",""executedOn"":""10/Oct/12 1:17 PM"",""cycleName"":""asdf"",""cycleID"":15,""label"":""1, 2, 3, 4, 5, 6, 7, 89, 5, 34"",""issueId"":10012,""versionName"":""sdf"",""issueID"":10012,""defects"":["
    Debug.Print , "str3=" & str3
    str4 = "{""key"":""asdf"",""asdf"":""asdf"",""summary"":""asdf""},{""key"":""asdf"",""status"":""Closed"",""summary"":""asdf""}],""executedByDisplay"":""asdf"",""executionStatus"":""2"",""htmlComment"":""asdf"",""projectID"":10002,""executedBy"":""asdf"",""component"":"""",""versionID"":10001,""issueKey"":""Edf"",""scheduleID"":18,""comment"":""asdf""},"
    Debug.Print , "str4=" & str4
    str5 = "{""summary"":""asdf"",""executedOn"":""10/Oct/12 1:20 PM"",""cycleName"":""asdf"",""cycleID"":15,""label"":""1, 2"",""issueId"":10011,""versionName"":""asdf"",""issueID"":10011,""defects"":["
    Debug.Print , "str5=" & str5
    str6 = "{""key"":""asdf"",""status"":""Closed"",""summary"":""asdf""},{""key"":""asdf"",""status"":""Closed"",""summary"":""asdf - asdf""}],""executedByDisplay"":""asdf"",""executionStatus"":""2"",""htmlComment"":""asdf"",""projectID"":10002,""executedBy"":""asdf"",""component"":"""",""versionID"":10001,""issueKey"":""asdf"",""scheduleID"":17,""comment"":""asdf""},"
    Debug.Print , "str6=" & str6
    str7 = "{""summary"":""asdfasdf"",""executedOn"":""10/Oct/12 1:26 PM"",""cycleName"":""asdf"",""cycleID"":15,""label"":""1,2"",""issueId"":10010,""versionName"":""asdf"",""issueID"":10010,""defects"":["
    Debug.Print , "str7=" & str7
    str8 = "{""key"":""asdf"",""status"":""Closed"",""summary"":""asdfa""},{""key"":""asdf"",""status"":""Closed"",""summary"":""asdf""}],""executedByDisplay"":""asdfasf"",""executionStatus"":""2"",""htmlComment"":""asdfafd"",""projectID"":10002,""executedBy"":""asdf"",""component"":"""",""versionID"":10001,""issueKey"":""afgaf"",""scheduleID"":16,""comment"":""asdf""}]}"
    Debug.Print , "str8=" & str8

    strTest = str1 & str2 & str3 & str4 & str5 & str6 & str7 & str8

    Set o = lib.parse(strTest)
    Debug.Print , "strTest=" & strTest
    Debug.Assert Err.Number = 0
    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    Debug.Print , "VALIDATED - WATCH OUT FOR LINE WRAP WITH C&P TO JSONLint"

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson13b()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson13b"

    Debug.Print , "vbajson13b: Test case needed. String Builder Class and Office x64 - TBD."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson14()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strEmbed As String

    Debug.Print "=> vbajson14"
    strEmbed = "[{""ty:pe"":""t1"",""title"":""データ1"",""attr"":[""1-1"",""1-2""]},{""type"":""t2"",""title"":""データ2"",""attr"":[""2-1"",""2-2""]}]"""
    Debug.Print , "strEmbed=" & strEmbed

    Set o = lib.parse(strEmbed)

    Debug.Print , "lib.toString(o)=" & lib.toString(o)
    'Debug.Assert Err.Number = 0
    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.GetParseError
        Debug.Print , "FAILED"
    End If

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson15()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson15"

    Debug.Print , "vbajson15: Test case needed."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson16()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson16"
    Debug.Print , "m_SDecimal= " & GetSDecimal
    Debug.Print , "m_SThousand= " & GetSThousand

    strTest = "{""InternationalNumber1"":32769.05}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    If lib.GetParseError <> vbNullString Then
        Debug.Print , "lib.GetParserError=" & lib.GetParseError
        Debug.Print , "FAILED"
        GoTo PROC_EXIT
    End If

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    End If

PROC_EXIT:
    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson16a()

    Dim lib As New jsonlib
    Dim o As Object
    Dim strTest As String

    Debug.Print "=> vbajson16a"
    Debug.Print , "m_SDecimal= " & GetSDecimal
    Debug.Print , "m_SThousand= " & GetSThousand

    strTest = "{""InternationalNumber2"":-1234567.89}"
    Debug.Print , "strTest=" & strTest
    ' read the JSON into an object:
    Set o = lib.parse(strTest)

    If lib.GetParseError <> vbNullString Then
        Debug.Print , "lib.GetParseError=" & lib.GetParseError
        Debug.Print , "FAILED"
        GoTo PROC_EXIT
    End If

    ' get the parsed text back:
    Debug.Print , "lib.toString(o)=" & lib.toString(o)

    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    End If

PROC_EXIT:
    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson17()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson17"

    Debug.Print , "vbajson17: Test case needed."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson18()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson18"

    Debug.Print , "vbajson18: Test case needed."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson19()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson19"

    Debug.Print , "vbajson19: Test case needed."

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson20()

    Dim lib As New jsonlib
    Dim o As Object

    Debug.Print "=> vbajson20"

    Debug.Print , "vbajson20: Test case needed."

    Set o = Nothing
    Set lib = Nothing

End Sub




