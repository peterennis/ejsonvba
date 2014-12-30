Attribute VB_Name = "basTESTejsonvbaErrors"
Option Explicit
Option Compare Text
Option Private Module

Public Sub RunAllejsonvbaErrorTests()

Debug.Print "=> Bypass RunAllejsonvbaErrorTests!" & vbCrLf
Exit Sub
'GoTo TEST:
    vbajson1_fail
    Debug.Print "=> vbajson1_fail Finished!" & vbCrLf
    vbajson7_fail
    Debug.Print "=> vbajson7_fail Finished!" & vbCrLf
    vbajson8_fail
    Debug.Print "=> vbajson8_fail Finished!" & vbCrLf
    parse_error_001
    Debug.Print "=> parse_error_001 Finished!" & vbCrLf
    parse_error_002
    Debug.Print "=> parse_error_002 Finished!" & vbCrLf
    parse_error_003
    Debug.Print "=> parse_error_003 Finished!" & vbCrLf
    parse_test3_fail
    Debug.Print "=> parse_test3_fail Finished!" & vbCrLf
    parse_error_004
    Debug.Print "=> parse_error_004 Finished!" & vbCrLf
    parse_error_005
    Debug.Print "=> parse_error_005 Finished!" & vbCrLf
TEST:
'    parse_error_006
'    Debug.Print "=> parse_error_006 Finished!" & vbCrLf
'    parse_error_007
'    Debug.Print "=> parse_error_007 Finished!" & vbCrLf
'    parse_error_008
'    Debug.Print "=> parse_error_008 Finished!" & vbCrLf
'    parse_error_009
'    Debug.Print "=> parse_error_009 Finished!" & vbCrLf
'    parse_error_010
'    Debug.Print "=> parse_error_010 Finished!" & vbCrLf

End Sub

Private Sub vbajson1_fail()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim o As Object
    Dim strJson As String

    Debug.Print "=> vbajson1_fail"

    ' read the JSON into an object:
    strJson = "{ bla:""hi"", ""items"": [{""it"":1,""itx"":2},{""i3"":""x""}] }"
    Debug.Print , "strJson=" & strJson & " DOES NOT VALIDATE AT jsonlint.com"
    Debug.Print , "EXPECTING STRING"
    
    'lib.DebugState = True
    Set o = lib.Parse(strJson)

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
    Debug.Print , "lib.ToString(o)=" & lib.ToString(o)

    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
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

Private Sub vbajson7_fail()

    Debug.Print "=> vbajson7_fail"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim o As Object
    Dim strTest As String

    strTest = "{""total_rows"":36778,""offset"":26220,""rows"":[" & _
                "{""id"":""6b80c0b76"",""key"":""a@bbb.net"",""value"":{""entryid"":""81151F241C2500"",""subject"":""test subject"",""senton"":""2009-7-09 22:03:43""}}," & _
                "{""id"":""b10ed9bee"",""key"":""b@bbb.net"",""value"":{""entryid"":A7C3CF74EA95C9F"",""subject"":""test subject2"",""senton"":""2009-4-21 10:18:26""}}]}"
    Debug.Print "strTest=" & strTest

    ' read the JSON into an object:
    'lib.DebugState = True
    Set o = lib.Parse(strTest)
   
    ' get the parsed text back:
    Debug.Print , "lib.ToString(o)=" & lib.ToString(o)

    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED"
    End If

    Set o = Nothing
    Set lib = Nothing

End Sub

Private Sub vbajson8_fail()

    Debug.Print "=> vbajson8_fail"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim o As Object

    On Error GoTo PROC_ERR

    ' Create a 2-d array, such as:
    Dim arr(0 To 1, 0 To 1) As String
    arr(0, 0) = "a"
    arr(0, 1) = "b"
    arr(1, 0) = "c"
    arr(1, 1) = "d"

    ' Try to convert to JSON with
    Debug.Print , "lib.ToString(arr)=" & lib.ToString(arr)
    ' Run-time error 13 Type Mismatch ERROR raised here: ToString = Replace(obj, ",", ".")
    
    Debug.Print , "FAILED. - This array definition is not supported in the current version."

PROC_EXIT:
    Set o = Nothing
    Set lib = Nothing
    Exit Sub

PROC_ERR:
    If Err = 13 Then
        Resume Next
    Else
        Stop
    End If

End Sub

Private Sub parse_error_001()

    Debug.Print "=> parse_error_001"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    parseString = " " & vbCrLf & vbTab & " {"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED {}"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_002()

    Debug.Print "=> parse_error_002"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    parseString = " " & vbCrLf & vbTab & " ["
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED []"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_003()

    Debug.Print "=> parse_error_003"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    parseString = " " & vbCrLf & vbTab & " <"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_test3_fail()

    Dim lib As New ejsonlib
    Dim json As Object
    Dim strEmbed As String
    Dim errString As String

    Debug.Print "=> parse_test3_fail"

    strEmbed = " [[], {""test1"":'v1', 'test2':'v222', test3:""v33333""}, null , ""test"", 123, 567.8910, 4.7e+10, true,  false]"
    Debug.Print , "strEmbed=" & strEmbed

    'lib.DebugState = True
    Set json = lib.Parse(" " & vbCrLf & vbTab & strEmbed)

    Debug.Print , "lib.ToString(json)=" & lib.ToString(json)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_004()

    Debug.Print "=> parse_error_004"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    parseString = "{" & "Bug" & "}"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_005()

    Debug.Print "=> parse_error_005"

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    parseString = "{" & """Bug" & "}"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

