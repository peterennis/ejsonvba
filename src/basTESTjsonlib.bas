Attribute VB_Name = "basTESTjsonlib"
Option Explicit

Public Sub RunAlljsonlibTests()

'    toString_test1
'    Debug.Print "=> toString_test1 Finished!" & vbCrLf
'    toString_test2
'    Debug.Print "=> toString_test2 Finished!" & vbCrLf
'    parse_test1
'    Debug.Print "=> parse_test1 Finished!" & vbCrLf
'    parse_test2
'    Debug.Print "=> parse_test2 Finished!" & vbCrLf
    parse_test3
    Debug.Print "=> parse_test3 Finished!" & vbCrLf
Exit Sub
    parse_test4
    Debug.Print "=> parse_test4 Finished!" & vbCrLf
    parse_test5
    Debug.Print "=> parse_test5 Finished!" & vbCrLf
    skip_test
    Debug.Print "=> skip_test Finished!" & vbCrLf

End Sub

'
'   jsonlib.toString tests
'
Private Sub toString_test1()

    Dim a As String
    Dim b As Date
    Dim lib As New jsonlib
    Dim errString As String

    Debug.Print "=> toString_test1"

    b = Now()

    Debug.Print , "toString_test1=" & lib.toString(Array("a", "b", Array(1, b, "3")))
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED"
    End If

    Set lib = Nothing

End Sub

Private Sub toString_test2()

    Dim a As Object
    Dim b As Object
    Dim c As New Collection
    Dim lib As New jsonlib
    Dim errString As String

    Debug.Print "=> toString_test2"

    Set a = CreateObject("Scripting.Dictionary")
    Set b = CreateObject("Scripting.Dictionary")

    a("aaa") = "abc"
    a("bbb") = Array(0, 1, b)
    b("ccc") = "def"
    Set b("ddd") = c
    c.Add "ghi"
    c.Add 999

    Debug.Print , "lib.toString(a)=" & lib.toString(a)
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED"
    End If

    Set lib = Nothing
    Set c = Nothing
    Set b = Nothing
    Set a = Nothing

End Sub

'
'   jsonlib.parse tests
'
Private Sub parse_test1()

    Dim lib As New jsonlib
    Dim json As Object
    Dim errString As String
    Dim parseString As String

    Debug.Print "=> parse_test1"
    parseString = " " & vbCrLf & vbTab & " {}"
    Debug.Print , "parseString=" & parseString

    Set json = lib.parse(" " & vbCrLf & vbTab & " {}")
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED {}"
        GoTo PROC_EXIT
    End If
    Debug.Assert TypeName(json) = "Dictionary"
    Debug.Print , "TypeName(json)=" & TypeName(json), "json.Count=" & json.Count
    Debug.Print
    
    parseString = " " & vbCrLf & vbTab & " []"
    Debug.Print , "parseString=" & parseString

    Set json = lib.parse(" " & vbCrLf & vbTab & " []")
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED []"
        GoTo PROC_EXIT
    End If
    Debug.Assert TypeName(json) = "Collection"
    Debug.Print , "TypeName(json)=" & TypeName(json), "json.Count=" & json.Count

PROC_EXIT:
    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_test2()

    Dim lib As New jsonlib
    Dim json As Object
    Dim errString As String

    Debug.Print "=> parse_test2"

    Set json = lib.parse(" " & vbCrLf & vbTab & " {}")

    Debug.Print , "lib.toString(json)=" & lib.toString(json)
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_test3()

    Dim lib As New jsonlib
    Dim json As Object
    Dim strEmbed As String
    Dim strEmbedValid As String
    Dim errString As String

    Debug.Print "=> parse_test3"

    strEmbedValid = " [[], {""test1"":""v1"", ""test2"":""v222"", ""test3"":""v33333""}, null , ""test"", 123, 567.8910, 4.7e+10, true,  false]"
    Debug.Print , "strEmbedValid=" & strEmbedValid
    strEmbed = " [[], {""test1"":'v1', 'test2':'v222', test3:""v33333""}, null , ""test"", 123, 567.8910, 4.7e+10, true,  false]"
    Debug.Print , "strEmbed=" & strEmbed

    Set json = lib.parse(" " & vbCrLf & vbTab & strEmbed)

    Debug.Print , "lib.toString(json)=" & lib.toString(json)
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_test4()

    Dim lib As New jsonlib
    Dim json As Object
    Dim errString As String

    Debug.Print "=> parse_test4"

    Set json = lib.parse("[{""type"":""t1"",""title"":""データ1"",""attr"":[""1-1"",""1-2""]},{""type"":""t2"",""title"":""データ2"",""attr"":[""2-1"",""2-2""]}]")

    Debug.Print lib.toString(json)
    'Debug.Assert Err.Number = 0
    errString = lib.GetParserErrors
    If errString = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_test5()

    Dim lib As New jsonlib
    Dim json As Object
    Dim text As String
    Dim res1 As String
    Dim res2 As String
    Dim errString As String

    Debug.Print "=> parse_test5"

    With CreateObject("ADODB.Stream")
        .Open
        .Charset = "UTF-8"
        .LoadFromFile ActiveWorkbook.Path & "\test\test1.json"
        text = .ReadText(-1)
        .Close
    End With

    Debug.Print text

    Set json = lib.parse(text)
    Debug.Assert Err.Number = 0
    res1 = lib.toString(json)

    Set json = lib.parse(lib.toString(json))
    Debug.Assert Err.Number = 0
    res2 = lib.toString(json)

    Debug.Print res1
    Debug.Print res2

    Debug.Assert (res1 = res2)

    Set json = Nothing
    Set lib = Nothing

End Sub

'
'   internal function tests
'
'       before executing this test, modify jsonlib.skipChar to 'Friend' or 'Public'
'
Private Sub skip_test()

    Dim lib As New jsonlib
    Dim str As String
    Dim index As Long
    Dim errString As String

    Debug.Print "=> skip_test"

    str = vbCrLf & vbCr & vbLf & " " & "abc"
    index = 1

    lib.skipChar str, index
    Debug.Assert index = 6

    Debug.Print index, Mid(str, index, 1)

    Set lib = Nothing

End Sub

