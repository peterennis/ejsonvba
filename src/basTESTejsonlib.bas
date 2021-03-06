Attribute VB_Name = "basTESTejsonlib"
Option Explicit
Option Compare Text
Option Private Module

Public Sub RunAllejsonlibTests()

Debug.Print "=> Bypass RunAllejsonlibTests!" & vbCrLf
Exit Sub
    ToString_test1
    Debug.Print "=> ToString_test1 Finished!" & vbCrLf
    ToString_test2
    Debug.Print "=> ToString_test2 Finished!" & vbCrLf
    parse_test1
    Debug.Print "=> parse_test1 Finished!" & vbCrLf
    parse_test2
    Debug.Print "=> parse_test2 Finished!" & vbCrLf
    parse_test3
    Debug.Print "=> parse_test3 Finished!" & vbCrLf
    parse_test3a
    Debug.Print "=> parse_test3a Finished!" & vbCrLf
    parse_test4
    Debug.Print "=> parse_test4 Finished!" & vbCrLf
    parse_test5
    Debug.Print "=> parse_test5 Finished!" & vbCrLf
    skip_test
    Debug.Print "=> skip_test Finished!" & vbCrLf

End Sub

'
'   ejsonlib.ToString tests
'
Private Sub ToString_test1()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim a As String
    Dim b As Date

    Debug.Print "=> ToString_test1"

    b = Now()

    Debug.Print , "ToString_test1=" & lib.ToString(Array("a", "b", Array(1, b, "3")))
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED"
    End If

    Set lib = Nothing

End Sub

Private Sub ToString_test2()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim a As Object
    Dim b As Object
    Dim c As New Collection

    Debug.Print "=> ToString_test2"

    Set a = CreateObject("Scripting.Dictionary")
    Set b = CreateObject("Scripting.Dictionary")

    a("aaa") = "abc"
    a("bbb") = Array(0, 1, b)
    b("ccc") = "def"
    Set b("ddd") = c
    c.Add "ghi"
    c.Add 999

    Debug.Print , "lib.ToString(a)=" & lib.ToString(a)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED"
    End If

    Set lib = Nothing
    Set c = Nothing
    Set b = Nothing
    Set a = Nothing

End Sub

'
'   ejsonlib.parse tests
'
Private Sub parse_test1()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    Debug.Print "=> parse_test1"
    parseString = " " & vbCrLf & vbTab & " {}"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
        Debug.Print , "FAILED {}"
        GoTo PROC_EXIT
    End If

    Debug.Assert TypeName(json) = "Dictionary"
    Debug.Print , "TypeName(json)=" & TypeName(json), "json.Count=" & json.Count
    Debug.Print
    
    parseString = " " & vbCrLf & vbTab & " []"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)
    If lib.ParseError = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.ParseError
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

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim parseString As String

    Debug.Print "=> parse_test2"
    parseString = " " & vbCrLf & vbTab & " {}"
    Debug.Print , "parseString=" & parseString

    'lib.DebugState = True
    Set json = lib.Parse(parseString)

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

Private Sub parse_test3()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim strEmbed As String
    Dim errString As String

    Debug.Print "=> parse_test3"

    strEmbed = " [[], {""test1"":'v1', 'test2':'v222', ""test3"":""v33333""}, null , ""test"", 123, 567.8910, 4.7e+10, true,  false]"
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

Private Sub parse_test3a()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim strEmbedValid As String
    Dim errString As String

    Debug.Print "=> parse_test3a STRICT JSON"

    strEmbedValid = " [[], {""test1"":""v1"", ""test2"":""v222"", ""test3"":""v33333""}, null , ""test"", 123, 567.8910, 4.7e+10, true,  false]"
    Debug.Print , "strEmbedValid=" & strEmbedValid

    Set json = lib.Parse(" " & vbCrLf & vbTab & strEmbedValid)

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

Private Sub parse_test4()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim json As Object
    Dim errString As String
    Dim strEmbed As String

    Debug.Print "=> parse_test4"
    strEmbed = "[{""type"":""t1"",""title"":""データ1"",""attr"":[""1-1"",""1-2""]},{""type"":""t2"",""title"":""データ2"",""attr"":[""2-1"",""2-2""]}]"""
    Debug.Print , "strEmbed=" & strEmbed

    Set json = lib.Parse("[{""type"":""t1"",""title"":""データ1"",""attr"":[""1-1"",""1-2""]},{""type"":""t2"",""title"":""データ2"",""attr"":[""2-1"",""2-2""]}]")

    Debug.Print , "lib.ToString(json)=" & lib.ToString(json)
    errString = lib.ParseError
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

    Dim lib As ejsonlib
    Set lib = New ejsonlib
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

    Debug.Print , "text=" & text

    Set json = lib.Parse(text)
    Debug.Assert Err.Number = 0
    res1 = lib.ToString(json)

    Set json = lib.Parse(lib.ToString(json))
    Debug.Assert Err.Number = 0
    res2 = lib.ToString(json)

    errString = lib.ParseError
    If errString = "" Then
        Debug.Print , res1
        Debug.Print , res2
        Debug.Assert (res1 = res2)
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , errString
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

'
'   internal function tests
'
'       before executing this test, modify ejsonlib.skipChar to 'Friend' or 'Public'
'
Private Sub skip_test()

    Dim lib As ejsonlib
    Set lib = New ejsonlib
    Dim str As String
    Dim index As Long
    Dim errString As String

    Debug.Print "=> skip_test"

    str = vbCrLf & vbCr & vbLf & " " & "abc"
    index = 1

    lib.SkipChar str, index
    Debug.Assert index = 6

    Debug.Print , "index=" & index, "Mid(str, index, 1)=" & Mid(str, index, 1)

    Set lib = Nothing

End Sub

