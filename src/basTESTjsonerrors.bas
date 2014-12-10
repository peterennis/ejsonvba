Attribute VB_Name = "basTESTjsonerrors"
Option Explicit
Option Compare Text
Option Private Module

Public Sub RunAllvbajsonErrorTests()

GoTo TEST:
    parse_error_001
    Debug.Print "=> parse_error_001 Finished!" & vbCrLf
    parse_error_002
    Debug.Print "=> parse_error_002 Finished!" & vbCrLf
    parse_error_003
    Debug.Print "=> parse_error_003 Finished!" & vbCrLf
TEST:
    parse_error_004
    Debug.Print "=> parse_error_004 Finished!" & vbCrLf
'    parse_error_005
'    Debug.Print "=> parse_error_005 Finished!" & vbCrLf
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

Private Sub parse_error_001()

    Debug.Print "=> parse_error_001"

    Dim lib As jsonlib
    Set lib = New jsonlib
    Dim json As Object
    Dim parseString As String

    parseString = " " & vbCrLf & vbTab & " {"
    Debug.Print , "parseString=" & parseString

    lib.DebugState = True
    Set json = lib.parse(parseString)
    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.GetParseError
        Debug.Print , "FAILED {}"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_002()

    Debug.Print "=> parse_error_002"

    Dim lib As jsonlib
    Set lib = New jsonlib
    Dim json As Object
    Dim parseString As String

    parseString = " " & vbCrLf & vbTab & " ["
    Debug.Print , "parseString=" & parseString

    lib.DebugState = True
    Set json = lib.parse(parseString)
    If lib.GetParseError = "" Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , lib.GetParseError
        Debug.Print , "FAILED []"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_003()

    Debug.Print "=> parse_error_003"

    Dim lib As jsonlib
    Set lib = New jsonlib
    Dim json As Object
    Dim parseString As String

    parseString = " " & vbCrLf & vbTab & " <"
    Debug.Print , "parseString=" & parseString

    lib.DebugState = True
    Set json = lib.parse(parseString)
    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

Private Sub parse_error_004()

    Debug.Print "=> parse_error_004"

    Dim lib As jsonlib
    Set lib = New jsonlib
    Dim json As Object
    Dim parseString As String

    parseString = "{" & "Bug" & "}"
    Debug.Print , "parseString=" & parseString

    lib.DebugState = True
    Set json = lib.parse(parseString)
    If lib.GetParseError = vbNullString Then
        Debug.Print , "VALIDATED"
    Else
        Debug.Print , "FAILED"
    End If

    Set json = Nothing
    Set lib = Nothing

End Sub

