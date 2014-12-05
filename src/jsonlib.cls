VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "jsonlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const INVALID_RPC_CALL  As Long = 7

Private Const jsonlibVERSION As String = "0.1.4"
Private Const jsonlibVERSION_DATE As String = "December 4, 2014"

Private pParseError As String
Private m_SDecimal As String
Private m_SThousand As String
Private m_StrLen As Long
Private m_Str() As Integer
'

Private Sub Class_Initialize()
    Debug.Print , "Class_Initialize"
    'Debug.Print , "jsonlibVERSION=" & jsonlibVERSION
    'Debug.Print , "jsonlibVERSION_DATE=" & jsonlibVERSION_DATE
End Sub

Private Sub Class_Terminate()
    Debug.Print , "Class_Terminate"
End Sub

Public Property Get GetParseError() As String
    GetParseError = pParseError
End Property

Public Property Get ClearParseError() As String
   pParseError = vbNullString
   ClearParseError = "pParseError = vbNullString"
End Property

'
'   parse string and create JSON object (Dictionary or Collection in VB)
'
Public Function parse(ByRef str As String) As Object

    m_SDecimal = GetSDecimal
    m_SThousand = GetSThousand

    Dim index As Long
    index = 1

    'Call GenerateStringArray(str)

    pParseError = vbNullString
    On Error Resume Next

    Call skipChar(str, index)

    Select Case Mid(str, index, 1)
    Case "{"
        Set parse = parseObject(str, index)
    Case "["
        Set parse = parseArray(str, index)
    Case Else
        pParseError = "str=" & str & " index=" & index & " Invalid JSON"
        Debug.Print , "E1 parse: " & pParseError
        Exit Function
    End Select

End Function

Private Sub GenerateStringArray(ByRef str As String)

    Dim i As Long

    m_StrLen = Len(str)
    ReDim m_Str(1 To m_StrLen)

    For i = 1 To m_StrLen
        m_Str(i) = AscW(Mid$(str, i, 1))
    Next i

End Sub

'
'   parse collection of key/value
'
Private Function parseObject(ByRef str As String, ByRef index As Long) As Object

    Set parseObject = CreateObject("Scripting.Dictionary")

    Dim sKey As String

    ' "{"
    Call skipChar(str, index)

    If Mid(str, index, 1) <> "{" Then
        pParseError = pParseError & "Invalid Object at position " & index & " : " & Mid(str, index)
        Debug.Print , "E2 parseObject: " & pParseError
        Exit Function
    End If

    index = index + 1

    Do

        Call skipChar(str, index)

        If "}" = Mid(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid(str, index, 1) Then
            index = index + 1
            Call skipChar(str, index)
        ElseIf index > Len(str) Then
            pParseError = pParseError & "Missing '}' at index " & index
            Debug.Print , "E3 parseObject: " & pParseError
            Exit Function
            Exit Do
        End If

        ' add key/value pair
        sKey = parseKey(str, index)
        On Error Resume Next
        
        parseObject.Add sKey, parseValue(str, index)
        If Err.Number <> 0 Then
            pParseError = pParseError & Err.Description & ": " & sKey & vbCrLf
            Debug.Print , "E4 parseObject: " & pParseError
            Exit Do
        End If

    Loop

End Function

'
'   parse list
'
Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection

    Set parseArray = New Collection

    ' "["
    Call skipChar(str, index)

    If Mid(str, index, 1) <> "[" Then
        pParseError = pParseError & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
        Debug.Print , "E5 parseArray: " & pParseError
        Exit Function
    End If

    index = index + 1

    Do

        Call skipChar(str, index)

        If "]" = Mid(str, index, 1) Then
            index = index + 1
            Exit Do
        ElseIf "," = Mid(str, index, 1) Then
            index = index + 1
            Call skipChar(str, index)
        ElseIf index > Len(str) Then
            pParseError = pParseError & "Missing ']' at index " & index
            Debug.Print , "E6 parseArray: " & pParseError
            Exit Function
        End If

        ' add value
        On Error Resume Next
        parseArray.Add parseValue(str, index)
        If Err.Number <> 0 Then
            pParseError = pParseError & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
            Debug.Print , "E7 parseArray: " & pParseError
            Exit Function
        End If

    Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef index As Long)

    Call skipChar(str, index)

    Select Case Mid(str, index, 1)
        Case "{"
            Set parseValue = parseObject(str, index)
        Case "["
            Set parseValue = parseArray(str, index)
        Case """", "'"
            parseValue = parseString(str, index)
        Case "t", "f"
            parseValue = parseBoolean(str, index)
        Case "n"
            parseValue = parseNull(str, index)
        Case Else
            parseValue = parseNumber(str, index)
    End Select

End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef index As Long) As String

    Dim quote As String
    Dim char As String
    Dim code As String

    Call skipChar(str, index)

    quote = Mid(str, index, 1)
    'Debug.Print , "quote=" & quote, index
    index = index + 1

    Do While index > 0 And index <= Len(str)
        char = Mid(str, index, 1)
        Select Case (char)
        Case "\"
            index = index + 1
            char = Mid(str, index, 1)
            Select Case (char)
                Case "\"
                    parseString = parseString & "\"
                    index = index + 1
                Case """", "/", "'", "\\"
                    parseString = parseString & char
                    index = index + 1
                Case "b"
                    parseString = parseString & vbBack
                    index = index + 1
                Case "f"
                    parseString = parseString & vbFormFeed
                    index = index + 1
                Case "n"
                    parseString = parseString & vbLf
                    index = index + 1
                Case "r"
                    parseString = parseString & vbCr
                    index = index + 1
                Case "t"
                    parseString = parseString & vbTab
                    index = index + 1
                Case "u"
                    index = index + 1
                    code = Mid(str, index, 4)
                    parseString = parseString & ChrW(Val("&h" + code))
                    index = index + 4
            End Select
        Case quote
            index = index + 1
            Exit Function
        Case Else
            parseString = parseString & char
            index = index + 1
        End Select
    Loop

End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef index As Long)

    Dim myValue As String
    Dim char As String

    'Debug.Print "parseNumber", str

    Call skipChar(str, index)

    'Debug.Print "m_StrLen=" & m_StrLen
    'Debug.Print "Len(str)=" & Len(str)
    
    Do While index > 0 And index <= Len(str)
        char = Mid(str, index, 1)
        If InStr("+-0123456789.eE", char) Then
            myValue = myValue & char
            'Debug.Print myValue
            index = index + 1
            If index > 200 Then
                Debug.Print "STOP"
                Stop
            End If
        Else
'            'check what is the grouping separator
'            If Not m_SDecimal = "." Then
'                myValue = Replace(myValue, ".", m_SDecimal)
'                Debug.Print "SD", myValue
'            End If
'
'            If m_SThousand = "." Then
'                myValue = Replace(myValue, ".", m_SDecimal)
'                Debug.Print "ST", myValue
'            End If
'
            parseNumber = CDec(myValue)
            Exit Function
        End If
    Loop

End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean

    Call skipChar(str, index)

    If Mid(str, index, 4) = "true" Then
        parseBoolean = True
        index = index + 4
    ElseIf Mid(str, index, 5) = "false" Then
        parseBoolean = False
        index = index + 5
    Else
        pParseError = pParseError & Err.Description & " - Invalid Boolean at position " & index & _
                        " : " & Mid(str, index) & vbCrLf
        Debug.Print , "E8 parseBoolean: " & pParseError
        Exit Function
    End If

End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef index As Long)

    Call skipChar(str, index)

    If Mid(str, index, 4) = "null" Then
        parseNull = Null
        index = index + 4
    Else
        pParseError = pParseError & Err.Description & " - Invalid Null value at position " & index & _
                        " : " & Mid(str, index) & vbCrLf
        Debug.Print , "E9 parseNull: " & pParseError
        Exit Function
    End If

End Function

Private Function parseKey(ByRef str As String, ByRef index As Long) As String

    Dim dquote As Boolean
    Dim squote As Boolean
    Dim char As String

    Call skipChar(str, index)

    Do While index > 0 And index <= Len(str)
        char = Mid(str, index, 1)
        Select Case (char)
            Case """"
                dquote = Not dquote
                index = index + 1
                If Not dquote Then
                    Call skipChar(str, index)
                    If Mid(str, index, 1) <> ":" Then
                        pParseError = pParseError & Err.Description & " - Invalid Key value at position " & index & _
                                    " : " & parseKey & vbCrLf
                    Debug.Print , "E10 parseKey: " & pParseError
                    Exit Function
                    End If
                End If
            Case "'"
                squote = Not squote
                index = index + 1
                If Not squote Then
                    Call skipChar(str, index)
                    If Mid(str, index, 1) <> ":" Then
                        pParseError = pParseError & Err.Description & " - Invalid Key value at position " & index & _
                                    " : " & parseKey & vbCrLf
                        Debug.Print , "E11 parseKey: " & pParseError
                        Exit Function
                    End If
                End If
            Case ":"
'                If Not dquote And Not squote Then
'                    index = index + 1
'                    Exit Do
'                ElseIf dquote And Not squote Then
'                    parseKey = parseKey & char
'                    index = index + 1
'                End If

                index = index + 1
                If Not dquote And Not squote Then
                    Exit Do
                ElseIf dquote And Not squote Then
                    parseKey = parseKey & char
                    index = index + 1
                'End If
                Else
                    parseKey = parseKey & char
                End If
            Case Else
                If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", char) Then
                Else
                    parseKey = parseKey & char
                End If
                index = index + 1
        End Select
    Loop

End Function

'
'   skip special character
'
'Private Sub skipChar(ByRef str As String, ByRef index As Long)
'Public Sub skipChar(ByRef str As String, ByRef index As Long)
Friend Sub skipChar(ByRef str As String, ByRef index As Long)

'    While index > 0 And index <= Len(str) And InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Mid(str, index, 1))
'        index = index + 1
'    Wend

    Dim bComment As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment As Boolean
    Do While index > 0 And index <= Len(str)
        Select Case Mid(str, index, 1)
            Case vbCr, vbLf
                If Not bLongComment Then
                    bStartComment = False
                    bComment = False
                End If
         
            Case vbTab, " ", "(", ")"

            Case "/"
                If Not bLongComment Then
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                    Else
                        bStartComment = True
                        bComment = False
                        bLongComment = False
                    End If
                Else
                    If bStartComment Then
                        bLongComment = False
                        bStartComment = False
                        bComment = False
                End If
            End If

        Case "*"
            If bStartComment Then
                bStartComment = False
                bComment = True
                bLongComment = True
            Else
                bStartComment = True
            End If

        Case Else
            If Not bComment Then
                Exit Do
            End If
        End Select

        index = index + 1
    Loop

End Sub

Public Function toString(ByRef obj As Variant) As String

    Select Case VarType(obj)
        Case vbNull
            toString = "null"
        Case vbDate
            toString = """" & CStr(obj) & """"
        Case vbString
            toString = """" & encode(obj) & """"
        Case vbObject
            Dim bFI, i
            bFI = True
            If TypeName(obj) = "Dictionary" Then
                toString = toString & "{"
                Dim keys
                keys = obj.keys
                For i = 0 To obj.Count - 1
                    If bFI Then bFI = False Else toString = toString & ","
                    Dim key
                    key = keys(i)
                    toString = toString & """" & key & """:" & toString(obj(key))
                Next i
                toString = toString & "}"
            ElseIf TypeName(obj) = "Collection" Then
                toString = toString & "["
                Dim myValue
                For Each myValue In obj
                    If bFI Then bFI = False Else toString = toString & ","
                    toString = toString & toString(myValue)
                Next myValue
                toString = toString & "]"
            End If
        Case vbBoolean
            If obj Then toString = "true" Else toString = "false"
        Case vbVariant, vbArray, vbArray + vbVariant
            Dim sEB
            toString = multiArray(obj, 1, "", sEB)
        Case Else
            toString = Replace(obj, ",", ".")
    End Select

End Function

Private Function encode(str) As String

    Dim i As Integer
    Dim j As Integer
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c As String
    Dim p As Boolean

    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To Len(str)
        p = True
        c = Mid(str, i, 1)
        For j = 0 To 7
            If c = Chr(aL1(j)) Then
                encode = encode & "\" & Chr(aL2(j))
                p = False
                Exit For
            End If
        Next

        If p Then
            Dim a
            a = AscW(c)
            If a > 31 And a < 127 Then
                encode = encode & c
            ElseIf a > -1 Or a < 65535 Then
                encode = encode & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
            End If
        End If
    Next

End Function

Private Function multiArray(aBD, iBC, sPS, ByRef sPT)   ' Array BoDy, Integer BaseCount, String PoSition

    Dim iDU As Integer  ' Integer DimensionUBound
    Dim iDL As Integer  ' Integer DimensionLBound
    Dim i As Integer

    On Error Resume Next
    iDL = LBound(aBD, iBC)
    iDU = UBound(aBD, iBC)

    Dim sPB1 As String  ' String PointBuffer1
    Dim sPB2 As String  ' String PointBuffer2

    If Err.Number = 9 Then
        sPB1 = sPT & sPS
        For i = 1 To Len(sPB1)
            If i <> 1 Then sPB2 = sPB2 & ","
            sPB2 = sPB2 & Mid(sPB1, i, 1)
        Next
'        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
        multiArray = multiArray & toString(aBD(sPB2))
    Else
        sPT = sPT & sPS
        multiArray = multiArray & "["
        For i = iDL To iDU
            multiArray = multiArray & multiArray(aBD, iBC + 1, i, sPT)
            If i < iDU Then multiArray = multiArray & ","
        Next
        multiArray = multiArray & "]"
        sPT = Left(sPT, iBC - 2)
    End If
    Err.Clear

End Function

Private Function JsonRpcCall(url As String, methName As String, args(), Optional user As String, Optional pwd As String) As Object

    Dim r As Object
    Dim cli As Object
    Dim pText As String
    Static reqId As Integer

    reqId = reqId + 1

    Set r = CreateObject("Scripting.Dictionary")
    r("jsonrpc") = "2.0"
    r("method") = methName
    r("params") = args
    r("id") = reqId

    pText = toString(r)

    Set cli = CreateObject("MSXML2.XMLHTTP.6.0")
    ' Set cli = New MSXML2.XMLHTTP60
    If Len(user) > 0 Then   ' If Not IsMissing(user) Then
        cli.Open "POST", url, False, user, pwd
    Else
        cli.Open "POST", url, False
    End If
    cli.setRequestHeader "Content-Type", "application/json"
    cli.Send pText

    If cli.Status <> 200 Then
        Err.Raise vbObjectError + INVALID_RPC_CALL + cli.Status, , cli.statusText
    End If

    Set r = parse(cli.responseText)
    Set cli = Nothing

    If r("id") <> reqId Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response id"

    If r.Exists("error") Or Not r.Exists("result") Then
        Err.Raise vbObjectError + INVALID_RPC_CALL, , "Json-Rpc Response error: " & r("error")("message")
    End If

    If Not r.Exists("result") Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response, missing result"

    Set JsonRpcCall = r("result")

End Function

Public Function toUnicode(str As String) As String

'    Dim x As Long
'    Dim uStr As New cStringBuilder
'    Dim uChrCode As Integer
'
'    For x = 1 To Len(str)
'        uChrCode = Asc(Mid(str, x, 1))
'        Select Case uChrCode
'            Case 8:   ' backspace
'                uStr.Append "\b"
'            Case 9: ' tab
'                uStr.Append "\t"
'            Case 10:  ' line feed
'                uStr.Append "\n"
'            Case 12:  ' formfeed
'                uStr.Append "\f"
'            Case 13: ' carriage return
'                uStr.Append "\r"
'            Case 34: ' quote
'                uStr.Append "\"""
'            Case 39:  ' apostrophe
'                uStr.Append "\'"
'            Case 92: ' backslash
'                uStr.Append "\\"
'            Case 123, 125:  ' "{" and "}"
'                uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
'            Case Is < 32, Is > 127: ' non-ascii characters
'                uStr.Append ("\u" & Right("0000" & Hex(uChrCode), 4))
'            Case Else
'                uStr.Append Chr$(uChrCode)
'        End Select
'    Next
'    toUnicode = uStr.toString
'    Exit Function

End Function
