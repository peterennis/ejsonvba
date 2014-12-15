Attribute VB_Name = "basTESTvbajsonlog"
Option Explicit
Option Compare Text
Option Private Module

' Fork of vba-json project
' Author: Peter Ennis
' GitHub: https://github.com/peterennis/eJsonVBA/tree/vba-json
' Date : Nov 9, 2014

' Original source from here: https://code.google.com/p/vba-json/source/detail?r=2
' Author: ryoyoko
' Date: Feb 14, 2009
' New BSD License : http://opensource.org/licenses/BSD-3-Clause

' Change integration:
' Add the file JSON.bas from Michael Glaser, review and integrate changes as appropriate
' VBJSON (http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html)
' is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

'"ID","Type","Status","Priority","Milestone","Owner","Summary","AllLabels","Link"
'"vbajson1","Defect","FIXED","Medium","","","outcome","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=1
    ' Reported by tkleinmi...@fenl.nl, Mar 19, 2009
    ' How can i read a parsed JSON string as an array?
    ' ANSWER
    ' -----------
    ' See test vbajson1 and results
'"vbajson2","Defect","New","HIGH","","","parseString bug","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=2
    ' Reported by webmas...@ediy.co.nz, Mar 24, 2009
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
'"vbajson4","Defect","REOPENED","Medium","","","improve parseNumber() for other decimal settings","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=4
    ' Reported by telmo.ca...@gmail.com, Jun 12, 2009
    ' I have added to parseNumber():
    '        If InStr(myValue, ".") Or InStr(myValue, "e") Or InStr(myValue, "E") Then
    '            ' for PT Local Settings where decimal is ","
    '            If CStr(1.2) = "1,2" Then myValue = Replace(myValue, ".", ",", 1, -1, 1)
    '            parseNumber = CDbl(myValue)
    '        Else
    '            parseNumber = CInt(myValue)
'"vbajson5","Defect","OPEN","Medium","","","Added suport for JSON-RPC 2.0 in jsonlib","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=5
    ' Reported by telmo.ca...@gmail.com, Jun 16, 2009
    ' ANSWER
    ' -----------
    ' Code sample added, commented out - needs test case
'"vbajson6","Defect","FIXED","Medium","","","Enter one-line summary","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=6
    ' Reported by yama...@gmail.com, Sep 4, 2009
    ' I found loop permanently when a string "key" include a colon.
    ' so i changed "parseKey()" tentatively. as following:
    '     Case ":"
    '        If Not dquote And Not squote Then
    '           index = index + 1
    '           Exit Do
    '        ElseIf dquote And Not squote Then
    '            parseKey = parseKey & char
    '           index = index + 1
    '        End If
'"vbajson7","Defect","New","HIGH","","","Cannot parse a JSON string containing an array...","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=7
    'Reported by c...@gmx.net, Oct 18, 2009
    'What steps will reproduce the problem?
    '1. Put this string in a variable:
    '
    '{"total_rows":36778,"offset":26220,"rows":[
    '{"id":"6b80c0b76","key":"a@bbb.net","value":{"entryid":"81151F241C2500","subject":"test subject","senton":"2009-7-09 22:03:43"}},
    '{"id":"b10ed9bee","key":"b@bbb.net","value":{"entryid":A7C3CF74EA95C9F","subject":"test subject2","senton":"2009-4-21 10:18:26"}}]}
    '
    '2. Instantiate a jsonlib object:  "Dim lib As New jsonlib"
    '3. Define a new JSON object: "Dim json As Object"
    '4. Instantiate the JSON object by invoking the jsonlib's "parse" method, the JSON string is the  parameter: "Set json = lib.parse(mystring)"
    '
    'What is the expected output? What do you see instead?
    'I would expect to be able to access the elements in the json object; the parse method returns an error.
    '
    'What version of the product are you using? On what operating system?
    'r2 from Feb 14,2009 - OS = Windows XP
    '
    'Please provide any additional information below.
    'Parsing JSON strings containing a single record works perfectly, I'm using your VBA library to read/write/delete data in CouchDB.
    '
    ' ANSWER:
    ' There is a " missing before A7C3CF74EA95C9F so this will not parse correctly.
    '---------
    '"vbajson7b"
    '#1 amrita.c...@gmail.com
    ' {"Cusip":[123,456,789],"Date":[1,2,3],"CloseType":["stock","bond","stock"]}
    'THIS IS MY JSON STRING
    'but when i try to parse(mystring)..I get back the same string
'"vbajson8","Defect","FIXED","Medium","","","Cannot convert a 2-d array to JSON","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=8
    'Reported by mheber...@gmail.com, Jan 15, 2010
    '-What steps will reproduce the problem?
    '1. Create a 2-d array, such as:
    'Dim arr(0 To 1, 0 To 1) As String
    'arr(0, 0) = "a"
    'arr(0, 1) = "b"
    'arr(1, 0) = "c"
    'arr(1, 1) = "d"
    '2. Try to convert to JSON with
    'Debug.Print lib.toString(arr)
    '-What is the expected output? What do you see instead?
    'I expect [["a", "b"], ["c", "d"]]
    'but I get a "Type Mismatch" error. If I change the array type to Variant, I
    'get the following:
    '[[,],[,]]
    'but it returns an error.
    '-What version of the product are you using? On what operating system?
    'json.xls downloaded on 1/15/2010 from the Google Code Site.
    '-Please provide any additional information below.
    'Thanks for looking into it!
    '
    'Oct 14, 2014 #1 walid.na...@gmail.com
    'One work-around seems to be to use an array of arrays instead of a 2-d array.
    'Also, arrays seem to need to be of type Variant for this to work.
    'So
    '    Dim arr(0 To 3) As Variant
    '    arr(0) = "a"
    '    arr(1) = "b"
    '    arr(2) = "c"
    '    arr(3) = "d"
    'works. And
    '    Dim arr(1 To 2) As Variant
    '    arr(1) = Array("a", "b")
    '    arr(2) = Array("c", "d")
    'works.
'"vbajson9","Defect","CLOSED","Medium","","","Thank you for this code!","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=9
    ' Reported by compwiz...@gmail.com, Apr 7, 2010
'"vbajson10","Defect","FIXED","Medium","","","improve parseNumber() with Long number","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=10
    'Reported by akun...@gmail.com, May 8, 2010
    '-What steps will reproduce the problem?
    '  1. Set json = lib.parse("{"BigNumber":32769}")
    '  2. an Exception was raise because the CInt Cannot process the Big number
    '-What is the expected output? What do you see instead?
    '  Debug.Assert json.Item("BigNumber") = 32769
    '-What version of the product are you using? On what operating system?
    'Please provide any additional information below.
    '
    'I use CLng to replace CInt
    'Private Function parseNumber(ByRef str As String, ByRef index As Long)
    '
    '    Dim value   As String
    '    Dim char    As String
    '
    '    Call skipChar(str, index)
    '    Do While index > 0 And index <= Len(str)
    '        char = Mid(str, index, 1)
    '        If InStr("+-0123456789.eE", char) Then
    '            value = value & char
    '            index = index + 1
    '        Else
    '            If InStr(value, ".") Or InStr(value, "e") Or InStr(value, "E") Then
    '                parseNumber = CDbl(value)
    '            Else
    '                parseNumber = CLng(value) 'CInt(value)
    '            End If
    '            Exit Function
    '        End If
    '    Loop
    'End Function
    'Aug 27, 2012 #1 djm...@googlemail.com
    'This needs CDbl instead of CInt. Otherwise you can get numbers which are written as a single number but are too big for long. (because it's not written using scientific format, the check for "e" or "E" won't pick it up).
'"vbajson11","Defect","FIXED","Medium","","","double backslash parse problem","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=11
    'Reported by akun...@gmail.com, May 12, 2010
    '-What steps will reproduce the problem?
    'I create an Test Case:
    '
    'Sub parse_test7()
    '
    '    Dim lib As New jsonlib
    '    Dim json As Object
    '
    '    Set json = lib.parse("{""Path"":""C:\\sample\\sample.jpg""}")
    '    Debug.Assert Err.Number = 0
    '
    '    Debug.Print lib.toString(json)
    '
    '    Set json = Nothing
    '    Set lib = Nothing
    '
    'End Sub
    '
    '-What is the expected output? What do you see instead?
    'expected:
    '{"Path":"C:\sample\sample.jpg"}
    'instead:
    '{"Path":"C:samplesample.jpg"}
    '-What version of the product are you using? On what operating system?
    'Please provide any additional information below.
    '
    'I manual fix by follow:
    '@@ -149,7 +147,10 @@
    '             index = index + 1
    '             char = Mid(str, index, 1)
    '             Select Case (char)
    '-            Case """", "\\", "/"
    '+            Case "\"
    '+                parseString = parseString & "\"
    '+                index = index + 1
    '+            Case """", "/", "'"
    '                 parseString = parseString & char
    '                 index = index + 1
    '             Case "b"
'"vbajson12","Defect","FIXED","Medium","","","String wont parse","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=12
    'Reported by ehb...@gmail.com, Aug 23, 2010
    'Why wont the parser parse the string below ?
    'It creates the root ListsState but leaves it empty, no sub objects are created.
    '
    '{"ListsState":{"MenuLocation":["Kelim","ChecklistTools"],"CurentLoadedChecklist":"ToolsConfig","InnerDoc":{"DapiotRegel":{"ClassName":"White","CHLTitle":"???? ???","Fields":{}},"ToolsConfig":{"ClassName":"White","CHLTitle":"??????","Fields":{"ToolsConfigHeliID":"036","ToolsConfigCrewSize":"3","ToolsConfigOperativeWgt":"1,500","ToolsConfigNumOf669":"0","ToolsConfigNumOf669Doc":"0","ToolsConfigNumOf669Med":"0","ToolsConfigNumOf669Equip":"0","ToolsConfigNumOfSol":"0","ToolsConfigNumOfPax":"0","ToolsConfigCargo":"0","ToolsConfigCar":"0","ToolsConfigFuelExtTanks":"0","ToolsConfigFuelTotal":"0","ToolsConfigCarUnits_Save":"?\"?"}}}}}
    '
    ' ANSWER
    ' -----------
    ' For VBA the " has to be "" and for JSON \ needs to be escaped as \\
    ' See vbajson12 test. I do not have the locale setup to verify the international characters.
'"vbajson13","Defect","OPEN","Medium","","","here's an update for office 64-bit support","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=13
    ' #2 sajja.pr...@gmail.com
    ' Hi I am trying to parse the below but I am getting an error that Object Not Found.
    '{"schedules":[{"summary":"Sign in","executedOn":"10/Oct/12 1:50 PM","cycleName":"asdf","cycleID":15,"label":"1, 2, 3, 4, 5","issueId":123,"versionName":"asdf","issueID":123,"defects":[
    '{"key":"124","status":"Closed","summary":"Title"},{"key":"asdf","status":"Closed","summary":"asdfasdf"}],"executedByDisplay":"Name of person","executionStatus":"2","htmlComment":"asdfasd","projectID":"asdf","executedBy":"asdasg","component":"","versionID":"adasd","issueKey":"asdf","scheduleID":73,"comment":"adsfasdf"},
    '{"summary":"asdf","executedOn":"10/Oct/12 1:17 PM","cycleName":"asdf","cycleID":15,"label":"1, 2, 3, 4, 5, 6, 7, 89, 5, 34","issueId":10012,"versionName":"sdf","issueID":10012,"defects":[
    '{"key":"asdf","asdf":"asdf","summary":"asdf"},{"key":"asdf","status":"Closed","summary":"asdf"}],"executedByDisplay":"asdf","executionStatus":"2","htmlComment":"asdf","projectID":10002,"executedBy":"asdf","component":"","versionID":10001,"issueKey":"Edf","scheduleID":18,"comment":"asdf"},
    '{"summary":"asdf","executedOn":"10/Oct/12 1:20 PM","cycleName":"asdf","cycleID":15,"label":"1, 2","issueId":10011,"versionName":"asdf","issueID":10011,"defects":[
    '{"key":"asdf","status":"Closed","summary":"asdf"},{"key":"asdf","status":"Closed","summary":"asdf - asdf"}],"executedByDisplay":"asdf","executionStatus":"2","htmlComment":"asdf","projectID":10002,"executedBy":"asdf","component":"","versionID":10001,"issueKey":"asdf","scheduleID":17,"comment":"asdf"},
    '{"summary":"asdfasdf","executedOn":"10/Oct/12 1:26 PM","cycleName":"asdf","cycleID":15,"label":"1,2","issueId":10010,"versionName":"asdf","issueID":10010,"defects":[
    '{"key":"asdf","status":"Closed","summary":"asdfa"},{"key":"asdf","status":"Closed","summary":"asdf"}],"executedByDisplay":"asdfasf","executionStatus":"2","htmlComment":"asdfafd","projectID":10002,"executedBy":"asdf","component":"","versionID":10001,"issueKey":"afgaf","scheduleID":16,"comment":"asdf"}]}
    '
    ' ANSWER
    ' ---------
    ' It validates correctly as JSON, but in the VBA IDE the string is too long.
    ' Broke it up and tested in vbajson13
    '
    ' OPEN
    ' ---------
    ' Reported by jho...@gmail.com, Jan 12, 2011
    ' String Builder Class and Office x64
    ' #1 myungdae...@soonjeonggame.com - Mar 10, 2014
    'use this
    '
    '#If Win64 Then
    'Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    '      (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    '#Else
    'Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    '      (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    '#End If
    '
    ' #5 glsmca.c...@gmail.com - Aug 27, 2014
    'Hello,
    'I am getting the error "Run-time error '424': Object Required".
    'How to fix this?
    '
    ' #7 ohr.hach...@gmail.com - Nov 12, 2014
    'i 'm getting this error: type mismatch at this line: CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
    'i added already the PtrSafe to the function decleration
'"vbajson14","Defect","FIXED","Medium","","","Unable to parse strings containing colons - Infinite loop","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=14
    'Reported by fadeyi.f...@gmail.com, Jun 1, 2012
    '-What steps will reproduce the problem?
    '1. In the parse_test4 subroutine, change any of the ""test"" elems to ""te:st""
    '2. Run the subroutine
    '-What is the expected output? What do you see instead?
    'The expected output should be the normal json output from the test. Instead, Excel freezes - because it is really good at detecting redundancies so it just loops them as fast as possible till windows asks if you want to crash the instance and restart.
    '-What version of the product are you using? On what operating system?
    'The latest version - with Microsoft Office 2007 running on Windows 7
    '
    'Please provide any additional information below.
    '
    'Great parser.It 's worked really well so far. I'll fix this bug in my instance, but I'm not sure if I should add my fix here or not. I'm pretty sure it won't be elegant :)
    'Jun 1, 2012 #1 fadeyi.f...@gmail.com
    'I fixed this by modifying the case for ":" in the method parsekey as follows:
    '        Case ":"
    '            index = index + 1
    '            If Not dquote And Not squote Then
    '                Exit Do
    '            Else
    '                parseKey = parseKey & char
    '            End If
'"vbajson15","Defect","New","Medium","","","Unable to handle multi-dimensional arrays","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=15
'"vbajson16","Defect","New","Medium","","","parseNumber and regional settings","Priority-Medium, Type-Defect",https://code.google.com/p/vba-json/issues/detail?id=16
    'Reported by bl.lio...@gmail.com, Oct 11, 2012
    'If you set "," as the decimal point in Control panel / Regional and Language settings then CDbl("12.34") will throw an error, but CDbl("12,34") will be parsed correctly.
    '
    'Some language uses comma for decimal point by default, so you can make more globalized parseNumber if you replace this:
    '  parseNumber = CDbl(Value)
    'to this:
    '  parseNumber = CDbl(Replace(Value, ".", Mid(CStr(0.1), 2, 1)))
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
' %010 -
' %009 -
' %008 - Review Ref: https://tools.ietf.org/html/rfc7159 - Proposed Standard - obsoletes rfc7158
' %007 - Review Ref: https://tools.ietf.org/html/rfc7158 - Proposed Standard - obsoletes rfc4627
' %006 - Review Ref: https://tools.ietf.org/html/rfc4627 - Original spec from Douglas Crockford
' %005 - Review Ref: http://bolinfest.com/essays/json.html - Json ES3 and ES5
' %004 - Review Ref: http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-404.pdf
' %002 - *** Ref: http://www.codeproject.com/Articles/720368/VB-JSON-Parser-Improved-Performance
' %001 - Have test result "VALIDATED" be verified automatically with online parser - TBD
' Issues:
' #010 -
' #009 -
' #008 -
' #006 - Error in multiArray if on error commented out when executing RunAlljsonlibTests
' #002 - vbajson2 still kills Excel
'=============================================================================================================================

' 20141214 - v016 - Use p for property variables and b for boolean
' 20141211 - v014 -
    ' FIXED - #001 - Run-time error '424' Object required in test vbajson1
    ' FIXED - #007 - parse_error_004 fails for the wrong reason
    ' FIXED - #004 - vbajson7 is a FAIL, test case vbajson7_fail
    ' DONE - %003 - Review http://www.ietf.org/rfc/rfc4627.txt
' 20141205 - v014 -
    ' Ref: http://www.intl-spectrum.com/resource/397/Parsing-JSON-data-in-MultiValue.aspx
    ' Ref: http://www.b4x.com/android/forum/threads/jsonparser-for-vb-net.44151/
    ' How do I preserve spaces in retrieving json keys? Ref: http://www.csspy.com/22_22639040/
    ' Added test vbajson1a
    ' http://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel
    ' FIXED #005 - vbajson1a runtime error 424 object required
' 20141201 - v012 - Add task to integrate improvements from task %002
' 20141126 - v011 - Move history to basTESTvbajsonlog
    ' FIXED vbajson14
    ' Add basTESTRUNNER module
' 20141125 - v011 - FIXED vbajson9
    ' FIXED vbajson10
    ' FIXED vbajson11
    ' FIXED vbajson12
    ' FIXED vbajson13, but vbajson13b OPEN and TBD for string builder class and Office x64
' 20141124 - v011 - FIXED vbajson3 - s/vbNewLine/vbLf
    ' FIXED vbajson4
    ' FIXED vbajson5. Test case needed.
    ' FIXED vbajson6. Test case needed.
    ' FIXED vbajson7b - @amrita.c... - paste the string into http://jsonlint.com/ validator and it validates.
' 20141121 - v011 - FIXED #003 - parse_test3 breaks RunAllTests


