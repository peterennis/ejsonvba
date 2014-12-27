VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ejsonlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Copyright (c) 2013 Peter F. Ennis
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation;
'version 3.0.
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, visit
'http://www.gnu.org/licenses/lgpl-3.0.txt

'=======================================================================
' Author:   Peter F. Ennis
' Date:     December 19, 2014
' Comment:  ejsonvba => extended JSON for VBA
' GitHub:   https://github.com/peterennis/ejsonvba
'=======================================================================

Private Const ejsonlibVERSION As String = "0.0.0"
Private Const ejsonlibVERSION_DATE As String = "December 26, 2014"
'

Private Sub Class_Initialize()
    On Error GoTo 0
    Debug.Print "Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo 0
    Debug.Print "Class_Terminate"
End Sub

Public Property Get Parse(strVBA As String) As Object
    On Error GoTo 0
    Debug.Print "Get Parse"
End Property

Public Property Get ParseError() As String
    On Error GoTo 0
    Debug.Print "Get ParseError"
End Property

Public Property Get ToString(v As Variant) As String
    On Error GoTo 0
    Debug.Print "Get ToString"
End Property

Public Property Get WhiteSpace(str As String) As String
    On Error GoTo 0
    Debug.Print "Get WhiteSpace"
End Property

Public Function SkipChar(str As String, ind As Long) As Boolean
    On Error GoTo 0
    Debug.Print "Function SkipChar"
End Function

