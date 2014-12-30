Attribute VB_Name = "basTESTRUNNER"
Option Explicit
Option Compare Text
Option Private Module

Public Sub RUN_ALL_TESTS()

    RunTDDejsonvbaTests
    RunAllejsonlibTests
    RunAllejsonvbaTests
    RunAllejsonvbaErrorTests

End Sub
