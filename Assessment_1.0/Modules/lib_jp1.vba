' Here there are functions and subs that could be generally usefull.
' Author: Jukka-Pekka Kerkkänen (et al; see links) 06/2018

Option Explicit

' Creates a Range-object based on Long parameters given.
Function createRangeByNumbers(ByVal col1 As Long, ByVal col2 As Long, _
                              ByVal row1 As Long, ByVal row2 As Long) As Range

  Dim colLet1, colLet2 As String

  colLet1 = Lib_jp1.colLetter(col1)
  colLet2 = Lib_jp1.colLetter(col2)

  ' Range is an object - Set needed here!
  Set createRangeByNumbers = Range(colLet1 & row1 & ":" & colLet2 & row2)

End Function

'http://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter

Function colLetter(ByVal lngCol As Long) As String
  Dim vArr
  vArr = Split(Cells(1, lngCol).Address(True, False), "$")
  colLetter = vArr(0)
End Function
