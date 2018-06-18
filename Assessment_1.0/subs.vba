' Author: Jukka-Pekka Kerkkänen 11/2016

Option Explicit
Sub Worksheet_Activate()
 MsgBox "Moi"
End Sub
Sub showTestAverages(ByVal Target As Range)
  Dim firstColOfAverages, numbOfTests, j As Integer
  Dim row, col As Long

  numbOfTests = Functions.numberOfTestsInTotal

  firstColOfAverages = Constants.firstTestCol + numbOfTests * 3 + 2

  ' Tests if target is one of the custom buttons
  If Target.row = 2 And Target.Column = firstColOfAverages Then
    Subs.stopAutoUpdates
    Cells(1, firstColOfAverages).Value = Constants.allTests

    For j = firstColOfAverages To firstColOfAverages + 12 Step 2
      Cells(6, j).Value = Constants.allTestsNotice_fi

    Next j

    Subs.countTestAverages
    Subs.enableAutoUpdates ' Outside this prevents copy&paste!

  ElseIf Target.row = 2 And Target.Column = firstColOfAverages + 5 Then
    Subs.stopAutoUpdates
    Cells(1, firstColOfAverages).Value = Constants.ATests

    For j = firstColOfAverages To firstColOfAverages + 12 Step 2
      Cells(6, j).Value = Constants.ATestsNotice_fi

    Next j
    Subs.countTestAverages
    Subs.enableAutoUpdates

  ElseIf Target.row = 2 And Target.Column = firstColOfAverages + 10 Then
    Subs.stopAutoUpdates
    Cells(1, firstColOfAverages).Value = Constants.BTests

    For j = firstColOfAverages To firstColOfAverages + 12 Step 2
      Cells(6, j).Value = Constants.BTestsNotice_fi

    Next j
    Subs.countTestAverages
    Subs.enableAutoUpdates
  End If


  'Selection.Interior.Color = RGB(255, 255, 255)
End Sub
Sub countTestAverages()

  Dim col, row, firstCol, lastCol, lastRow, numbOfStud, counter As Long
  Dim avType, periodOfYear As Long
  Dim av As Double
  Dim colLet As String

  firstCol = Functions.firstAverageCol
  lastCol = firstCol + 13       '14 average cols altogether

  ' Toggles between 1 and 2 (pergentage / marks)
  avType = 2

  numbOfStud = Functions.haeOppilaidenLkm
  lastRow = Constants.firstNameRow + numbOfStud - 1
  counter = 1

  For col = firstCol To lastCol Step 1
    ' colLet = Functions.colLetter((col))
    If avType = 1 Then
      avType = 2
    Else
      avType = 1
    End If

    ' Counter helps to define the right period(s) of year:
    If counter < 3 Then
      periodOfYear = 1
    ElseIf counter < 5 Then
      periodOfYear = 2
    ElseIf counter < 7 Then
      periodOfYear = 3
    ElseIf counter < 9 Then
      periodOfYear = 4
    ElseIf counter < 11 Then
      periodOfYear = 12
    ElseIf counter < 13 Then
      periodOfYear = 34
    Else
      periodOfYear = 1234
    End If

    For row = Constants.firstNameRow To lastRow Step 1

      av = Functions.countAvOfTests((avType), (periodOfYear), (row))

      ' Empty if no tests done, otherwise number:
      If av = Constants.noTestsDone Then
        Cells(row, col).Value = Constants.noValue

      Else
        Cells(row, col).Value = av
      End If

    Next row
    counter = counter + 1
  Next col

  ' Total averages:
  For col = firstCol To lastCol Step 1

    av = Functions.countAverageJP(Cells(Constants.firstNameRow, col))

    ' Empty if no tests done, otherwise number:
    If av = Constants.noTestsDone Then
      Cells(row, col).Value = Constants.noValue

    Else
      Cells(row, col).Value = av
    End If

  Next col

End Sub
' Cleares a row of the averages and the totals row.
' If the parameter = Constants.clearAll, cleares all the rows.
Sub clearAverages(ByVal row As Long)
  Dim col, firstCol, lastCol, firstRow, lastRow, numbOfStud As Long

  firstCol = Functions.firstAverageCol
  lastCol = firstCol + 13       '14 average cols altogether

  numbOfStud = Functions.haeOppilaidenLkm

  firstRow = Constants.firstNameRow
  lastRow = Constants.firstNameRow + numbOfStud ' Including the averages row

  ' Either one row or all of them:
  If row = Constants.clearAll Then

    Lib_jp1.createRangeByNumbers(firstCol, lastCol, firstRow, lastRow).ClearContents

    ' Notice just ones:
    Cells(firstRow, firstCol + 5).Value = Constants.changeDetectedNotice

  Else ' Only targeted row and totals row are cleared if it has not been done before.
    If Cells(row, firstCol + 5).Value <> Constants.changeDetectedNotice Then

      Lib_jp1.createRangeByNumbers(firstCol, lastCol, row, row).ClearContents
      Lib_jp1.createRangeByNumbers(firstCol, lastCol, lastRow, lastRow).ClearContents

      Cells(row, firstCol + 5).Value = Constants.changeDetectedNotice
    End If
  End If

End Sub

' Checks where a change has occurred and clears
' averages if they needed (a manual) update.
Sub clearAveragesIfNeeded(Target As Range)
  Dim numbOfTests, firstColOfTests, lastColOfTests, lastStudRow, j As Integer
  Dim row, col As Long

  firstColOfTests = 3
  lastColOfTests = Functions.firstAverageCol - 3
  lastStudRow = Functions.haeOppilaidenLkm + Constants.firstNameRow - 1

  ' Tests if target is one of the test Constants cells.
  ' In that case all the averages must be recalculated.
  If Target.Column >= firstColOfTests And Target.Column <= lastColOfTests Then
    If Target.row > 1 And Target.row < 9 Then

      Call Subs.clearAverages(Constants.clearAll)

    ElseIf Target.row >= Constants.firstNameRow And Target.row <= lastStudRow Then

      Call Subs.clearAverages(Target.row)

    End If
  End If
End Sub


Sub lisaa_uusi_rivi()

  Dim vikarivi, seur, ed As Integer

  Dim col, lastCol As Long
  Dim colLet As String

  ' Disable automatic calculation
  Subs.stopAutoUpdates

  vikarivi = Functions.haeOppilaidenLkm() + Constants.firstNameRow - 1
  seur = vikarivi + 1
  ed = vikarivi - 1

  Rows(vikarivi & ":" & vikarivi).Copy
  Rows(seur & ":" & seur).Insert Shift:=xlDown
  'Rows(seur & ":" & seur).ClearContents
  Application.CutCopyMode = False

  Range("B" & vikarivi & ":W" & vikarivi).AutoFill Destination:= _
      Range("B" & vikarivi & ":W" & seur), Type:=xlFillDefault

  Range("A" & ed & ":A" & vikarivi).AutoFill Destination:= _
      Range("A" & ed & ":A" & seur), Type:=xlFillDefault

  ' Empties all copied results and grades:
  Range("B" & seur & ":C" & seur).ClearContents
  col = 3 ' column C
  lastCol = Functions.numberOfTestsInTotal * 3

  ' Results
  For col = 3 To lastCol Step 3
    Cells(seur, col).Value = ""
  Next col

  ' Grades
  For col = lastCol + 21 To lastCol + 30
    Cells(seur, col).Value = ""
  Next col

  ' Enable automatic calculation
  Subs.enableAutoUpdates


End Sub

Sub poista_vika_rivi()
  Dim vikarivi As Integer

  ' Disable automatic calculation
  Subs.stopAutoUpdates

  vikarivi = Functions.haeOppilaidenLkm() + firstNameRow - 1

  ' Tarkistetaan ensin, onko oppilaan nimi tyhjä:
  If Cells(vikarivi, 2).Value = "" Then
      ' poista_vika_rivi Macro
      Rows(vikarivi & ":" & vikarivi).Delete Shift:=xlUp
  Else
      MsgBox "Oppilaan nimi ei ole tyhjä! Ei voi poistaa."
  End If

  ' Enable automatic calculation
  Subs.enableAutoUpdates

End Sub




Sub addTest()
'
' add_test Makro

  Dim numberOfPupils, numberOfTests, lastCol As Integer

  ' Disable automatic calculation
  Subs.stopAutoUpdates

  numberOfPupils = Functions.haeOppilaidenLkm
  numberOfTests = Functions.numberOfTestsInTotal
  lastCol = 5 + (numberOfTests - 1) * 3

  Columns(Subs.colLetter(lastCol - 2) & ":" & Subs.colLetter(lastCol)).Copy
  Columns(Subs.colLetter(lastCol + 1) & ":" & Subs.colLetter(lastCol + 1)). _
    Insert Shift:=xlToRight

  Application.CutCopyMode = False

  Range(Subs.colLetter(lastCol + 1) & firstNameRow & ":" & Subs.colLetter(lastCol + 1) & _
    firstNameRow - 1 + numberOfPupils).ClearContents


  Range(Subs.colLetter(lastCol + 3) & "1:" & Subs.colLetter(lastCol + 3) & "1"). _
    Value = numberOfTests + 1


  ' enable automatic calculation
  Subs.enableAutoUpdates


End Sub

Sub delLastEmptyTest()
'
' Deleting last (to the right) test if it has no points

  Dim numberOfPupils, numberOfTests, lastCol, row, col As Integer
  Dim coord As String

  Subs.stopAutoUpdates

  numberOfPupils = Functions.haeOppilaidenLkm
  numberOfTests = Functions.numberOfTestsInTotal
  lastCol = 5 + (numberOfTests - 1) * 3

  ' Tarkistetaan ensin, onko pisteiden keskiarvo tyhjä, eli kokeessa ei merkintöjä:
  row = firstNameRow + numberOfPupils
  col = lastCol - 2

  If lastCol > 5 And Cells(row, col).Value = "" Then

    ' don't know why but lastCol caused an error without +0! Should be the same type, but...
    Columns(Subs.colLetter(lastCol - 2) & ":" & Subs.colLetter(lastCol + 0)). _
      Delete Shift:=xlToLeft
  Else
    MsgBox "Kokeessa on pistemerkintöjä tai se on ainut kappale! Ei voi poistaa."

  End If

  Subs.enableAutoUpdates

End Sub


Sub hideAverages(hide As Boolean)
'
' showHiddenColumns Makro
'
Dim numberOfTests, lastCol, firstAverageCol As Integer
  Dim colStr As String

  numberOfTests = Functions.numberOfTestsInTotal
  lastCol = 5 + (numberOfTests - 1) * 3
  firstAverageCol = lastCol + 3

  colStr = Subs.colLetter(firstAverageCol) & ":" & _
    Subs.colLetter(firstAverageCol + 15)

  Columns(colStr).EntireColumn.Hidden = hide
End Sub
Sub hideTests(hide As Boolean)
'
' hide_averages Makro
  Dim numberOfTests, lastCol, firstCol As Integer
  Dim colStr As String

  numberOfTests = Functions.numberOfTestsInTotal
  lastCol = 5 + (numberOfTests - 1) * 3
  firstCol = 3

  colStr = Subs.colLetter(firstCol) & ":" & _
    Subs.colLetter(lastCol + 0)

  Columns(colStr).EntireColumn.Hidden = hide
End Sub
Public Sub selectColumnsTest(startCol As Integer, numberOfCols As Integer)
  Columns(startCol).Resize(, numberOfCols).Select
End Sub

'http://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter

Function colLetter(lngCol As Integer) As String
  Dim vArr
  vArr = Split(Cells(1, lngCol).Address(True, False), "$")
  colLetter = vArr(0)
End Function


Sub stopAutoUpdates()
  With Excel.Application
   .ScreenUpdating = False
   .EnableEvents = False
   .Calculation = Excel.xlCalculationManual
  End With
End Sub

Sub enableAutoUpdates()
  With Excel.Application
   .ScreenUpdating = True
   .EnableEvents = True
   .Calculation = Excel.xlCalculationAutomatic
  End With
End Sub
