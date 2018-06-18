' Author: Jukka-Pekka Kerkkänen 11/2016

Option Explicit

Function numberOfTestsInPeriod(n As Integer) As Integer

  Dim row, total, col, j, max, period As Integer
  max = 20
  row = 4
  total = 0
  col = 5

  For j = col To col + 3 * max Step 3
    period = Cells(row, j).Value
    If period = n Then
      total = total + 1
    End If

  Next j
  numberOfTestsInPeriod = total

End Function

Function haeOppilaidenLkm() As Integer
    Dim nimisarake As Range
    Dim solu As Range

    ' Huom: sama kuin funktion nimi!
    haeOppilaidenLkm = 0

    'Katsotaan 50 ekaa riviä. Tuskin isompia ryhmiä heti tulee. Ei toki tukeva näin..
    ' Set nimisarake = Range("A" & Constants.firstNameRow & ":A & Constants.firstNameRow+50")
    Set nimisarake = Range("A10:A60")

    For Each solu In nimisarake.Cells
        If solu.Value > haeOppilaidenLkm Then
            haeOppilaidenLkm = solu.Value
        ElseIf solu.Value = 0 Then
            Exit For
        End If
    Next solu
End Function


Function numberOfTestsInTotal() As Integer
  numberOfTestsInTotal = _
    Functions.numberOfTestsInPeriod(1) + Functions.numberOfTestsInPeriod(2) + _
    Functions.numberOfTestsInPeriod(3) + Functions.numberOfTestsInPeriod(4)

End Function
Function firstAverageCol() As Integer
  Dim numbOfExtraCols As Integer

  'Extra cols: 2 before the first test + 2 after
  numbOfExtraCols = Constants.firstTestCol + 1
  firstAverageCol = Functions.numberOfTestsInTotal * 3 + numbOfExtraCols + 1

End Function




' Calculates the average of cells that begin from the
' cell given as parameter and continues down the number
' of pupils cells
Function countAverageJP(firstRowCell As Range) As Double

  ' This makes Excel recalculate this function in case of change in worksheet
  Application.Volatile True

  Dim arvoalue, cell As Range
  Dim lkm, oppilaidenLkm, firstRow, lastRow As Long
  Dim col As Long
  Dim sum As Double
  Dim colLet As String


  ' MsgBox firstRowCell.Column
  col = firstRowCell.Column  ' returns long!
  colLet = Functions.colLetter(col)

  firstRow = firstRowCell.row
  firstRow = 10

  oppilaidenLkm = Functions.haeOppilaidenLkm
  lastRow = firstRow + oppilaidenLkm - 1

  ' Values in the beginning
  lkm = 0
  sum = 0

  ' Arvoalue alkaa parametrisolusta ja jatkuu siitä alas oppilaiden
  ' lukumäärän verran:
  Set arvoalue = Range(colLet & firstRow & ":" & colLet & lastRow)

  For Each cell In arvoalue.Cells
    If IsNumeric(cell.Value) And cell.Value <> Constants.noTestsDone And cell.Value <> Constants.blank Then
      lkm = lkm + 1
      sum = sum + cell.Value
    End If
  Next cell

  If lkm = 0 Then
    countAverageJP = Constants.noTestsDone
  Else
    countAverageJP = sum / lkm
  End If

End Function

' Calculates the average of the tests taking into account the period, test type and weight:
Function countAvOfTests(avType As Integer, period As Integer, currRow As Long) As Double

  'MsgBox "Muutos tapahtui!"

  Dim testsMode, firstColOfAverages, numbOfTests, thisPeriod, number, j As Integer
  Dim thisRow, dataCol As Long
  Dim weight, weightTotal, av As Double
  Dim avOfPro As Boolean
  Dim testType As String

  numbOfTests = Functions.numberOfTestsInTotal
  number = 0
  weightTotal = 0
  av = 0  ' Average to be returned

  firstColOfAverages = Constants.firstTestCol + numbOfTests * 3 + 2

  ' testsMode number is hidden in the worksheet (font color = bg color)
  If IsNumeric(Cells(1, firstColOfAverages).Value) And _
    Cells(1, firstColOfAverages).Value > 0 Then
    testsMode = Cells(1, firstColOfAverages).Value
  Else
    testsMode = Constants.allTests
  End If

  thisRow = currRow

  If avType = Constants.avOfProcents Then
    avOfPro = True
  Else
    avOfPro = False
  End If

  ' lets go through all the tests
  dataCol = Constants.firstTestCol + 2

  ' Sum of all weights is needed for calculation of weighted average.
  ' The same checkings must be done as in the next for. Not very nice.
  For j = dataCol To dataCol + 3 * (numbOfTests - 1) Step 3

    testType = Cells(Constants.testTypeRow, j).Value
    weight = Cells(Constants.testWeightRow, j).Value
    thisPeriod = Cells(Constants.testPeriodRow, j).Value

    ' First tests if the period is right.
    ' Without double parenthesis typemismatch error ..'
    If Functions.thisPeriodIsOk(period, (thisPeriod)) Then
      If testTypeIsOk((testsMode), testType) Then

        ' Only includes weight if grade is not empty (student must have made the test):
        If IsNumeric(Cells(thisRow, j - 1).Value) And IsNumeric(Cells(thisRow, j).Value) Then

          ' If weight value is not good, it's replaced by 100 (default) and a warning is
          ' shown.
          If Not IsNumeric(Cells(Constants.testWeightRow, j).Value) Then
            Cells(Constants.testWeightRow, j).Value = 100
            MsgBox " Weight value was erroneous and has been replaced by 100 (default)."
          End If

          weightTotal = weightTotal + Cells(Constants.testWeightRow, j).Value

        End If  ' isnumeric
      End If  ' test of right type
    End If  ' test in right period
  Next j

  ' Then the second round and the real calculation:
  For j = dataCol To dataCol + 3 * (numbOfTests - 1) Step 3
    testType = Cells(Constants.testTypeRow, j).Value
    weight = Cells(Constants.testWeightRow, j).Value
    thisPeriod = Cells(Constants.testPeriodRow, j).Value

    ' First tests if the period is right.
    If Functions.thisPeriodIsOk(period, (thisPeriod)) Then
      If Functions.testTypeIsOk((testsMode), testType) Then

        ' Checks if the values are numerical
        If IsNumeric(Cells(thisRow, j - 1).Value) And IsNumeric(Cells(thisRow, j).Value) Then

          ' Number of tests included in the average:
          number = number + 1

          ' if the average is of percents or of grades:
          If avType = Constants.avOfProcents Then
            av = av + weight / weightTotal * Cells(thisRow, j - 1).Value
          Else
            av = av + weight / weightTotal * Cells(thisRow, j).Value
          End If  ' average calculated of percentages
        End If  ' isnumeric
      End If  ' test of right type
    End If  ' test in right period
  Next j

  If number = 0 Then
    countAvOfTests = Constants.noTestsDone
  Else
    countAvOfTests = av

  End If

End Function
' Tests if testType is right and returns true or false.
' If testType is something else than "A" (for instance "Exam") it's handled
' as "B".
Function testTypeIsOk(testsMode As Integer, testType As String) As Boolean

  If testsMode = Constants.allTests _
    Or (testType = "A" And testsMode = Constants.ATests) _
    Or (testType = "a" And testsMode = Constants.ATests) Then

    testTypeIsOk = True

  ElseIf (testType <> "A" And testType <> "a" And testsMode = Constants.BTests) Then
    testTypeIsOk = True
  Else
    testTypeIsOk = False
  End If
End Function
' Tests if thisPeriod is right for the average and returns true or false.
Function thisPeriodIsOk(period As Integer, thisPeriod As Integer)
  thisPeriodIsOk = False

  ' Go through cases where value is true:
  If thisPeriod = period Or period = Constants.testPeriod1234 Then
    thisPeriodIsOk = True
  ElseIf period = Constants.testPeriod12 Then
    If thisPeriod = Constants.testPeriod1 Or thisPeriod = Constants.testPeriod2 Then
      thisPeriodIsOk = True
    End If
  ElseIf period = Constants.testPeriod34 Then
    If thisPeriod = Constants.testPeriod3 Or thisPeriod = Constants.testPeriod4 Then
      thisPeriodIsOk = True
    End If
  End If

End Function

'http://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter

Function colLetter(lngCol As Long) As String
  Dim vArr
  vArr = Split(Cells(1, lngCol).Address(True, False), "$")
  colLetter = vArr(0)
End Function
