'Simple VBA script for OA system. Analyze employee's attendance.
'Copyright (C) 2022  Runtong Wang, Miao Wang
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License along
'with this program; if not, write to the Free Software Foundation, Inc.,
'51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.'

Attribute VB_Name = "Module1"
Function findAbnormalDays(startDate As Date, endDate As Date, rownum As Integer)
    Dim tripSheet As Worksheet
    Dim errandSheet As Worksheet
    Dim leaveSheet As Worksheet
    Dim summarySheet As Worksheet
    Dim detailSheet As Worksheet
    Dim resSheet As Worksheet
    Dim excusedSheet As Worksheet
    Dim dateSheet As Worksheet
    Dim wbBook As Workbook
    Dim datestr As String
    Dim absent As Integer
    Dim lateArrive As Integer
    Dim earlyLeave As Integer
    Dim name As String
    Dim j As Integer
    Set wbBook = ActiveWorkbook
    Set tripSheet = wbBook.Sheets(2)
    Set errandSheet = wbBook.Sheets(3)
    Set leaveSheet = wbBook.Sheets(4)
    Set summarySheet = wbBook.Sheets(5)
    Set detailSheet = wbBook.Sheets(6)
    Set resSheet = wbBook.Sheets(7)
    Set excusedSheet = wbBook.Sheets(8)
    wbBook.Sheets.Add After:=excusedSheet
    Set dateSheet = wbBook.Sheets(9)
    Dim curDate As Date
    curDate = startDate
    absent = 0
    lateArrive = 0
    earlyLeave = 0
    j = 0
    name = summarySheet.Range("A" & rownum)
    While curDate <= endDate
        If Weekday(curDate) = 2 Or Weekday(curDate) = 7 Then
            GoTo ENDLOOP
        End If
        
        'skip check excused date
        If excusedSheet.UsedRange.Rows.Count > 3 Then
        For i = 1 To excusedSheet.Cells(excusedSheet.Rows.Count, "A").End(xlUp).Row
            If IsDate(excusedSheet.Range("A" & i)) Then
                If DateDiff("d", excusedSheet.Range("A" & i), curDate) <= DateDiff("d", excusedSheet.Range("A" & i), excusedSheet.Range("B" & i)) And DateDiff("d", excusedSheet.Range("A" & i), curDate) >= 0 Then GoTo ENDLOOP
            End If
        Next
        End If
        
        'filter current date
        datestr = Application.WorksheetFunction.Text(curDate, "YYYY-MM-DD")
        resSheet.Range("A1:C" & resSheet.Cells(resSheet.Rows.Count, "C").End(xlUp).Row).AutoFilter field:=1, Criteria1:=datestr & "*"
        resSheet.Range("A1:A" & resSheet.Cells(resSheet.Rows.Count, "A").End(xlUp).Row).Copy dateSheet.Range("A1")
        resSheet.Range("A1:C" & resSheet.Cells(resSheet.Rows.Count, "C").End(xlUp).Row).AutoFilter
        
        'check if checkin or checkout data exist
        If dateSheet.UsedRange.Rows.Count <= 1 Then
        absent = absent + 1
        detailSheet.Range("B" & rownum).Offset(0, j).Interior.ColorIndex = 3
        detailSheet.Range("B" & rownum).Offset(0, j + 1).Interior.ColorIndex = 3
        GoTo ENDLOOP
        End If
        
        For i = 2 To dateSheet.Cells(dateSheet.Rows.Count, "A").End(xlUp).Row
            datetimestr = dateSheet.Range("A" & i)
            dateSheet.Range("B" & i) = TimeValue(Split(datetimestr, " ")(1))
        Next
        'check if late early
        detailSheet.Range("B" & rownum).Offset(0, j) = dateSheet.Application.WorksheetFunction.min(dateSheet.Range("B1:B" & i))
        detailSheet.Range("B" & rownum).Offset(0, j + 1) = dateSheet.Application.WorksheetFunction.Max(dateSheet.Range("B1:B" & i))
        If dateSheet.Application.WorksheetFunction.Max(dateSheet.Range("B1:B" & i)) < TimeValue("17:00:00") Then
            earlyLeave = earlyLeave + 1
            detailSheet.Range("B" & rownum).Offset(0, j + 1).Interior.ColorIndex = 6
        End If
        If dateSheet.Application.WorksheetFunction.min(dateSheet.Range("B1:B" & i)) > TimeValue("8:00:00") Then
            lateArrive = lateArrive + 1
            detailSheet.Range("B" & rownum).Offset(0, j).Interior.ColorIndex = 6
        End If
ENDLOOP:
        curDate = DateAdd("d", 1, curDate)
        j = j + 2
        dateSheet.Cells.Clear
    Wend
    Debug.Print absent
    Debug.Print earlyLeave
    Debug.Print lateArrive
    summarySheet.Range("B" & rownum) = absent
    summarySheet.Range("C" & rownum) = lateArrive
    summarySheet.Range("D" & rownum) = earlyLeave
    findAbnormalDays = 0
    dateSheet.Delete
End Function

Function calculateattendence(startDate As Date, endDate As Date)
    Dim dataSheet As Worksheet
    Dim tripSheet As Worksheet
    Dim errandSheet As Worksheet
    Dim leaveSheet As Worksheet
    Dim summarySheet As Worksheet
    Dim detailSheet As Worksheet
    Dim resSheet As Worksheet
    Dim excusedSheet As Worksheet
    Dim wbBook As Workbook
    Dim name As String
    Dim j As Integer
    Dim res As Variant
    Set wbBook = ActiveWorkbook
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    Set dataSheet = wbBook.Sheets(1)
    Set tripSheet = wbBook.Sheets(2)
    Set errandSheet = wbBook.Sheets(3)
    Set leaveSheet = wbBook.Sheets(4)
    Set summarySheet = wbBook.Sheets(5)
    Set detailSheet = wbBook.Sheets(6)
    wbBook.Sheets.Add After:=detailSheet
    Set resSheet = wbBook.Sheets(7)
    wbBook.Sheets.Add After:=resSheet
    Set excusedSheet = wbBook.Sheets(8)
    For j = 2 To summarySheet.UsedRange.Rows.Count
        name = summarySheet.Range("A" & j)
        
        'filter and save errand data
        errandSheet.Range("A3:F" & errandSheet.Cells(errandSheet.Rows.Count, "F").End(xlUp).Row).AutoFilter field:=2, Criteria1:=name
        errandSheet.Range("D3:E" & errandSheet.Cells(errandSheet.Rows.Count, "E").End(xlUp).Row).Copy excusedSheet.Range("A1")
        errandSheet.Range("A3:F" & errandSheet.Cells(errandSheet.Rows.Count, "F").End(xlUp).Row).AutoFilter
        
        'filter and save leave data
        leaveSheet.Range("A1:J" & leaveSheet.Cells(leaveSheet.Rows.Count, "J").End(xlUp).Row).AutoFilter field:=1, Criteria1:=name
        leaveSheet.Range("E2:F" & leaveSheet.Cells(leaveSheet.Rows.Count, "F").End(xlUp).Row).Copy excusedSheet.Range("A" & excusedSheet.Cells(excusedSheet.Rows.Count, "A").End(xlUp).Row + 1)
        leaveSheet.Range("A1:J" & leaveSheet.Cells(leaveSheet.Rows.Count, "J").End(xlUp).Row).AutoFilter
        
        'filter and save trip data
        tripSheet.Range("A3:C" & tripSheet.Cells(tripSheet.Rows.Count, "C").End(xlUp).Row).AutoFilter field:=1, Criteria1:=name
        tripSheet.Range("B5:C" & tripSheet.Cells(tripSheet.Rows.Count, "C").End(xlUp).Row).Copy excusedSheet.Range("A" & excusedSheet.Cells(excusedSheet.Rows.Count, "A").End(xlUp).Row + 1)
        tripSheet.Range("A3:C" & tripSheet.Cells(tripSheet.Rows.Count, "C").End(xlUp).Row).AutoFilter
        detailSheet.Range("A" & j) = name
        dataSheet.Range("A1:I" & dataSheet.Cells(dataSheet.Rows.Count, "I").End(xlUp).Row).AutoFilter field:=3, Criteria1:=name
        dataSheet.Range("A1:A" & dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row).Copy resSheet.Range("A1")
        dataSheet.Range("A1:I" & dataSheet.Cells(dataSheet.Rows.Count, "I").End(xlUp).Row).AutoFilter
        res = findAbnormalDays(startDate, endDate, j)
        resSheet.Cells.Clear
        excusedSheet.Cells.Clear
    Next
    resSheet.Delete
    excusedSheet.Delete
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    calculateattendence = 0
End Function

Sub attendanceAgent()
    Dim summarySheet As Worksheet
    Dim detailSheet As Worksheet
    Dim wbBook As Workbook
    Dim startDate As Date
    Dim endDate As Date
    Dim res As Integer
    Dim i As Integer
    Set wbBook = ActiveWorkbook
    Set summarySheet = wbBook.Sheets(5)
    wbBook.Sheets.Add After:=summarySheet
    Set detailSheet = wbBook.Sheets(6)
    startDate = summarySheet.Range("H5")
    endDate = summarySheet.Range("I5")
    Dim curDate As Date
    Dim curRange As Range
    curDate = startDate
    i = 0
    Set curRange = detailSheet.Range("B1")
    While curDate <= endDate
        curRange.Offset(0, i) = curDate
        i = i + 2
        curDate = DateAdd("w", 1, curDate)
    Wend
    res = calculateattendence(startDate, endDate)
End Sub
