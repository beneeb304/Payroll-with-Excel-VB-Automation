Imports Microsoft.Office.Interop

Module Main

    '------------------------------------------------------------
    '-                File Name : Main.vb                       -
    '-                Part of Project: Assignment9              -
    '------------------------------------------------------------
    '-                Written By: Benjamin Neeb                 -
    '-                Written On: April 30, 2021                -
    '------------------------------------------------------------
    '- File Purpose:                                            -
    '-                                                          -
    '- This file contains the main Sub for the console          -
    '- application. This file performs all of the interaction   -
    '- with data and Excel.                                     -
    '------------------------------------------------------------
    '- Program Purpose:                                         -
    '-                                                          -
    '- This program creates an initial set of employee data.    -
    '- The data is then written to an Excel file. No            -
    '- calculations are performed in VB, but rather formulas    -
    '- are written to Excel to conduct calculations.            -
    '------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):             -
    '- (None)                                                   -
    '------------------------------------------------------------

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    Const intCOLUMNCOUNT As Integer = 5     'Amount of columns used in Excel
    Const intOFFSET As Integer = 2          'Amount to offset pieces of data

    '-----------------------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-----------------------------------------------------------------------------------

    Sub Main()

        '------------------------------------------------------------
        '-              Subprogram Name: Main                       -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: April 30, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine creates employees, essential Excel       -
        '- objects, and calls all subs in the file.                 -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- myEmps:      List of clsEmployee that holds all of the   -
        '-              employees used in the program.              -
        '- aSheet:      Instance of Excel as an application.        -
        '- CheckExcel:  Object used to determine if Excel is        -
        '-              already open and running.                   -
        '------------------------------------------------------------

        'Create list of employees
        Dim myEmps As List(Of clsEmployee) = New List(Of clsEmployee)

        'Add employees to list
        myEmps.Add(New clsEmployee("Sue", 103, 15.25, {8, 8, 8, 8, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Scott", 105, 15.0, {10, 10, 0, 10, 10, 10, 0}))
        myEmps.Add(New clsEmployee("Bill", 106, 12.0, {8, 8, 8, 8, 9, 0, 0}))
        myEmps.Add(New clsEmployee("Tina", 107, 16.0, {8, 8, 8, 8, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Ron", 109, 15.5, {0, 0, 9, 9, 9, 9, 9}))
        myEmps.Add(New clsEmployee("Barb", 110, 13.0, {0, 10, 0, 10, 10, 10, 0}))
        myEmps.Add(New clsEmployee("Cathy", 111, 14.5, {8, 8, 8, 8, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Al", 112, 15.0, {10, 10, 10, 10, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Dave", 133, 15.5, {0, 0, 8, 8, 8, 8, 8}))
        myEmps.Add(New clsEmployee("Haley", 134, 16.5, {8, 8, 8, 8, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Drew", 136, 12.25, {10, 10, 0, 0, 10, 10, 0}))
        myEmps.Add(New clsEmployee("John", 137, 13.0, {8, 8, 8, 8, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Mary", 138, 14.0, {8, 8, 8, 8, 8, 0, 0}))
        myEmps.Add(New clsEmployee("Ann", 139, 15.0, {0, 0, 0, 10, 10, 10, 10}))
        myEmps.Add(New clsEmployee("Chuck", 140, 15.0, {0, 8, 8, 8, 8, 8, 0}))

        'Create Excel sheet
        Dim aSheet As Excel.Application
        Dim CheckExcel As Object

        'See if Excel is already in memory. If it isn't, we got an error
        Try
            CheckExcel = GetObject(, "Excel.Application")

            'It is, so show it
            aSheet = CheckExcel
            aSheet.Visible = True
        Catch Ex As Exception
            'Create a new instance of Excel
            aSheet = New Excel.Application()
            aSheet.Visible = True
        End Try

        'Add a new workbook and a new sheet
        aSheet.Workbooks.Add()
        aSheet.Sheets.Add()

        'Write headers to Excel sheet
        WriteHeaders(aSheet)

        'Write data to Excel sheet
        WriteData(aSheet, myEmps)

        'Write stats to Excel sheet
        WriteStats(aSheet, myEmps.Count + 1)

        'Clean things up
        aSheet = Nothing
    End Sub

    Sub WriteHeaders(aSheet As Excel.Application)

        '------------------------------------------------------------
        '-              Subprogram Name: WriteHeaders               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: April 30, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine writes the column headers to the Excel   -
        '- sheet.                                                   -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- aSheet:      Excel application to which write data.      -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- arrHeaders:  String array that holds the column headers. -
        '------------------------------------------------------------

        'Create header array
        Dim arrHeaders() As String = {"Name", "ID", "Payrate", "Hours", "Total"}

        'Write each to Excel
        For i = 1 To intCOLUMNCOUNT
            aSheet.Cells(1, i) = arrHeaders(i - 1)
        Next
    End Sub

    Sub WriteData(aSheet As Excel.Application, myEmps As List(Of clsEmployee))

        '------------------------------------------------------------
        '-              Subprogram Name: WriteHeaders               -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: April 30, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine writes the employee data to the Excel    -
        '- sheet.                                                   -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- myEmps:      List of clsEmployees.                       -
        '- aSheet:      Excel application to which write data.      -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- intRow:      Integer variable to keep track of what row  -
        '-              which is currently being written.           -
        '------------------------------------------------------------

        'int to keep track of current row
        Dim intRow As Integer

        'Write each line of data
        For i = 0 To myEmps.Count - 1

            intRow = i + intOFFSET

            'Write Name
            aSheet.Cells(intRow, 1) = myEmps(i).strName

            'Write ID
            aSheet.Cells(intRow, 2) = myEmps(i).intID

            'Write Payrate
            aSheet.Cells(intRow, 3) = myEmps(i).sngRate

            'Write Hours
            aSheet.Cells(intRow, 4) = myEmps(i).TotalHours

            'Write Total (including calculation for overtime)
            aSheet.Cells(intRow, 5) = "=IF((D" & intRow & " > 40), ((C" & intRow & " * 40) + ((D" & intRow & " - 40) * C" & intRow & " * 1.5)), C" & intRow & " * D" & intRow & ")"
        Next
    End Sub

    Sub WriteStats(aSheet As Excel.Application, intSize As Integer)

        '------------------------------------------------------------
        '-              Subprogram Name: WriteStats                 -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: April 30, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine writes the employee statistics to the    -
        '- Excel sheet.                                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- intSize:     Integer to keep track of the size of the    -
        '-              data rows.                                  -
        '- aSheet:      Excel application to which write data.      -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- intRow:      Integer variable to keep track of what row  -
        '-              which is currently being written.           -
        '------------------------------------------------------------

        Dim intRow As Integer = intSize + intOFFSET

        'Write average
        aSheet.Cells(intRow, 2) = "Aver:"
        aSheet.Cells(intRow, 3) = "=ROUND(AVERAGE(C2:C" & intSize & "), 1)"
        aSheet.Cells(intRow, 4) = "=ROUND(AVERAGE(D2:D" & intSize & "), 1)"
        aSheet.Cells(intRow, 5) = "=ROUND(AVERAGE(E2:E" & intSize & "), 2)"

        'Move to next row
        intRow += 1

        'Write min
        aSheet.Cells(intRow, 2) = "Min:"
        aSheet.Cells(intRow, 3) = "=ROUND(MIN(C2:C" & intSize & "), 1)"
        aSheet.Cells(intRow, 4) = "=ROUND(MIN(D2:D" & intSize & "), 1)"
        aSheet.Cells(intRow, 5) = "=ROUND(MIN(E2:E" & intSize & "), 2)"

        'Move to next row
        intRow += 1

        'Write max
        aSheet.Cells(intRow, 2) = "Max:"
        aSheet.Cells(intRow, 3) = "=ROUND(MAX(C2:C" & intSize & "), 1)"
        aSheet.Cells(intRow, 4) = "=ROUND(MAX(D2:D" & intSize & "), 1)"
        aSheet.Cells(intRow, 5) = "=ROUND(MAX(E2:E" & intSize & "), 2)"

        'Move to next row
        intRow += 1

        'Write total
        aSheet.Cells(intRow, 2) = "Total:"
        aSheet.Cells(intRow, 3) = "=ROUND(SUM(C2:C" & intSize & "), 1)"
        aSheet.Cells(intRow, 4) = "=ROUND(SUM(D2:D" & intSize & "), 1)"
        aSheet.Cells(intRow, 5) = "=ROUND(SUM(E2:E" & intSize & "), 2)"
    End Sub
End Module