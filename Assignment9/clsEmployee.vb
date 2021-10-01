Public Class clsEmployee

    '------------------------------------------------------------
    '-                File Name : clsEmployee.vb                -
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
    '- Global Variable Dictionary (alphabetically):             -
    '- intID:       Integer variable that holds the employee's  -
    '-              identification number.                      -
    '- arrHours:    Single array that holds the employee's      -
    '-              hours worked for each day of the week.      -
    '- sngRate:     Single variable that holds teh employee's   -
    '-              pay rate.                                   -
    '- strName:     String variable that holds the employee's   -
    '-              name.                                       -
    '------------------------------------------------------------

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    Const intWEEKLYDAYS As Integer = 6

    '---------------------------------------------------------------------------------------
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '---------------------------------------------------------------------------------------

    Public strName As String
    Public intID As Integer
    Public sngRate As Single
    Public arrHours(6) As Single

    '-----------------------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-----------------------------------------------------------------------------------

    Public Sub New(ByVal newName As String, ByVal newID As Integer, ByVal newRate As Single, ByVal newHours() As Single)

        '------------------------------------------------------------
        '-              Subprogram Name: New                        -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: April 30, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine is a named constructor that instantiates -
        '- a new instance of an employee                            -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- newID:       Integer that holds the new employee's ID.   -
        '- newHours:    Single array that holds the new employee's  -
        '-              hours worked.                               -
        '- newName:     String that holds the new employee's name.  -
        '- newRate:     Single that holds the new employee's pay    -
        '-              rate.                                       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        strName = newName
        intID = newID
        sngRate = newRate
        arrHours = newHours
    End Sub

    Function TotalHours() As Single

        '------------------------------------------------------------
        '-              Subprogram Name: TotalHours                 -
        '------------------------------------------------------------
        '-                Written By: Benjamin Neeb                 -
        '-                Written On: April 30, 2021                -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '- This Subroutine adds the employee's hours for each day   -
        '- to return a sum as a single.                             -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- sngSum:      Single to hold the employee's total worked  -
        '-              hours.                                      -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- sngSum:      The subtotal amount of the employee's total -
        '-              hours worked for the week.                  -
        '------------------------------------------------------------

        Dim sngSum As Single = 0

        'Get each number in the array and add them to the integer
        For i = 0 To intWEEKLYDAYS
            sngSum += arrHours(i)
        Next

        'Return the sum
        Return sngSum
    End Function
End Class
