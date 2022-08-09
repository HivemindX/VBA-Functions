Option Explicit

Function GetColumnNumber(sColumnName As String) As Long

'#### Input is a column name like A or BZ, the output is the column number that matches the column name lile 1 or 78

    Dim lResult As Long

    On Error GoTo labelError

    If Len(sColumnName) = 1 Then
        ' If the column name is only a single character it is trivial to determine the result using ASCII codes (A=65, etc)
        lResult = Asc(sColumnName) - 64
    Else
        ' If the column name is more than one character (eg: XFD the current max allowed) we determine the number indicated
        ' by the character in the first position and then recursively call the function with the remaining characters
        ' For example if the address is CA we first determine that the value indicated by the C is 78 and then add the value
        ' given by A which is 1 for a result of 79
        lResult = 26 ^ (Len(sColumnName) - 1) * GetColumnNumber(Left(sColumnName, 1)) + GetColumnNumber(Right(sColumnName, Len(sColumnName) - 1))
    End If

    GetColumnNumber = lResult

labelEnd:
    Exit Function

labelError:
    Resume labelEnd

End Function

End Function

Function GetLastCol(ws As Worksheet, Optional lRow As Long) As Long

    Dim lResult As Long
    Dim lLastCol As Long
    
    On Error GoTo labelError
    
    If lRow = 0 Then
        For lRow = 1 To GetLastRow(ws)
            lLastCol = ws.Cells(lRow, ws.Columns.Count).End(xlToLeft).Column
            If lLastCol > lResult Then
                lResult = lLastCol
            End If
        Next lRow
    Else
        lResult = ws.Cells(lRow, ws.Columns.Count).End(xlToLeft).Column
    End If
    
    GetLastCol = lResult
labelEnd:
    Exit Function

labelError:
    Resume labelEnd

End Function

Function GetLastRow(ws As Worksheet, Optional lCol As Long) As Long

    Dim lResult As Long
    Dim lLastRow As Long
    
    On Error GoTo labelError
    
    If lCol = 0 Then
        For lCol = 1 To ws.Columns.Count
            lLastRow = ws.Cells(ws.Rows.Count, lCol).End(xlUp).row
            If lLastRow > lResult Then
                lResult = lLastRow
            End If
        Next lCol
    Else
        lResult = ws.Cells(ws.Rows.Count, lCol).End(xlUp).row
    End If
    
    GetLastRow = lResult
labelEnd:
    Exit Function

labelError:
    Resume labelEnd

End Function