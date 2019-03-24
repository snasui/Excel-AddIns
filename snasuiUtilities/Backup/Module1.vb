Imports Microsoft.Office.Interop.Excel
Module Module1
    'Declare xlApp as Excel Appplication, public variable for using in all module
    Public xlApp As Excel.Application

    Public Function TotalVal(ByRef rng As Range) As Long
        Dim r As Range, l As Long
        l = 0
        For Each r In rng
            If r.Value >= 3 Then
                l = l + r.Value
            End If
        Next
        TotalVal = l
    End Function

End Module
