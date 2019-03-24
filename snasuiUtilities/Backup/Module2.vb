Imports Microsoft.Office.Interop.Excel

Module Module2
    Public Sub DeleteAllNonBuildInStyles()
        Dim sty As Style, wb As Workbook, i As Integer
        'wb = Globals.ThisAddIn.Application.ActiveWorkbook
        wb = xlApp.ActiveWorkbook
        For Each sty In wb.Styles
            If Not sty.BuiltIn Then
                sty.Delete()
                i = i + 1
            End If
        Next
        If i > 0 Then
            MsgBox("Total " & i & " styles are deleted.", vbInformation)
        Else
            MsgBox("Not found none build-in style.", vbInformation)
        End If
    End Sub
    Public Sub DeleteNamedRangesError()
        Dim n As Name, j As Long
        j = 0
        For Each n In xlApp.ActiveWorkbook.Names
            If InStr(n.RefersTo, "#REF!") Then
                n.Delete()
                j = j + 1
            End If
        Next
        If j > 0 Then
            MsgBox("Total " & j & " name(s) are deleted", vbIgnore)
        Else
            MsgBox("Not found eror name.", vbInformation)
        End If
    End Sub

    Public Sub DeleteAllShapes()
        Dim shps As Shape, k As Long
        Dim wb As Workbook = xlApp.ActiveWorkbook
        Dim sh As Worksheet
        For Each sh In wb.Worksheets
            For Each shps In sh.Shapes
                shps.Delete()
                k = k + 1
            Next
        Next
        If k > 0 Then
            MsgBox("Total " & k & " shape(s) are deleted.", vbInformation)
        Else
            MsgBox("Not found shape", vbInformation)
        End If
    End Sub
End Module
