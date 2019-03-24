Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.Windows.Forms

Module SubProc
    Public Sub DeleteAllNonBuildInStyles()
        Dim sty As Style, wb As Workbook, i As Integer
        If xlApp.Workbooks.Count = 0 Then Exit Sub
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
        If xlApp.Workbooks.Count = 0 Then Exit Sub
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
        If xlApp.Workbooks.Count = 0 Then Exit Sub
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

    Public Sub DeleteBlankRows()
        Dim rng As Range, k As Long, j As Long
        If xlApp.Workbooks.Count = 0 Then Exit Sub
        Dim wb As Workbook = xlApp.ActiveWorkbook
        Dim sh As Worksheet
        For Each sh In wb.Worksheets
            For j = sh.UsedRange.Resize(, 1).Cells.Count To 1 Step -1
                rng = sh.UsedRange.Resize(, 1).Cells(j)
                If Len(rng.Value) = 0 And rng.End(XlDirection.xlToRight).Column = sh.Columns.Count Then
                    rng.EntireRow.Delete()
                    k = k + 1
                End If
            Next
        Next
        If k > 0 Then
            MsgBox("Total " & k & " rows(s) are deleted.", vbInformation)
        Else
            MsgBox("Not found blank row", vbInformation)
        End If
    End Sub

    Public Sub ResetLastCellAllSheets()
        If xlApp.Workbooks.Count = 0 Then Exit Sub
        Dim wb As Workbook = xlApp.ActiveWorkbook
        Dim sh As Worksheet, x As Long, y As Long
        Try
            For Each sh In wb.Worksheets
                With sh.UsedRange
                    x = .Rows.Count
                    y = .Columns.Count
                End With
            Next
            wb.Save()
        Catch ex As Exception
        End Try
        MsgBox("Finished.", vbInformation)
    End Sub

    Public Sub ListErrorFormula()
        Dim nwb As Workbook, tb As Workbook = xlApp.ActiveWorkbook
        Dim sh As Worksheet, arr(99999, 2) As Object
        Dim r As Range, l As Long, tgSh As Worksheet
        l = 0
        For Each sh In tb.Worksheets
            For Each r In sh.UsedRange
                If xlApp.WorksheetFunction.IsError(r) Then
                    arr(l, 0) = sh.Name
                    arr(l, 1) = r.Address
                    arr(l, 2) = "'" & r.Formula
                    l = l + 1
                End If
            Next r
        Next sh
        If l > 0 Then
            nwb = xlApp.Workbooks.Add
            tgSh = nwb.Worksheets(1)
            With tgSh
                .Range("a1:c1").Value = Split("SheetName|Cell|Formulas", "|")
                .Range("a2").Resize(l, 3).Value = arr
            End With
        Else
            MsgBox("Not found error formula.", vbInformation)
        End If
    End Sub

    Sub ImportFilesAndSheets()
        Dim strFile As Array, i As Integer
        Dim tb As Workbook, wb As Workbook
        Dim sh As Worksheet, tgSh As Worksheet, lstRow As Long
        tb = xlApp.ActiveWorkbook
        tgSh = tb.Sheets(1)
        xlApp.ScreenUpdating = False
        xlApp.Calculation = XlCalculation.xlCalculationManual
        Try
            tgSh.UsedRange.ClearContents()
            strFile = xlApp.GetOpenFilename(FileFilter:="Excel Files (*.xls*),*.xls*",
                                        MultiSelect:=True)
            If strFile Is Nothing Then
                Exit Sub
            End If
            For i = 1 To UBound(strFile)
                wb = xlApp.Workbooks.Open(strFile(i))
                For Each sh In wb.Worksheets
                    If sh.UsedRange.Rows.Count > 0 Then
                        lstRow = tgSh.Range("a" & tgSh.Rows.Count).End(XlDirection.xlUp).Offset(1, 0).Row
                        sh.UsedRange.Copy()
                        If tgSh.Range("a1").Value = "" Then
                            tgSh.Range("a1").PasteSpecial(XlPasteType.xlPasteValues)
                        Else
                            tgSh.Range("a" & lstRow).PasteSpecial(XlPasteType.xlPasteValues)
                        End If
                        xlApp.CutCopyMode = False
                    End If
                Next
                wb.Close(SaveChanges:=False)
            Next
        Catch ex As Exception
        End Try
        xlApp.ScreenUpdating = True
        xlApp.Calculation = XlCalculation.xlCalculationSemiautomatic
    End Sub

    Public Sub LinkFromExternal()
        Dim j As Long, twb As Workbook = xlApp.ActiveWorkbook
        Dim nwb As Workbook = xlApp.Workbooks.Add
        Dim lng As Long
        Dim rng As Range, tgSh As Worksheet = nwb.Worksheets(1)
        tgSh.Range("a1:c1").Value = {"SheetName", "Cell", "Formula"}
        For Each sh As Worksheet In twb.Worksheets
            For Each rng In sh.UsedRange
                If rng.HasFormula And InStr(rng.Formula, "\[") Then
                    With tgSh
                        lng = .Range("a" & .Rows.Count).End(XlDirection.xlUp).Row + 1
                        .Range("a" & lng).Value = sh.Name
                        .Range("b" & lng).Value = rng.Address(0, 0)
                        .Range("c" & lng).Value = "'" & rng.Formula
                        j = j + 1
                    End With
                End If
            Next
        Next
        If j > 0 Then
            MsgBox("Found " & j & " items.", vbInformation)
        Else
            MsgBox("Not fonnd exoternal links.", vbInformation)
        End If
    End Sub

    Public Sub ListAllSheets()
        Dim wb As Workbook = xlApp.ActiveWorkbook
        Dim sh As Worksheet, tgSh As Worksheet
        Dim nwb As Workbook = xlApp.Workbooks.Add
        Dim i As Integer
        tgSh = nwb.Worksheets(1)
        With tgSh.Range("a1:b1")
            .Value = {"Sheet Name", "Status"}
            .Font.Color = Color.Blue
            .Font.Bold = True
        End With
        For Each sh In wb.Worksheets
            With tgSh.Range("a" & tgSh.Rows.Count).End(XlDirection.xlUp).Offset(1, 0)
                .Value = sh.Name
                Select Case sh.Visible
                    Case XlSheetVisibility.xlSheetVisible
                        .Offset(0, 1).Value = "Visible"
                    Case XlSheetVisibility.xlSheetHidden
                        .Offset(0, 1).Value = "Hidden"
                        .Resize(1, 2).Font.Color = Color.Red
                    Case XlSheetVisibility.xlSheetVeryHidden
                        .Offset(0, 1).Value = "VeryHidden"
                        .Resize(1, 2).Font.Color = Color.Red
                End Select
                i = i + 1
            End With
        Next

        tgSh.Range("a1").CurrentRegion.EntireColumn.AutoFit()

        If i > 0 Then
            MsgBox("Found " & i & " sheets.", vbInformation)
        End If
    End Sub
    Public Sub ResponseForumMsg()
        Dim wb As Workbook = xlApp.ActiveWorkbook
        'Dim nwb As Workbook = xlApp.Workbooks.Add
        'Dim tgSh As Worksheet = nwb.Worksheets(1)
        Dim rng As Range, txtMsg As String = ""
        'wb.Activate()
        For Each rng In wb.Application.Selection
            If rng.HasArray Then
                txtMsg = txtMsg & "[*] ที่ " & rng.Address(0, 0) & " คีย์ " & vbCrLf
                txtMsg = txtMsg & "[font=consolas]" & rng.Formula & "[/font]" & vbCrLf
                txtMsg = txtMsg & "Ctrl+Shift+Enter > Copy ลงด้านล่าง" & vbCrLf
            Else
                txtMsg = txtMsg & "[*] ที่ " & rng.Address(0, 0) & " คีย์ " & vbCrLf
                txtMsg = txtMsg & "[font=consolas]" & rng.Formula & "[/font]" & vbCrLf
                txtMsg = txtMsg & "Enter > Copy ลงด้านล่าง" & vbCrLf
            End If
        Next
        txtMsg = ":D ตัวอย่างสูตรตามด้านล่างครับ " & vbCrLf & "[list=1]" & vbCrLf & txtMsg & "[/list]"
        Clipboard.SetText(txtMsg)
        MsgBox("Please paste to destination.", vbInformation)
        'tgSh.Range("b2").Value = txtMsg
        'tgSh.Range("b2").EntireColumn.AutoFit()
        'nwb.Activate()
    End Sub

    Public Sub InsertBlankrows()
        Dim rngAll As Range, sh As Worksheet = xlApp.ActiveSheet
        Dim currCell As Range, allRows As Integer, j As Integer
        Dim inSertRows As Integer, x As String, y As String
        Dim rowOrCol As Integer, firstCell As Range, firstRow As Integer
        Try
            inSertRows = InputBox("Please enter row(s) for insert")
        Catch ex As Exception
            MsgBox("Please enter number only.", vbExclamation)
            Exit Sub
        End Try
        rowOrCol = MsgBox("Insert entire row click 'Yes', Shift cells down click 'No'",
                        vbYesNo + vbQuestion)
        If Not IsNumeric(inSertRows) Then
            MsgBox("Please enter number only.", vbExclamation)
        End If

        currCell = xlApp.ActiveCell
        allRows = currCell.CurrentRegion.Rows.Count
        firstRow = currCell.CurrentRegion.Range("A1").Row
        firstCell = sh.Cells(firstRow, currCell.Column)

        rngAll = firstCell.Resize(allRows, 1)

        For j = allRows To 2 Step -1
            x = firstCell.Offset(j, 0).Value
            y = firstCell.Offset(j - 1, 0).Value
            If x <> y Then
                If rowOrCol = vbYes Then
                    firstCell.Offset(j, 0).Resize(inSertRows, 1).EntireRow.Insert()
                Else
                    firstCell.Offset(j, 0).Resize(inSertRows, 1).Insert(Shift:=XlInsertShiftDirection.xlShiftDown)
                End If
            End If
        Next
        MsgBox("Finished.", vbInformation)
    End Sub

    Public Sub FillBlankCells()
        Dim rng As Range = Nothing
        Dim sh As Worksheet = xlApp.ActiveSheet
        Dim rngAll As Range = sh.Application.Selection
        Dim rngFillBlanks As Range = Nothing
        xlApp.ScreenUpdating = False
        Try
            If rngAll.Count = 1 Then
                rngFillBlanks = sh.UsedRange.SpecialCells(XlCellType.xlCellTypeBlanks)
            Else
                rngFillBlanks = rngAll.SpecialCells(XlCellType.xlCellTypeBlanks)
            End If
            For Each rng In rngFillBlanks
                If rng.Row <> 1 Then
                    rng.NumberFormat = rng.End(XlDirection.xlUp).NumberFormat
                    rng.Value = rng.End(XlDirection.xlUp)
                End If
            Next
        Catch ex As Exception
            MsgBox("Not found blank cells.", vbInformation)
            xlApp.ScreenUpdating = True
            Exit Sub
        End Try
        xlApp.ScreenUpdating = True
        MsgBox("Finished.", vbInformation)
    End Sub

    Public Sub SplitFile()
        Dim wb As Workbook = Nothing
        Dim sh As Worksheet = Nothing
        Dim fpth As String = Nothing
        fpth = xlApp.GetOpenFilename(FileFilter:="Excel File (*.xls*),*xls*", MultiSelect:=False)
        xlApp.ScreenUpdating = False
        xlApp.Calculation = XlCalculation.xlCalculationManual
        Try
            If fpth = "False" Then Exit Sub
            wb = xlApp.Workbooks.Open(Filename:=fpth, UpdateLinks:=False)
            For Each sh In wb.Worksheets
                sh.Copy()
                xlApp.ActiveWorkbook.SaveAs(Filename:=sh.Name)
                xlApp.ActiveWorkbook.Close(SaveChanges:=False)
            Next
        Catch ex As Exception
        End Try
        wb.Close(SaveChanges:=False)
        xlApp.ScreenUpdating = True
        xlApp.Calculation = XlCalculation.xlCalculationSemiautomatic
        MsgBox("Finished.", vbInformation)
    End Sub

    Public Sub BreakAllLinks()
        Dim allLinks As Object = Nothing
        Dim i As Integer = 0
        Try
            allLinks = xlApp.ActiveWorkbook.LinkSources(Type:=XlLinkType.xlLinkTypeExcelLinks)
            For Each item In allLinks
                xlApp.ActiveWorkbook.BreakLink(Name:=item, Type:=XlLinkType.xlLinkTypeExcelLinks)
                i = i + 1
            Next
            If i > 0 Then
                MsgBox("Finished with " & i & " link(s).", vbInformation)
            End If
        Catch ex As Exception
            MsgBox("Not found link from other file.", vbInformation)
        End Try
    End Sub

    Public Sub OpenLinkOnTag(lnk As String)
        If xlApp.Workbooks.Count = 0 Then Exit Sub
        Dim wb As Workbook = xlApp.ActiveWorkbook
        wb.FollowHyperlink(Address:=lnk)
    End Sub
End Module
