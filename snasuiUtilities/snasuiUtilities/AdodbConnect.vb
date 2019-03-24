Imports Microsoft.Office.Interop.Excel
Module AdodbConnect
    '!!! This method must be add ActiveX data object x.x library first.
    Sub TestingConnectionSqlConnection()
        Dim wb As Workbook = xlApp.ActiveWorkbook
        Dim sh As Worksheet = wb.ActiveSheet
        Dim rng As Range = sh.Range("a2")
        Dim myCon As New ADODB.Connection
        Dim myRecset As New ADODB.Recordset
        Dim cmd As New ADODB.Command
        Dim i As Integer, j As Integer

        '1. Initial connection string
        myCon.ConnectionString = "Provider = SQLNCLI11;" &
                "Data Source=localhost;" &
                "Initial Catalog=AdventureWorks2012;" &
                "User ID=sa;" &
                "Password=s@ntip0ng;"

        '2 SQL Statement
        Dim strsql = "Select Top 5 * From [AdventureWorks2012].[Sales].[Customer]"

        '3. Open connection
        myCon.Open()

        '4. Assign connection and sql statement to command object
        cmd.ActiveConnection = myCon
        cmd.CommandText = strsql

        '5. Execute data and keep in recordset
        myRecset = cmd.Execute

        '6. Paste data from recordset to destination
        If myRecset.Fields.Count = 0 Then
            MsgBox("Data not found.", vbInformation)
        Else
            For j = 0 To myRecset.Fields.Count - 1
                rng.Offset(-1, i).Value = myRecset.Fields(j).Name
                i = i + 1
            Next
            rng.CopyFromRecordset(myRecset)
            rng.CurrentRegion.EntireColumn.AutoFit()
        End If

        '7. Clear memory on stystem
        myCon.Close()
        myCon = Nothing
    End Sub
End Module
