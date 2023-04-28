Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Microsoft.Office.Interop

Module Module1
    Public myconn As New MySql.Data.MySqlClient.MySqlConnection
    Public myConnectionString As String
    Public strSQL As String

    Public currentDate As DateTime = DateTime.Now
    Public strpassword = "mariahannah25"
    Public xlsFiles As String = "C:\Users\hanna\OneDrive\Documents\Event Driven Programming (EDP)\"

    Public Sub Connect2DB()
        myConnectionString = "server=localhost;" _
                & "user id=root;" _
                & "port=3306;" _
                & "password=mariahannah25;" _
                & "database=database"
        Try
            myconn.ConnectionString = myConnectionString
            myconn.Open()

        Catch ex As MySqlException
            Select Case ex.Number
                Case 0
                    MsgBox("Cannot connect to the server!")
                Case 1045
                    MsgBox("Invalid username or password")
            End Select

        End Try
    End Sub

    Public Sub Disconnect2DB()
        myconn.Close()
        myconn.Dispose()

    End Sub


    Public Sub ExportToExcel(ByVal dgv As DataGridView, ByVal templatefilename As String)
        Dim xlsApp As Excel.Application = Nothing
        Dim xlsWB As Excel.Workbook = Nothing
        Dim xlsSheet As Excel.Worksheet = Nothing
        Dim misValue As Object = System.Reflection.Missing.Value

        xlsApp = New Excel.Application
        xlsApp.Visible = False
        xlsWB = xlsApp.Workbooks.Add(misValue)
        xlsSheet = xlsWB.Sheets("Sheet1")

        'Copy DataGridView values to Excel
        Dim columnsCount As Integer = dgv.Columns.Count
        For column As Integer = 0 To columnsCount - 1
            xlsSheet.Cells(1, column + 1) = dgv.Columns(column).HeaderText
            For row As Integer = 0 To dgv.Rows.Count - 1
                If dgv.Rows(row).Cells(column).Value IsNot Nothing Then
                    xlsSheet.Cells(row + 2, column + 1) = dgv.Rows(row).Cells(column).Value.ToString()
                Else
                    xlsSheet.Cells(row + 2, column + 1) = ""
                End If
            Next
        Next


        'Autofit columns in Excel
        Dim columnsRange As Range = xlsSheet.Columns
        columnsRange.AutoFit()

        'Save the Excel file
        templatefilename = templatefilename.Replace(".xlsx", "")
        templatefilename = templatefilename.Replace(".xls", "")
        Dim myfilename As String = templatefilename & " " & currentDate.ToString("mm-dd-yy hh-mm-ss") & ".xlsx"
        xlsSheet.Protect(strpassword)
        xlsApp.ActiveWindow.View = XlWindowView.xlPageLayoutView
        xlsApp.ActiveWindow.DisplayGridlines = False
        xlsWB.SaveAs(xlsFiles & myfilename, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlsWB.Close(True, misValue, misValue)
        xlsApp.Quit()

        releaseObject(xlsSheet)
        releaseObject(xlsWB)
        releaseObject(xlsApp)

        Dim filePath As String = Path.Combine(xlsFiles, myfilename)

        'Open the Excel file
        'System.Diagnostics.Process.Start("excel.exe", """" & xlsFiles & myfilename & """")
        MsgBox("File saved to " + filePath)
    End Sub

    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module

