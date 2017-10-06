Imports System.IO
Imports OfficeOpenXml
Imports System.Collections.Generic
Imports System.Text
Imports System.Drawing
Imports System

Public Class GlobalFunction
    Public Property Response As Object

    Function GetIP() As DataTable
        Dim file As String = Application.StartupPath & "\ip.txt"
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim i As Integer = 0
        Dim sr As StreamReader = New StreamReader(file)
        Dim line As String = ""
        Do While sr.Peek() >= 0
            line = sr.ReadLine()
            If i = 0 AndAlso line <> "" Then
                Dim strColumn() As String = line.Split("|")
                For j As Integer = 0 To strColumn.Length - 1
                    dt.Columns.Add(strColumn(j))
                Next
            Else
                Dim strValue() As String = line.Split("|")
                dr = dt.NewRow
                If strValue.Length = dt.Columns.Count Then
                    For j As Integer = 0 To strValue.Length - 1
                        dr(j) = strValue(j)
                    Next
                    If dr(0).ToString <> "" Then
                        dt.Rows.Add(dr)
                    End If
                End If
            End If

            i = i + 1
        Loop
        sr.Close()
        Return dt
    End Function

    Public Function GetDateTime() As String
        Return DateTime.Now.ToString("yyyyMMdd hh:MM:ss")
    End Function

#Region "Export Excel"
    Protected Sub SetHeaderStyle(RowHeader As ExcelRange)
        RowHeader.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
        RowHeader.Style.Font.Bold = True
        RowHeader.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
        RowHeader.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#2f6073"))
        RowHeader.Style.Font.Color.SetColor(Color.White)
    End Sub

    Protected Function GetDataForReport() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("firstname")
        dt.Columns.Add("lastname")
        dt.Columns.Add("address")
        dt.Columns.Add("telno")

        Dim dr As DataRow = dt.NewRow()
        dr("firstname") = "นายก"
        dr("lastname") = "ใจดี"
        dr("address") = "ห้วยขวาง"
        dr("telno") = "01234567890"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("firstname") = "นายข"
        dr("lastname") = "ใจดีมาก"
        dr("address") = "บางกะปิ"
        dr("telno") = "01234567891"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("firstname") = "นายค"
        dr("lastname") = "ใจดีมากกว่า"
        dr("address") = "ปิ่นเกล้า"
        dr("telno") = "01234567892"
        dt.Rows.Add(dr)

        Return dt

    End Function

    Public Sub ExportToExcel()
        'Dim xlPackage As New ExcelPackage()
        'Dim oBook As ExcelWorkbook = xlPackage.Workbook
        'For Each dt As System.Data.DataTable In ds.Tables
        '    Dim ws As ExcelWorksheet = oBook.Worksheets.Add(dt.TableName)
        '    ws.Cells.LoadFromDataTable(dt, True)
        '    ws.Cells.Style.Locked = False
        'Next

        Dim dt As DataTable = GetDataForReport()
        Dim ep As New ExcelPackage()
        Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add("TestReport")

        Dim cnt As Integer = 1
        ws.Cells("A" & cnt & "").Value = "ชื่อ"
        SetHeaderStyle(ws.Cells("A" & cnt & ""))

        ws.Cells("B" & cnt & "").Value = "สกุล"
        SetHeaderStyle(ws.Cells("B" & cnt & ""))

        ws.Cells("C" & cnt & "").Value = "ที่อยู่"
        SetHeaderStyle(ws.Cells("C" & cnt & ""))

        ws.Cells("D" & cnt & "").Value = "เบอร์โทร"
        SetHeaderStyle(ws.Cells("D" & cnt & ""))

        For i As Integer = 0 To dt.Rows.Count - 1
            cnt += 1
            Dim _firstname As String = dt.Rows(i)("firstname").ToString()
            Dim _lastname As String = dt.Rows(i)("lastname").ToString()
            Dim _address As String = dt.Rows(i)("address").ToString()
            Dim _telno As String = dt.Rows(i)("telno").ToString()

            ws.Cells("A" & cnt & "").Value = _firstname
            ws.Cells("B" & cnt & "").Value = _lastname
            ws.Cells("C" & cnt & "").Value = _address
            ws.Cells("D" & cnt & "").Value = _telno
        Next

        Dim RowContent As ExcelRange = ws.Cells("A" & cnt & ":D" & cnt)
        RowContent.AutoFitColumns()

        Dim path As String = Application.StartupPath & "\Export"
        If Directory.Exists(path) = False Then Directory.CreateDirectory(path)
        Dim filename As String = DateTime.Now.ToString("yyyyMMddhhmmssfff") + ".xlsx"
        Dim excelFile As New FileInfo(path & "\" & filename)
        ep.SaveAs(excelFile)

        'Dim csvdata As Byte() = ConvertToCsv(ep)

        'Dim path As String = Application.StartupPath & "\Export"
        'If Directory.Exists(path) = False Then Directory.CreateDirectory(path)
        'Dim filename As String = DateTime.Now.ToString("yyyyMMddhhmmssfff") + ".csv"
        'Dim excelFile As New FileInfo(path & "\" & filename)
        'ep.SaveAs(excelFile)

        'File.AppendAllText(path & "\" & filename, BitConverter.ToString(csvdata))

    End Sub

    Public Function ConvertToCsv(package As ExcelPackage) As Byte()
        Dim worksheet = package.Workbook.Worksheets(1)

        Dim maxColumnNumber = worksheet.Dimension.[End].Column
        Dim currentRow = New List(Of String)(maxColumnNumber)
        Dim totalRowCount = worksheet.Dimension.[End].Row
        Dim currentRowNum = 1

        Dim memory = New MemoryStream()

        Using writer = New StreamWriter(memory, Encoding.ASCII)
            While currentRowNum <= totalRowCount
                BuildRow(worksheet, currentRow, currentRowNum, maxColumnNumber)
                WriteRecordToFile(currentRow, writer, currentRowNum, totalRowCount)
                currentRow.Clear()
                currentRowNum += 1
            End While
        End Using

        Return memory.ToArray()
    End Function

    Sub WriteRecordToFile(record As List(Of String), sw As StreamWriter, rowNumber As Integer, totalRowCount As Integer)
        Dim commaDelimitedRecord = ToDelimitedString(record, ",") 'record.ToDelimitedString(",")

        If rowNumber = totalRowCount Then
            sw.Write(commaDelimitedRecord)
        Else
            sw.WriteLine(commaDelimitedRecord)
        End If
    End Sub

    Sub BuildRow(worksheet As ExcelWorksheet, currentRow As List(Of String), currentRowNum As Integer, maxColumnNumber As Integer)
        For i As Integer = 1 To maxColumnNumber
            Dim cell = worksheet.Cells(currentRowNum, i)
            If cell Is Nothing Then
                ' add a cell value for empty cells to keep data aligned.
                AddCellValue(String.Empty, currentRow)
            Else
                AddCellValue(GetCellText(cell), currentRow)
            End If
        Next
    End Sub

    Function GetCellText(cell As ExcelRangeBase) As String
        Return If(cell.Value Is Nothing, String.Empty, cell.Value.ToString())
    End Function

    Sub AddCellValue(s As String, record As List(Of String))
        record.Add(String.Format("{0}{1}{0}", """"c, s))
    End Sub

    Function ToDelimitedString(list As List(Of String), Optional delimiter As String = ":", Optional insertSpaces As Boolean = False, Optional qualifier As String = "", Optional duplicateTicksForSQL As Boolean = False) As String
        Dim result = New StringBuilder()
        For i As Integer = 0 To list.Count - 1
            Dim initialStr As String = If(duplicateTicksForSQL, list(i).Replace("'", "''"), list(i))
            result.Append(If((qualifier = String.Empty), initialStr, String.Format("{1}{0}{1}", initialStr, qualifier)))
            If i < list.Count - 1 Then
                result.Append(delimiter)
                If insertSpaces Then
                    result.Append(" "c)
                End If
            End If
        Next
        Return result.ToString()
    End Function

#End Region
End Class
