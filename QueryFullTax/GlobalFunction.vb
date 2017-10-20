Imports System.IO
Imports System.Text

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
        Return DateTime.Now.ToString("yyyyMMdd H:mm:ss")
    End Function

    Public Function ExportToCSV(ip As String, dt As DataTable) As String
        Dim ret As String = ""
        Try
            Dim data As New StringBuilder
            Dim column_name As String = ""
            Dim j As Integer = 0
            For Each column As DataColumn In dt.Columns
                column_name &= String.Format("{0}{1}{0}", """"c, column.ColumnName) & ";"
                j += 1
            Next
            data.AppendLine(" " & column_name)

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim strData As String = ConvertTextFormat(dt.Rows(i))
                data.AppendLine(" " & strData)
            Next

            Dim path As String = Application.StartupPath & "\Export"
            Dim output As String = path & "\" & ip & ".csv"
            Dim csv As New StreamWriter(New FileStream(output, FileMode.CreateNew), Encoding.UTF8)
            csv.Write(Data.ToString)
            csv.Close()

        Catch ex As Exception
            ret = ex.ToString
        End Try
        Return ret
    End Function

    Function ConvertTextFormat(dr As DataRow) As String
        Dim ret As String = ""
        For i As Integer = 0 To dr.ItemArray.Count - 1
            Dim str As String = IIf(IsDBNull(dr(i)), "", dr(i).ToString.Replace(",", ""))
            ret &= String.Format("{0}{1}{0}", """"c, str) & ";"
        Next
        Return ret
    End Function

End Class
