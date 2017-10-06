Imports System.IO
Public Class GlobalFunction
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
End Class
