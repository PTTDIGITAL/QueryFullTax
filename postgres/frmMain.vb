Public Class frmMain

    Dim clsGlobalFunction As New GlobalFunction
    Dim clsConnectDatabase As New ConnectDatabase

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtQuery.Focus()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtQuery.Text = String.Empty
        txtResult.Text = String.Empty
    End Sub

    Private Sub btnExcute_Click(sender As Object, e As EventArgs) Handles btnExcute.Click
        Dim dtip As DataTable
        dtip = clsGlobalFunction.GetIP()
        For i As Integer = 0 To dtip.Rows.Count - 1
            Dim ip As String = dtip.Rows(i)("ip").ToString
            With clsConnectDatabase
                .Server = ip
                .Database = dtip.Rows(i)("db").ToString
                .Username = dtip.Rows(i)("username").ToString
                .Password = dtip.Rows(i)("password").ToString
                .Port = dtip.Rows(i)("port").ToString
            End With

            Dim sql As String = txtQuery.Text.Trim

            Dim ret As New ProcessReturnInfo
            ret = clsConnectDatabase.FillData(sql)
            If ret.IsSuccess = False Then
                txtResult.Text &= txtResult.Text & GetDateTime() & " IP Address" & ip & " Fail" & ret.ErrorMessage & vbCrLf
            Else
                txtResult.Text &= txtResult.Text & GetDateTime() & " IP Address " & ip & " Success" & vbCrLf
                Dim dtdata As New DataTable
                dtdata = ret.DT
                If Not dtdata Is Nothing AndAlso dtdata.Rows.Count > 0 Then
                    'Write To CSV
                End If
            End If
        Next
    End Sub

    Function GetDateTime()
        Return clsGlobalFunction.GetDateTime()
    End Function

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Using New Centered_MessageBox(Me)
            If MessageBox.Show("ต้องการปิดโปรแกรมใช่หรือไม่", "", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                MyBase.OnFormClosing(e)
            Else
                e.Cancel = True
            End If
        End Using
    End Sub
End Class
