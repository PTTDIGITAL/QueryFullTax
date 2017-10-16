﻿Imports System.IO
Public Class frmMain

    Dim clsGlobalFunction As New GlobalFunction
    Dim clsConnectDatabase As New ConnectDatabase

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtQuery.Focus()
        ProgressBar1.Value = 0
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtQuery.Text = String.Empty
        txtResult.Text = String.Empty
        ProgressBar1.Value = 0
    End Sub

    Private Sub btnExcute_Click(sender As Object, e As EventArgs) Handles btnExcute.Click
        Try
            Dim path As String = Application.StartupPath & "\Export"
            If Directory.Exists(path) Then
                For Each _file As String In Directory.GetFiles(path)
                    File.Delete(_file)
                Next
            Else
                Directory.CreateDirectory(path)
            End If

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

                Application.DoEvents()
                Dim ret As New ProcessReturnInfo
                ret = clsConnectDatabase.FillData(sql)
                If ret.IsSuccess = False Then
                    txtResult.Text &= GetDateTime() & " IP Address" & ip & " พบปัญหาในการเชื่อมต่อ " & ret.ErrorMessage & vbCrLf
                Else
                    txtResult.Text &= GetDateTime() & " IP Address " & ip & " สำเร็จ" & vbCrLf
                    Dim dtdata As New DataTable
                    dtdata = ret.DT
                    If Not dtdata Is Nothing AndAlso dtdata.Rows.Count > 0 Then
                        'Write To CSV
                        clsGlobalFunction.ExportToCSV(ip, dtdata)
                    Else
                        txtResult.Text &= GetDateTime() & " IP Address " & ip & " ไม่พบข้อมูล" & vbCrLf
                    End If
                End If
                If ProgressBar1.Value >= 97 Then ProgressBar1.Value = 97 Else ProgressBar1.Value = ProgressBar1.Value + 1
            Next

            Application.DoEvents()
            txtResult.Text &= GetDateTime() & " สิ้นสุดการเชื่อมต่อ" & vbCrLf
            ProgressBar1.Value = 100
        Catch ex As Exception
            Using New Centered_MessageBox(Me)
                MessageBox.Show("พบปัญหา  " & ex.ToString(), "", MessageBoxButtons.OK)
            End Using
        End Try
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
