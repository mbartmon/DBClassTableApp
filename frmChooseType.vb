Public Class frmChooseType

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        isSQL = False
        Dim frm As New frmMain
        frm.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click

        SetDbConnection(1)
        isSQL = True

        If Not IsNothing(sqlCN) Then
            Dim frm As New frmMainSQL
            frm.ShowDialog()
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As System.EventArgs) Handles Button3.Click

        Application.Exit()

    End Sub

    Private Sub Button4_Click(sender As Object, e As System.EventArgs) Handles Button4.Click

        SetDbConnection(2)
        isSQL = True

        If Not IsNothing(sqlCN) Then
            Dim frm As New frmMainSQL
            frm.ShowDialog()
        End If

    End Sub
End Class