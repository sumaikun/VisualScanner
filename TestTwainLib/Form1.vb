Public Class Form1


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FileNames = TwainLib.ScanImages(, , My.Settings.scan)

        For Each Itm In FileNames
            ListBox1.Items.Add(Itm)
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        My.Settings.scan = TwainLib.GetScanSource
        My.Settings.Save()
    End Sub
End Class
