Public Class Form1
    Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click
        Dim TestDll As XmlTools.Xml = New XmlTools.Xml
        Dim strResult As String = Nothing

        Select Case cbxFunctions.SelectedItem
            Case "ExtractFromXml"
                txtResult.Text = TestDll.ExtractFromXml(txtElement.Text, txtXmlRequest.Text, txtXmlType.Text)
        End Select
    End Sub
End Class
