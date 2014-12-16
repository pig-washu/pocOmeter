Imports System.Data


Class MainWindow


    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset


    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)

        cnn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test_nopid.xls;Extended Properties=Excel 8.0")

        rst.Open("SELECT * FROM [Sheet1$];", cnn, 1, 1)

        rst.MoveFirst()
        Do While Not rst.EOF

            Debug.Print(rst("MRN").Value)
            rst.MoveNext()

        Loop

        cnn.Close()

    End Sub


End Class
