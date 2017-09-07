Imports System.IO
Imports System.Net

Public Class Form2
    Dim host As String = System.Net.Dns.GetHostName()
    Dim LocalHostaddress As String = System.Net.Dns.GetHostByName(host).AddressList(0).ToString()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        post_data()
        Dim filepath As String
        filepath = System.IO.Path.GetTempPath & "Cedula_caraA.jpg"
        If My.Computer.FileSystem.FileExists(filepath) Then
            My.Computer.FileSystem.DeleteFile(filepath)
        End If
        My.Settings.scan = TwainLib.GetScanSource
        My.Settings.Save()
        Console.WriteLine(My.Settings.scan)
        TwainLib.ScanIt(My.Settings.scan, "Cedula_caraA")
        Dim fs As New IO.FileStream(System.IO.Path.GetTempPath & "Cedula_caraA.jpg", IO.FileMode.Open, IO.FileAccess.Read)
        PictureBox1.Image = Image.FromStream(fs)
        fs.Close()

        post_image("Cedula_caraA.png", "Cedula_caraA.jpg")

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        post_data()
        Dim filepath As String
        filepath = System.IO.Path.GetTempPath & "Cedula_caraB.jpg"
        If My.Computer.FileSystem.FileExists(filepath) Then
            My.Computer.FileSystem.DeleteFile(filepath)
        End If
        My.Settings.scan = TwainLib.GetScanSource
        My.Settings.Save()
        Console.WriteLine(My.Settings.scan)
        TwainLib.ScanIt(My.Settings.scan, "Cedula_caraB")
        Dim fs As New IO.FileStream(System.IO.Path.GetTempPath & "Cedula_caraB.jpg", IO.FileMode.Open, IO.FileAccess.Read)
        PictureBox2.Image = Image.FromStream(fs)
        fs.Close()

        post_image("Cedula_caraB.png", "Cedula_caraB.jpg")

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        post_data()
        Dim filepath As String
        filepath = System.IO.Path.GetTempPath & "tarjeta_credito.jpg"
        If My.Computer.FileSystem.FileExists(filepath) Then
            My.Computer.FileSystem.DeleteFile(filepath)
        End If
        My.Settings.scan = TwainLib.GetScanSource
        My.Settings.Save()
        Console.WriteLine(My.Settings.scan)
        TwainLib.ScanIt(My.Settings.scan, "tarjeta_credito")
        Dim fs As New IO.FileStream(System.IO.Path.GetTempPath & "tarjeta_credito.jpg", IO.FileMode.Open, IO.FileAccess.Read)
        PictureBox3.Image = Image.FromStream(fs)
        fs.Close()

        post_image("tarjeta_credito.png", "tarjeta_credito.jpg")
    End Sub


    Private Function post_data()
        'JESUS VEGA This is an example of how request must be doned on vb, this class is not on production'
        ' Create a request using a URL that can receive a post. 
        Dim request As WebRequest = WebRequest.Create("http://app.aoacolombia.com/Control/operativo/controllers/RecepcionController.php")
        ' Set the Method property of the request to POST.
        request.Method = "POST"
        request.ContentType = "multipart/form-data"
        ' Create POST data and convert it to a byte array.
        Dim email = "correo@soemthing.com"
        Dim password = "blablabla"
        Dim postData As String = String.Format("inputEmailHandle={0}&name={1}&inputPassword={2}", email, "jesus vega", password)
        Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes(postData)
        ' Set the ContentType property of the WebRequest.
        request.ContentType = "application/x-www-form-urlencoded"
        ' Set the ContentLength property of the WebRequest.
        request.ContentLength = byteArray.Length
        ' Get the request stream.
        Dim dataStream As IO.Stream = request.GetRequestStream()
        ' Write the data to the request stream.
        dataStream.Write(byteArray, 0, byteArray.Length)

        ' Close the Stream object.
        dataStream.Close()
        ' Get the response.
        Dim response As WebResponse = request.GetResponse()
        ' Display the status.
        Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        MessageBox.Show(CType(response, HttpWebResponse).StatusDescription, "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)
        ' Get the stream containing content returned by the server.
        dataStream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim reader As New IO.StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = reader.ReadToEnd()
        ' Display the content.
        'Console.WriteLine(responseFromServer)'
        'MessageBox.Show(responseFromServer, "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)'
        ' Clean up the streams.
        reader.Close()
        dataStream.Close()
        response.Close()
        Return response
    End Function

    Private Function post_image(fname As String, ffile As String)
        Dim boundary As String = "-----------------------------" & DateTime.Now.Ticks.ToString("x")
        Dim FileName = fname
        Dim request As HttpWebRequest = CType(WebRequest.Create("http://app.aoacolombia.com/Control/operativo/controllers/RecepcionController.php"), HttpWebRequest)
        request.Method = "POST"
        request.ContentType = "multipart/form-data; boundary=" & "---------------------------" & DateTime.Now.Ticks.ToString("x")

        request.KeepAlive = True
        Dim builder As New System.Text.StringBuilder()
        builder.Append(boundary & vbCrLf & "Content-Disposition: form-data; name=""variable1""" & vbCrLf & vbCrLf & "1" & vbCrLf)
        builder.Append(boundary & vbCrLf & "Content-Disposition: form-data; name=""file""; filename=""" & FileName & """" & vbCrLf)
        builder.Append("Content-Type: application/octet-stream")
        builder.Append(vbCrLf & vbCrLf)


        Dim _Buffer() As Byte = Nothing
        Dim ImageData As System.Drawing.Image
        Dim fs As New IO.FileStream(System.IO.Path.GetTempPath & ffile, IO.FileMode.Open, IO.FileAccess.Read)
        Dim _BinaryReader As New System.IO.BinaryReader(fs)
        Dim _TotalBytes As Long = New System.IO.FileInfo(System.IO.Path.GetTempPath & ffile).Length
        _Buffer = _BinaryReader.ReadBytes(CInt(Fix(_TotalBytes)))


        ImageData = Image.FromStream(fs)
        Dim encoded As String
        encoded = Convert.ToBase64String(_Buffer)
        fs.Close()
        builder.Append(encoded)
        builder.Append(vbCrLf)
        builder.Append(boundary & vbCrLf & "Content-Disposition: form-data; name=""privateip""" & vbCrLf & vbCrLf & LocalHostaddress)
        ' Footer Bytes
        Dim close As Byte() = System.Text.Encoding.UTF8.GetBytes("--")
        Dim postHeader As String = builder.ToString()
        Dim postHeaderBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(postHeader)
        Dim boundaryBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(vbCrLf & boundary & "--" & vbCrLf)
        Dim length As Long = postHeaderBytes.Length + boundaryBytes.Length
        request.ContentLength = length
        Dim requestStream As IO.Stream = request.GetRequestStream()
        Dim fulllength As Integer = postHeaderBytes.Length + boundaryBytes.Length
        ' Write out our post header
        requestStream.Write(postHeaderBytes, 0, postHeaderBytes.Length)
        ' Write out the trailing boundary
        requestStream.Write(boundaryBytes, 0, boundaryBytes.Length)
        'Dim dataStream As IO.Stream = request.GetRequestStream()'
        'dataStream.Write(bytes, 0, fs.Length)'

        Dim response = request.GetResponse()
        requestStream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim reader As New IO.StreamReader(requestStream)
        ' Read the content.
        Dim responseFromServer As String = reader.ReadToEnd()
        ' Display the content.
        Console.WriteLine(responseFromServer)
        MessageBox.Show(responseFromServer, "My Application", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk)

        Return response
    End Function




End Class