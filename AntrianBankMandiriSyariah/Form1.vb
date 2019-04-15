Imports System
Imports System.IO.Ports
Public Class Form1
    Dim bulan() As String = {"", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"}
    Dim hari() As String = {"", "Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu", "Minggu"}

    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim NoAntri As String = ""

    Dim comPORT As String
    Dim receivedData As String = ""
    Dim labelUP As String
    Dim LabelDown As String


    Dim fontSize As Integer = 0
    Dim tinggi = 10

    Dim TextToPrint As String = ""
    Dim jumlahkarakter As Integer = 40 '40 adalah jumlah karakter (lebar) yang ada pada struk

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        a = My.Settings.a
        b = My.Settings.b
        c = My.Settings.c
        view()
        refPrinterinstall()
        refCom()
        ComboBox2.Text = My.Settings.printerName

    End Sub
    

    


    Public Sub view()
        a = My.Settings.a
        b = My.Settings.b
        c = My.Settings.c
        TextBox1.Text = My.Settings.a
        TextBox2.Text = My.Settings.b
        TextBox5.Text = My.Settings.c

    End Sub

    Private Sub aUp()
        a += 1
        My.Settings.a = a
        view()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        aDown()
    End Sub
    Private Sub aDown()
        a -= 1
        My.Settings.a = a
        view()
    End Sub

    Private Sub bDown()
        b -= 1
        My.Settings.b = b
        view()
    End Sub

    Private Sub bUp()
        b += 1
        My.Settings.b = b
        view()
    End Sub

    Private Sub cDown()
        c -= 1
        My.Settings.c = c
        view()
    End Sub

    Private Sub cUp()
        c += 1
        My.Settings.c = c
        view()
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        aUp()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        NoAntri = TextBox1.Text
        labelUP = "A1"
        LabelDown = "TELLER"
        cetak()
        aUp()
    End Sub
    Public Sub cetak()

        Dim printControl = New Printing.StandardPrintController
        PrintDocument1.PrintController = printControl
        Try
            PrintDocument1.Print()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        bDown()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        bUp()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        cDown()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        cUp()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        For i = 0 To 0

            TextToPrint = "Tgl: " & DateTime.Now.Day & " " & bulan(DateTime.Now.Month) & " " & DateTime.Now.Year & "    " & Format(DateTime.Now, "HH:mm:ss")
            e.Graphics.DrawString(TextToPrint, New Drawing.Font("Arial", 9), Brushes.Black, 5, 2)
            TextToPrint = labelUP
            e.Graphics.DrawString(TextToPrint, New Drawing.Font("Arial", 16), Brushes.Black, 85, 20)
            Dim posHor As Integer = (185 - (NoAntri.Length * 23)) / 2
            e.Graphics.DrawString(NoAntri, New Drawing.Font("Arial", 30), Brushes.Black, posHor, 55)
            TextToPrint = LabelDown
            posHor = (185 - (TextToPrint.Length * 15.4)) / 2
            e.Graphics.DrawString(TextToPrint, New Drawing.Font("Arial", 20), Brushes.Black, posHor, 120)

            '==========================
            ' NoAntri = "888"
            'TextToPrint = "8"
            ' Dim posHor As Integer = (185 - (TextToPrint.Length * 15)) / 2
            'posHor = 0

            'e.Graphics.DrawString(TextToPrint, New Drawing.Font("Arial", 20), Brushes.Black, posHor, 0)

        Next


    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        NoAntri = TextBox2.Text
        labelUP = "A2"
        LabelDown = "GADAI"
        cetak()
        bUp()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            receivedData = ReceiveSerialData()
            If (Val(receivedData) = 1) Then
                NoAntri = TextBox1.Text
                labelUP = "A1"
                LabelDown = "TELLER"
                cetak()
                aUp()
            ElseIf (Val(receivedData) = 2) Then
                NoAntri = TextBox2.Text
                labelUP = "A2"
                LabelDown = "GADAI"
                cetak()
                bUp()
            ElseIf (Val(receivedData) = 3) Then
                NoAntri = TextBox5.Text
                labelUP = "A3"
                LabelDown = "CS"
                cetak()
                cUp()

            End If
            Label3.Text = "Terkoneksi"
        Catch ex As Exception
            Label3.Text = "Tidak terkoneksi"

        End Try
        
        

    End Sub
    Function ReceiveSerialData() As String
        Dim Incoming As String
        Try
            Incoming = SerialPort1.ReadExisting()
            If Incoming Is Nothing Then
                Return "nothing" & vbCrLf
            Else
                Return Incoming
            End If
        Catch ex As TimeoutException
            Return "Error: Serial Port read timed out."
        End Try

    End Function

    Private Sub connect_BTN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        connectPort()
    End Sub
    Private Sub connectPort()
        If (comPORT <> "") Then
            SerialPort1.Close()
            SerialPort1.PortName = comPORT
            SerialPort1.BaudRate = 9600
            SerialPort1.DataBits = 8
            SerialPort1.Parity = Parity.None
            SerialPort1.StopBits = StopBits.One
            SerialPort1.Handshake = Handshake.None
            SerialPort1.Encoding = System.Text.Encoding.Default
            SerialPort1.ReadTimeout = 10000

            SerialPort1.Open()

            Timer1.Enabled = True
        Else
            MsgBox("Select a COM port first")
        End If

        SerialPort1.Close()
    End Sub
  

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        NoAntri = TextBox5.Text
        labelUP = "A3"
        LabelDown = "CS"
        cetak()
        cUp()
    End Sub


    Public Sub refPrinterinstall()
        For Each prnt As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
            ComboBox2.Items.Add(prnt)
        Next

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        My.Settings.printerName = ComboBox2.Text
        PrintDocument1.PrinterSettings.PrinterName = My.Settings.printerName
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        comPORT = ComboBox1.Text
        If (comPORT <> "") Then
            SerialPort1.Close()
            SerialPort1.PortName = comPORT
            SerialPort1.BaudRate = 9600
            SerialPort1.DataBits = 8
            SerialPort1.Parity = Parity.None
            SerialPort1.StopBits = StopBits.One
            SerialPort1.Handshake = Handshake.None
            SerialPort1.Encoding = System.Text.Encoding.Default
            SerialPort1.ReadTimeout = 10000
            Try
                SerialPort1.Open()
            Catch ex As Exception
                MsgBox("Colokkan Perangkat Control!")
            End Try

            Timer1.Enabled = True
        End If
    End Sub
    Public Sub refCom()

        ComboBox1.Items.Clear()
        ComboBox1.Text = ""
        For Each sp As String In My.Computer.Ports.SerialPortNames
            ComboBox1.Items.Add(sp)
        Next
        If ComboBox1.Items.Count > 0 Then
            For Each port As String In ComboBox1.Items
                ComboBox1.Text = port
            Next
        Else
            MsgBox("Colokkan Perangkat Tombol!")
        End If
        
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        refCom()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        My.Settings.a = 1
        My.Settings.b = 1
        My.Settings.c = 1
        view()
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        Label4.Text = hari(DateTime.Now.DayOfWeek) & " " & DateTime.Now.Day & " " & bulan(DateTime.Now.Month) & " " & DateTime.Now.Year & "     " & Format(DateTime.Now, "HH:mm:ss")

        If (Format(My.Settings.tgl, "yyyy/MM/dd") <> Format(Date.Now, "yyyy/MM/dd")) Then
            My.Settings.a = 1
            My.Settings.b = 1
            My.Settings.c = 1
            My.Settings.tgl = Date.Now
            view()

        End If
    End Sub
End Class
