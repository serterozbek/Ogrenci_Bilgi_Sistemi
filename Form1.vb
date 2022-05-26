Imports System.Data.OleDb
Imports System.Xml
Imports Microsoft.Office.Interop


Public Class Form1
    'Excel
    Dim excel_uyg As New Excel.Application
    Dim exc_workbook As Excel.Workbook
    Dim exc_sheet As Excel.Worksheet
    'Excel

    'Word
    Dim word As New Word.Application
    Dim yenibelge As New Word.Document
    'Word
    Dim con As OleDbConnection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("ALGORİTMA")
        ComboBox1.Items.Add("NET'E GİRİŞ")
        ComboBox1.Items.Add("VERİTABANI")

        DataGridView1.Columns.Add("dersKodu", "Ders Kodu")
        DataGridView1.Columns.Add("dersAdi", "Ders Adı")
        DataGridView1.Columns.Add("okulNo", "Okul No")
        DataGridView1.Columns.Add("adi", "Adı")
        DataGridView1.Columns.Add("soyadi", "Soyadı")
        DataGridView1.Columns.Add("vize", "Vize")
        DataGridView1.Columns.Add("final", "Final")
        DataGridView1.Columns.Add("ortalama", "Ortalama")
        DataGridView1.Columns.Add("bagil", "Bağıl")
        DataGridView1.Columns.Add("harfNotu", "Harf Notu")

        DataGridView1.Columns("dersKodu").ReadOnly = True
        DataGridView1.Columns("dersAdi").ReadOnly = True
        DataGridView1.Columns("okulNo").ReadOnly = True
        DataGridView1.Columns("adi").ReadOnly = True
        DataGridView1.Columns("soyadi").ReadOnly = True
        DataGridView1.Columns("vize").ReadOnly = True
        DataGridView1.Columns("final").ReadOnly = True
        DataGridView1.Columns("ortalama").ReadOnly = True
        DataGridView1.Columns("bagil").ReadOnly = True

        DataGridView1.EditMode = False
        Button6.Enabled = False

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Button3.Enabled = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedItem
            Case "ALGORİTMA"
                algoritmaToDataGrid()
            Case "NET'E GİRİŞ"
                netegirisToDataGrid()
            Case "VERİTABANI"
                veritabaniToDataGrid()
            Case Else
                MsgBox("Böyle bir ders kayıtlı değil.")
        End Select
        Button6.Enabled = True

        DataGridView1.Columns("vize").ReadOnly = False
        DataGridView1.Columns("final").ReadOnly = False

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        ' ReadOnly Combobox
        e.SuppressKeyPress = True
    End Sub
    Public Sub algoritmaToDataGrid()
        excel_uyg.Workbooks.Open(Application.StartupPath + "\\algoritma.xlsx")
        exc_workbook = excel_uyg.ActiveWorkbook
        exc_sheet = exc_workbook.ActiveSheet

        Dim sayac, i As Integer
        sayac = 0
        i = 1
        Do While i < 1000
            sayac = sayac + 1
            If String.IsNullOrEmpty(exc_sheet.Cells(i, 1).Value) Then
                Exit Do
            End If
            i = i + 1
        Loop
        DataGridView1.Rows.Clear()
        For i = 1 To sayac - 2
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i - 1).Cells(0).Value = exc_sheet.Cells(i + 1, 1).Value
            DataGridView1.Rows(i - 1).Cells(1).Value = exc_sheet.Cells(i + 1, 2).Value
            DataGridView1.Rows(i - 1).Cells(2).Value = exc_sheet.Cells(i + 1, 3).Value
            DataGridView1.Rows(i - 1).Cells(3).Value = exc_sheet.Cells(i + 1, 4).Value
            DataGridView1.Rows(i - 1).Cells(4).Value = exc_sheet.Cells(i + 1, 5).Value
        Next
    End Sub
    Public Sub netegirisToDataGrid()
        excel_uyg.Workbooks.Open(Application.StartupPath + "\\netegiris.xlsx")
        exc_workbook = excel_uyg.ActiveWorkbook
        exc_sheet = exc_workbook.ActiveSheet

        Dim sayac, i As Integer
        sayac = 0
        i = 1
        Do While i < 1000
            sayac = sayac + 1
            If String.IsNullOrEmpty(exc_sheet.Cells(i, 1).Value) Then
                Exit Do
            End If
            i = i + 1
        Loop
        DataGridView1.Rows.Clear()
        For i = 1 To sayac - 2
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i - 1).Cells(0).Value = exc_sheet.Cells(i + 1, 1).Value
            DataGridView1.Rows(i - 1).Cells(1).Value = exc_sheet.Cells(i + 1, 2).Value
            DataGridView1.Rows(i - 1).Cells(2).Value = exc_sheet.Cells(i + 1, 3).Value
            DataGridView1.Rows(i - 1).Cells(3).Value = exc_sheet.Cells(i + 1, 4).Value
            DataGridView1.Rows(i - 1).Cells(4).Value = exc_sheet.Cells(i + 1, 5).Value
            DataGridView1.Rows(i - 1).Cells(5).Value = exc_sheet.Cells(i + 1, 6).Value
        Next
    End Sub

    Public Sub veritabaniToDataGrid()
        excel_uyg.Workbooks.Open(Application.StartupPath + "\\veritabani.xlsx")
        exc_workbook = excel_uyg.ActiveWorkbook
        exc_sheet = exc_workbook.ActiveSheet

        Dim sayac, i As Integer
        sayac = 0
        i = 1
        Do While i < 1000
            sayac = sayac + 1
            If String.IsNullOrEmpty(exc_sheet.Cells(i, 1).Value) Then
                Exit Do
            End If
            i = i + 1
        Loop

        DataGridView1.Rows.Clear()
        For i = 1 To sayac - 2
            DataGridView1.Rows.Add()
            DataGridView1.Rows(i - 1).Cells(0).Value = exc_sheet.Cells(i + 1, 1).Value
            DataGridView1.Rows(i - 1).Cells(1).Value = exc_sheet.Cells(i + 1, 2).Value
            DataGridView1.Rows(i - 1).Cells(2).Value = exc_sheet.Cells(i + 1, 3).Value
            DataGridView1.Rows(i - 1).Cells(3).Value = exc_sheet.Cells(i + 1, 4).Value
            DataGridView1.Rows(i - 1).Cells(4).Value = exc_sheet.Cells(i + 1, 5).Value
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        veritabani()
    End Sub

    Public Sub veritabani()
        Dim connectionStr As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=ogr.accdb"
        con = New OleDbConnection(connectionStr)

        Try
            For i = 0 To DataGridView1.Rows.Count - 1
                Dim derskodu As String = DataGridView1.Rows(i).Cells(0).Value
                Dim dersadi As String = DataGridView1.Rows(i).Cells(1).Value
                Dim okulno As String = DataGridView1.Rows(i).Cells(2).Value
                Dim adi As String = DataGridView1.Rows(i).Cells(3).Value
                Dim soyadi As String = DataGridView1.Rows(i).Cells(4).Value
                Dim vize As Double = DataGridView1.Rows(i).Cells(5).Value
                Dim final As Double = DataGridView1.Rows(i).Cells(6).Value
                Dim ort As Double = DataGridView1.Rows(i).Cells(7).Value
                Dim bagil As Double = DataGridView1.Rows(i).Cells(8).Value


                Select Case ComboBox1.SelectedItem
                    Case "ALGORİTMA"
                        con.Open()
                        Dim ins As String = "insert into algoritma(okulno,derskodu,dersadi,adi,soyadi,vize,final,ortalama,bagil) values(@okulno, @derskodu,@dersadi,@adi,@soyadi,@vize,@final,@ortalama,@bagil)"
                        Dim sorgu As OleDbCommand = New OleDbCommand(ins, con)
                        sorgu.Parameters.AddWithValue("@okulno", okulno)
                        sorgu.Parameters.AddWithValue("@derskodu", derskodu)
                        sorgu.Parameters.AddWithValue("@dersadi", dersadi)
                        sorgu.Parameters.AddWithValue("@adi", adi)
                        sorgu.Parameters.AddWithValue("@soyadi", soyadi)
                        sorgu.Parameters.AddWithValue("@vize", vize)
                        sorgu.Parameters.AddWithValue("@final", final)
                        sorgu.Parameters.AddWithValue("@ortalama", ort)
                        sorgu.Parameters.AddWithValue("@bagil", bagil)
                        sorgu.ExecuteNonQuery()
                        con.Close()

                    Case "NET'E GİRİŞ"
                        con.Open()
                        Dim ins As String = "insert into netegiris(okulno,derskodu,dersadi,adi,soyadi,vize,final,ortalama,bagil) values(@okulno, @derskodu,@dersadi,@adi,@soyadi,@vize,@final,@ortalama,@bagil)"
                        Dim sorgu As OleDbCommand = New OleDbCommand(ins, con)
                        sorgu.Parameters.AddWithValue("@okulno", okulno)
                        sorgu.Parameters.AddWithValue("@derskodu", derskodu)
                        sorgu.Parameters.AddWithValue("@dersadi", dersadi)
                        sorgu.Parameters.AddWithValue("@adi", adi)
                        sorgu.Parameters.AddWithValue("@soyadi", soyadi)
                        sorgu.Parameters.AddWithValue("@vize", vize)
                        sorgu.Parameters.AddWithValue("@final", final)
                        sorgu.Parameters.AddWithValue("@ortalama", ort)
                        sorgu.Parameters.AddWithValue("@bagil", bagil)
                        sorgu.ExecuteNonQuery()
                        con.Close()

                    Case "VERİTABANI"
                        con.Open()
                        Dim ins As String = "insert into veritabani(okulno,derskodu,dersadi,adi,soyadi,vize,final,ortalama,bagil) values(@okulno, @derskodu,@dersadi,@adi,@soyadi,@vize,@final,@ortalama,@bagil)"
                        Dim sorgu As OleDbCommand = New OleDbCommand(ins, con)
                        sorgu.Parameters.AddWithValue("@okulno", okulno)
                        sorgu.Parameters.AddWithValue("@derskodu", derskodu)
                        sorgu.Parameters.AddWithValue("@dersadi", dersadi)
                        sorgu.Parameters.AddWithValue("@adi", adi)
                        sorgu.Parameters.AddWithValue("@soyadi", soyadi)
                        sorgu.Parameters.AddWithValue("@vize", vize)
                        sorgu.Parameters.AddWithValue("@final", final)
                        sorgu.Parameters.AddWithValue("@ortalama", ort)
                        sorgu.Parameters.AddWithValue("@bagil", bagil)
                        sorgu.ExecuteNonQuery()
                        con.Close()
                    Case Else
                        MsgBox("Lütfen ders seçimi yapınız")
                End Select
            Next
            MsgBox("Veritabanına aktarma başarılı")
        Catch ex As Exception
            MsgBox("Bir hata meydana geldi. ")
        End Try

        DataGridView1.Columns("vize").ReadOnly = True
        DataGridView1.Columns("final").ReadOnly = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Word
        yenibelge = word.Documents.Add
        word.Visible = True
        Dim i As Integer
        For i = 2 To DataGridView1.Rows.Count
            yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(0).Value.ToString + " ")
            yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(1).Value.ToString + " ")
            yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(2).Value.ToString + " ")
            yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(3).Value.ToString + " ")
            yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(4).Value.ToString + " ")
            'vize
            If DataGridView1.Rows(i - 2).Cells(5).Value IsNot Nothing Then
                yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(5).Value.ToString + " ")
            Else
                yenibelge.Range.InsertAfter("  " + " ")
            End If

            'final
            If DataGridView1.Rows(i - 2).Cells(6).Value IsNot Nothing Then
                yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(6).Value.ToString + " ")
            Else
                yenibelge.Range.InsertAfter("  " + " ")
            End If

            'ortalama
            If DataGridView1.Rows(i - 2).Cells(7).Value IsNot Nothing Then
                yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(7).Value.ToString + " ")
            Else
                yenibelge.Range.InsertAfter("  " + " ")
            End If

            'bagil
            If DataGridView1.Rows(i - 2).Cells(8).Value IsNot Nothing Then
                yenibelge.Range.InsertAfter(DataGridView1.Rows(i - 2).Cells(8).Value.ToString + " ")
            Else
                yenibelge.Range.InsertAfter("  " + " ")
            End If

            yenibelge.Paragraphs.Add()
        Next

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Mail Gönder
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'XML
        Dim dosya As New XmlDocument
        Try
            Select Case ComboBox1.SelectedItem
                Case "ALGORİTMA"
                    dosya.Load("algoritma.xml")
                    For i = 0 To DataGridView1.Rows.Count - 2
                        Dim ogrenci As XmlElement = dosya.CreateElement("ogrenci")
                        ogrenci.SetAttribute("okulno", DataGridView1.Rows(i).Cells(2).Value.ToString)

                        Dim derskodu As XmlNode = dosya.CreateNode("element", "derskodu", DataGridView1.Rows(i).Cells(0).Value.ToString)
                        ogrenci.AppendChild(derskodu)

                        Dim dersadi As XmlNode = dosya.CreateNode("element", "dersadi", DataGridView1.Rows(i).Cells(1).Value.ToString)
                        ogrenci.AppendChild(dersadi)

                        Dim adi As XmlNode = dosya.CreateNode("element", "adi", DataGridView1.Rows(i).Cells(3).Value.ToString)
                        ogrenci.AppendChild(adi)

                        Dim soyadi As XmlNode = dosya.CreateNode("element", "soyadi", DataGridView1.Rows(i).Cells(4).Value.ToString)
                        ogrenci.AppendChild(soyadi)

                        If DataGridView1.Rows(i).Cells(5).Value IsNot Nothing Then
                            Dim vize As XmlNode = dosya.CreateNode("element", "vize", DataGridView1.Rows(i).Cells(5).Value.ToString)
                            ogrenci.AppendChild(vize)
                        Else
                            Dim vize As XmlNode = dosya.CreateNode("element", "vize", "")
                            ogrenci.AppendChild(vize)
                        End If

                        If DataGridView1.Rows(i).Cells(6).Value IsNot Nothing Then
                            Dim final As XmlNode = dosya.CreateNode("element", "final", DataGridView1.Rows(i).Cells(6).Value.ToString)
                            ogrenci.AppendChild(final)
                        Else
                            Dim final As XmlNode = dosya.CreateNode("element", "final", "")
                            ogrenci.AppendChild(final)
                        End If

                        If DataGridView1.Rows(i).Cells(7).Value IsNot Nothing Then
                            Dim ort As XmlNode = dosya.CreateNode("element", "ort", DataGridView1.Rows(i).Cells(7).Value.ToString)
                            ogrenci.AppendChild(ort)
                        Else
                            Dim ort As XmlNode = dosya.CreateNode("element", "ort", "")
                            ogrenci.AppendChild(ort)
                        End If

                        If DataGridView1.Rows(i).Cells(8).Value IsNot Nothing Then
                            Dim bagil As XmlNode = dosya.CreateNode("element", "bagil", DataGridView1.Rows(i).Cells(8).Value.ToString)
                            ogrenci.AppendChild(bagil)
                        Else
                            Dim bagil As XmlNode = dosya.CreateNode("element", "bagil", "")
                            ogrenci.AppendChild(bagil)
                        End If

                        dosya.DocumentElement.AppendChild(ogrenci)
                    Next
                    dosya.Save("algoritma.xml")
                    MsgBox("Algoritma xml kaydı başarılı")

                Case "NET'E GİRİŞ"
                    dosya.Load("netegiris.xml")
                    For i = 0 To DataGridView1.Rows.Count - 2
                        Dim ogrenci As XmlElement = dosya.CreateElement("ogrenci")
                        ogrenci.SetAttribute("okulno", DataGridView1.Rows(i).Cells(2).Value.ToString)

                        Dim derskodu As XmlNode = dosya.CreateNode("element", "derskodu", DataGridView1.Rows(i).Cells(0).Value.ToString)
                        ogrenci.AppendChild(derskodu)

                        Dim dersadi As XmlNode = dosya.CreateNode("element", "dersadi", DataGridView1.Rows(i).Cells(1).Value.ToString)
                        ogrenci.AppendChild(dersadi)

                        Dim adi As XmlNode = dosya.CreateNode("element", "adi", DataGridView1.Rows(i).Cells(3).Value.ToString)
                        ogrenci.AppendChild(adi)

                        Dim soyadi As XmlNode = dosya.CreateNode("element", "soyadi", DataGridView1.Rows(i).Cells(4).Value.ToString)
                        ogrenci.AppendChild(soyadi)

                        If DataGridView1.Rows(i).Cells(5).Value IsNot Nothing Then
                            Dim vize As XmlNode = dosya.CreateNode("element", "vize", DataGridView1.Rows(i).Cells(5).Value.ToString)
                            ogrenci.AppendChild(vize)

                        Else
                            Dim vize As XmlNode = dosya.CreateNode("element", "vize", "")
                            ogrenci.AppendChild(vize)
                        End If

                        If DataGridView1.Rows(i).Cells(6).Value IsNot Nothing Then
                            Dim final As XmlNode = dosya.CreateNode("element", "final", DataGridView1.Rows(i).Cells(6).Value.ToString)
                            ogrenci.AppendChild(final)
                        Else
                            Dim final As XmlNode = dosya.CreateNode("element", "final", "")
                            ogrenci.AppendChild(final)
                        End If

                        If DataGridView1.Rows(i).Cells(7).Value IsNot Nothing Then
                            Dim ort As XmlNode = dosya.CreateNode("element", "ort", DataGridView1.Rows(i).Cells(7).Value.ToString)
                            ogrenci.AppendChild(ort)
                        Else
                            Dim ort As XmlNode = dosya.CreateNode("element", "ort", "")
                            ogrenci.AppendChild(ort)
                        End If

                        If DataGridView1.Rows(i).Cells(8).Value IsNot Nothing Then
                            Dim bagil As XmlNode = dosya.CreateNode("element", "bagil", DataGridView1.Rows(i).Cells(8).Value.ToString)
                            ogrenci.AppendChild(bagil)
                        Else
                            Dim bagil As XmlNode = dosya.CreateNode("element", "bagil", "")
                            ogrenci.AppendChild(bagil)
                        End If

                        dosya.DocumentElement.AppendChild(ogrenci)
                    Next
                    dosya.Save("netegiris.xml")
                    MsgBox("Net'e Giriş xml kaydı başarılı")

                Case "VERİTABANI"
                    dosya.Load("veritabani.xml")
                    For i = 0 To DataGridView1.Rows.Count - 2
                        Dim ogrenci As XmlElement = dosya.CreateElement("ogrenci")
                        ogrenci.SetAttribute("okulno", DataGridView1.Rows(i).Cells(2).Value.ToString)

                        Dim derskodu As XmlNode = dosya.CreateNode("element", "derskodu", DataGridView1.Rows(i).Cells(0).Value.ToString)
                        ogrenci.AppendChild(derskodu)

                        Dim dersadi As XmlNode = dosya.CreateNode("element", "dersadi", DataGridView1.Rows(i).Cells(1).Value.ToString)
                        ogrenci.AppendChild(dersadi)

                        Dim adi As XmlNode = dosya.CreateNode("element", "adi", DataGridView1.Rows(i).Cells(3).Value.ToString)
                        ogrenci.AppendChild(adi)

                        Dim soyadi As XmlNode = dosya.CreateNode("element", "soyadi", DataGridView1.Rows(i).Cells(4).Value.ToString)
                        ogrenci.AppendChild(soyadi)

                        If DataGridView1.Rows(i).Cells(5).Value IsNot Nothing Then
                            Dim vize As XmlNode = dosya.CreateNode("element", "vize", DataGridView1.Rows(i).Cells(5).Value.ToString)
                            ogrenci.AppendChild(vize)

                        Else
                            Dim vize As XmlNode = dosya.CreateNode("element", "vize", "")
                            ogrenci.AppendChild(vize)
                        End If

                        If DataGridView1.Rows(i).Cells(6).Value IsNot Nothing Then
                            Dim final As XmlNode = dosya.CreateNode("element", "final", DataGridView1.Rows(i).Cells(6).Value.ToString)
                            ogrenci.AppendChild(final)
                        Else
                            Dim final As XmlNode = dosya.CreateNode("element", "final", "")
                            ogrenci.AppendChild(final)
                        End If

                        If DataGridView1.Rows(i).Cells(7).Value IsNot Nothing Then
                            Dim ort As XmlNode = dosya.CreateNode("element", "ort", DataGridView1.Rows(i).Cells(7).Value.ToString)
                            ogrenci.AppendChild(ort)
                        Else
                            Dim ort As XmlNode = dosya.CreateNode("element", "ort", "")
                            ogrenci.AppendChild(ort)
                        End If

                        If DataGridView1.Rows(i).Cells(8).Value IsNot Nothing Then
                            Dim bagil As XmlNode = dosya.CreateNode("element", "bagil", DataGridView1.Rows(i).Cells(8).Value.ToString)
                            ogrenci.AppendChild(bagil)
                        Else
                            Dim bagil As XmlNode = dosya.CreateNode("element", "bagil", "")
                            ogrenci.AppendChild(bagil)
                        End If

                        dosya.DocumentElement.AppendChild(ogrenci)
                    Next
                    dosya.Save("veritabani.xml")
                    MsgBox("Veritabanı xml kaydı başarılı")
                Case Else
                    MsgBox("Lütfen ders seçimi yapınız")
            End Select
        Catch ex As Exception
            MsgBox("Bir hata meydana geldi. ")
        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Not girişi
        Dim x, vize, final
        x = TextBox1.Text
        vize = TextBox2.Text
        final = TextBox3.Text

        If TextBox2.Text = "" Or TextBox3.Text = "" Then
            Beep()
            MsgBox("Boş veri eklenemez. Öğrenci sınava girmedi ise -1 değerini giriniz")
            Exit Sub
        End If

        If Not sayiMi(TextBox2.Text) Then
            Beep()
            MsgBox("Sayısal değer giriniz")
            Exit Sub
        End If

        If Not sayiMi(TextBox3.Text) Then
            Beep()
            MsgBox("Sayısal değer giriniz")
            Exit Sub
        End If

        If (vize > 100 Or vize < -1) Or (final > 100 Or final < -1) Then
            Beep()
            MessageBox.Show("sınav notu 0 ile 100 arasında olmalıdır. Öğrenci sınava girmedi ise -1 değerini giriniz.")
            Exit Sub
        End If

        Dim i As Integer = DataGridView1.CurrentRow.Index()
        DataGridView1.Rows(i).Cells(5).Value = vize
        DataGridView1.Rows(i).Cells(6).Value = final



    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'BAĞIL HESAPLAMA
        ' Verisetindeki tüm ortalama değerler alınır ve sınıf mevcuduna bölünerek sınıf ortalaması bulunur.

        Try
            Dim x As Double
            For i = 0 To DataGridView1.Rows.Count - 1
                x = x + Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value)
            Next

            Dim sinifort As Double
            sinifort = x / DataGridView1.Rows.Count - 1

            ' Varyans bulunur.
            ' Verisetindeki bütün sayılardan sinifortalaması çıkarılır.(a)
            ' Sonra sonucun karesi alınır.(b)
            ' Daha sonra karesi alınan sayılar toplanır. (c)
            Dim a, b, c As Double
            a = 0
            b = 0
            c = 0

            For i = 0 To DataGridView1.Rows.Count - 1
                a = Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value) - sinifort
                b = a * a
                c = c + b
            Next

            ' Karelerin toplamı n-1'e bölünür.
            Dim varyans As Double = c / DataGridView1.Rows.Count - 2
            Dim standartSapma As Double = Val(varyans ^ 1 / 2)

            For i = 0 To DataGridView1.Rows.Count - 1
                Dim ogrort As Double = Convert.ToDouble(DataGridView1.Rows(i).Cells(7).Value)
                Dim bagilnot As Double = (((ogrort - sinifort) / standartSapma) * 10) + 50
                DataGridView1.Rows(i).Cells(8).Value = bagilnot
            Next

            If 80 < sinifort And sinifort <= 100 Then
                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 57 Then
                        harfNotu = "AA"
                    ElseIf 52 <= bagilnot And bagilnot <= 56.99 Then
                        harfNotu = "BA"
                    ElseIf 47 <= bagilnot And bagilnot <= 51.99 Then
                        harfNotu = "BB"
                    ElseIf 42 <= bagilnot And bagilnot <= 46.99 Then
                        harfNotu = "CB"
                    ElseIf 37 <= bagilnot And bagilnot <= 41.99 Then
                        harfNotu = "CC"
                    ElseIf 32 <= bagilnot And bagilnot <= 36.99 Then
                        harfNotu = "DC"
                    ElseIf 27 <= bagilnot And bagilnot <= 31.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 27 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If



            If 70 < sinifort And sinifort <= 80 Then

                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 59 Then
                        harfNotu = "AA"
                    ElseIf 54 <= bagilnot And bagilnot <= 58.99 Then
                        harfNotu = "BA"
                    ElseIf 49 <= bagilnot And bagilnot <= 53.99 Then
                        harfNotu = "BB"
                    ElseIf 44 <= bagilnot And bagilnot <= 48.99 Then
                        harfNotu = "CB"
                    ElseIf 39 <= bagilnot And bagilnot <= 43.99 Then
                        harfNotu = "CC"
                    ElseIf 34 <= bagilnot And bagilnot <= 38.99 Then
                        harfNotu = "DC"
                    ElseIf 29 <= bagilnot And bagilnot <= 33.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 29 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu

                Next
            End If

            If 62.5 < sinifort And sinifort <= 70 Then

                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 61 Then
                        harfNotu = "AA"
                    ElseIf 56 <= bagilnot And bagilnot <= 60.99 Then
                        harfNotu = "BA"
                    ElseIf 51 <= bagilnot And bagilnot <= 55.99 Then
                        harfNotu = "BB"
                    ElseIf 46 <= bagilnot And bagilnot <= 50.99 Then
                        harfNotu = "CB"
                    ElseIf 41 <= bagilnot And bagilnot <= 45.99 Then
                        harfNotu = "CC"
                    ElseIf 36 <= bagilnot And bagilnot <= 40.99 Then
                        harfNotu = "DC"
                    ElseIf 31 <= bagilnot And bagilnot <= 35.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 31 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If

            If 57.5 < sinifort And sinifort < 62.5 Then

                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 63 Then
                        harfNotu = "AA"
                    ElseIf 58 <= bagilnot And bagilnot <= 62.99 Then
                        harfNotu = "BA"
                    ElseIf 53 <= bagilnot And bagilnot <= 57.99 Then
                        harfNotu = "BB"
                    ElseIf 48 <= bagilnot And bagilnot <= 52.99 Then
                        harfNotu = "CB"
                    ElseIf 43 <= bagilnot And bagilnot <= 47.99 Then
                        harfNotu = "CC"
                    ElseIf 38 <= bagilnot And bagilnot <= 42.99 Then
                        harfNotu = "DC"
                    ElseIf 33 <= bagilnot And bagilnot <= 37.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 33 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If

            If 52.5 < sinifort And sinifort <= 57.5 Then

                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 65 Then
                        harfNotu = "AA"
                    ElseIf 60 <= bagilnot And bagilnot <= 64.99 Then
                        harfNotu = "BA"
                    ElseIf 55 <= bagilnot And bagilnot <= 59.99 Then
                        harfNotu = "BB"
                    ElseIf 50 <= bagilnot And bagilnot <= 54.99 Then
                        harfNotu = "CB"
                    ElseIf 45 <= bagilnot And bagilnot <= 49.99 Then
                        harfNotu = "CC"
                    ElseIf 40 <= bagilnot And bagilnot <= 44.99 Then
                        harfNotu = "DC"
                    ElseIf 35 <= bagilnot And bagilnot <= 39.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 35 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If

            If 47.5 < sinifort And sinifort <= 52.5 Then

                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 67 Then
                        harfNotu = "AA"
                    ElseIf 62 <= bagilnot And bagilnot <= 66.99 Then
                        harfNotu = "BA"
                    ElseIf 57 <= bagilnot And bagilnot <= 61.99 Then
                        harfNotu = "BB"
                    ElseIf 52 <= bagilnot And bagilnot <= 56.99 Then
                        harfNotu = "CB"
                    ElseIf 47 <= bagilnot And bagilnot <= 51.99 Then
                        harfNotu = "CC"
                    ElseIf 42 <= bagilnot And bagilnot <= 46.99 Then
                        harfNotu = "DC"
                    ElseIf 37 <= bagilnot And bagilnot <= 41.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 37 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If

            If 42.5 < sinifort And sinifort <= 47.5 Then

                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 69 Then
                        harfNotu = "AA"
                    ElseIf 64 <= bagilnot And bagilnot <= 68.99 Then
                        harfNotu = "BA"
                    ElseIf 59 <= bagilnot And bagilnot <= 63.99 Then
                        harfNotu = "BB"
                    ElseIf 54 <= bagilnot And bagilnot <= 58.99 Then
                        harfNotu = "CB"
                    ElseIf 49 <= bagilnot And bagilnot <= 53.99 Then
                        harfNotu = "CC"
                    ElseIf 44 <= bagilnot And bagilnot <= 48.99 Then
                        harfNotu = "DC"
                    ElseIf 39 <= bagilnot And bagilnot <= 43.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 39 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If

            If sinifort < 42.5 Then
                For i = 0 To DataGridView1.Rows.Count - 1
                    Dim bagilnot As Double = DataGridView1.Rows(i).Cells(8).Value
                    Dim harfNotu As String = DataGridView1.Rows(i).Cells(9).Value
                    If bagilnot >= 71 Then
                        harfNotu = "AA"
                    ElseIf 66 <= bagilnot And bagilnot <= 70.99 Then
                        harfNotu = "BA"
                    ElseIf 61 <= bagilnot And bagilnot <= 65.99 Then
                        harfNotu = "BB"
                    ElseIf 56 <= bagilnot And bagilnot <= 60.99 Then
                        harfNotu = "CB"
                    ElseIf 51 <= bagilnot And bagilnot <= 55.99 Then
                        harfNotu = "CC"
                    ElseIf 46 <= bagilnot And bagilnot <= 50.99 Then
                        harfNotu = "DC"
                    ElseIf 41 <= bagilnot And bagilnot <= 45.99 Then
                        harfNotu = "DD"
                    ElseIf bagilnot < 41 Then
                        harfNotu = "FF"
                    Else
                        MsgBox("Hatalı veri girişi")
                    End If

                    DataGridView1.Rows(i).Cells(9).Value = harfNotu
                Next
            End If

        Catch ex As Exception
            MsgBox("Hata meydana geldi. Ortalama değerin hesaplandığından emin olunuz.")
        End Try


    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click

        Dim sonuc As Double
        For i = 0 To DataGridView1.Rows.Count - 1
            Dim vize As Double = DataGridView1.Rows(i).Cells(5).Value
            Dim final As Double = DataGridView1.Rows(i).Cells(6).Value
            'sonuc = dll.ortalama(vize, final).ToString()

            'DataGridView1.Rows(i).Cells(7).Value = (dll.ortalama(vize, final))
            sonuc = (vize * 0.4) + (final * 0.6)
            DataGridView1.Rows(i).Cells(7).Value = sonuc.ToString
        Next
        Button3.Enabled = True
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
    End Sub

    Public Function sayiMi(deger As String)
        Try
            Convert.ToDouble(deger)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If DataGridView1.Rows.Count <> 0 Then
            Dim i As Integer = DataGridView1.CurrentRow.Index()
            TextBox1.Text = DataGridView1.Rows(i).Cells(2).Value.ToString
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            'Vize notu boş ise
            If String.IsNullOrEmpty(DataGridView1.Rows(i).Cells(5).Value) Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = DataGridView1.Rows(i).Cells(5).Value.ToString()
            End If

            'final notu boş ise 
            If String.IsNullOrEmpty(DataGridView1.Rows(i).Cells(6).Value) Then
                TextBox3.Text = ""
            Else
                TextBox3.Text = DataGridView1.Rows(i).Cells(6).Value
            End If
        End If
    End Sub
End Class
