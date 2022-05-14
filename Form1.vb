Imports Microsoft.Office.Interop
Public Class Form1

    Dim excel_uyg As New Excel.Application
    Dim exc_workbook As Excel.Workbook
    Dim exc_sheet As Excel.Worksheet

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Algoritma")
        ComboBox1.Items.Add(".Net'e Giriş")
        ComboBox1.Items.Add("Veritabanı 1")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "Algoritma" Then
            Label1.Text = ComboBox1.SelectedItem
            algoritmaToDataGrid()

        ElseIf ComboBox1.SelectedItem = ".Net'e Giriş" Then
            Label1.Text = ComboBox1.SelectedItem
            netegirisToDataGrid()

        ElseIf ComboBox1.SelectedItem = "Veritabanı 1" Then
            Label1.Text = ComboBox1.SelectedItem
            veritabaniToDataGrid()

        Else messagebox.Show("Böyle bir ders kayıtlı değil.")
        End If

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        ' ReadOnly Combobox
        e.SuppressKeyPress = True
    End Sub

    Public Sub algoritmaToDataGrid()
        excel_uyg.Workbooks.Open("D:\\algoritma.xlsx")
        exc_workbook = excel_uyg.ActiveWorkbook
        exc_sheet = exc_workbook.ActiveSheet

        Dim sayac, i As Integer
        sayac = 0
        i = 1
        Do While i < 1000
            sayac = sayac + 1
            If exc_sheet.Cells(i, 1).Value = "" Then
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
        excel_uyg.Workbooks.Open("D:\\netegiris.xlsx")
        exc_workbook = excel_uyg.ActiveWorkbook
        exc_sheet = exc_workbook.ActiveSheet

        Dim sayac, i As Integer
        sayac = 0
        i = 1
        Do While i < 1000
            sayac = sayac + 1
            If exc_sheet.Cells(i, 1).Value = "" Then
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

    Public Sub veritabaniToDataGrid()
        excel_uyg.Workbooks.Open("D:\\veritabani.xlsx")
        exc_workbook = excel_uyg.ActiveWorkbook
        exc_sheet = exc_workbook.ActiveSheet

        Dim sayac, i As Integer
        sayac = 0
        i = 1
        Do While i < 1000
            sayac = sayac + 1
            If exc_sheet.Cells(i, 1).Value = "" Then
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
End Class
