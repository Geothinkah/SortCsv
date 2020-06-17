Public Class Form1


    '    This Is in a Windows Forms application to make the example easy to reproduce, but note that 
    'the functions which Do the work could easily be pasted into your console project.

    'The SortCsv() method takes the path Of the input file, the path Of the output file, 
    'And a Integer Or array Of integers representing the column numbers To sort On (In the order specified by the array).  
    'The columns numbers are treated at being 1 based.

    'The example also allows For delimiters In the data, which takes a little extra time And could be removed If you 
    'knew that your data would Not have a comma In it.


    'Sets up a large test file (200,000 records; ~3.5 MB)  
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        If Not IO.File.Exists("c:\test\testdata.csv") Then
            Dim sb As New Text.StringBuilder
            For i As Integer = 1 To 4
                sb.AppendLine("160001, John, Doe")
                sb.AppendLine("150227,Sue,Smith")
                sb.AppendLine("160102,Ben,Cartwright")
                sb.AppendLine("120222,Bill,Jones")
            Next

            IO.File.WriteAllText("c:\test\testdata.csv", sb.ToString)
        End If
    End Sub

    'Uses Button1 and DataGridView1 to display results of example  
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Button1.Click
        DataGridView1.AutoGenerateColumns = True
        ' SortCsv("c:\test\testdata.csv", "c:\test\testdataresult.csv", New Integer() {2, 3})
        SortCsv("c:\test\testdata.csv", "c:\test\testdataresult.csv", New Integer() {1, 3})
        DataGridView1.DataSource = CsvToTable("c:\test\testdataresult.csv")
    End Sub

    'Sorts the csv by turning it into a datatable, applying a view filter, and then writing the result back to a csv  
    Public Sub SortCsv(ByVal sourceFile As String, ByVal destinationFile As String, ByVal ParamArray sortColumns() As Integer)
        Dim dt As DataTable = CsvToTable(sourceFile)
        If sortColumns.Length > 0 Then
            Dim sortStr As String = String.Empty
            For i As Integer = 0 To sortColumns.Length - 1
                If sortStr.Length > 0 Then sortStr &= ", "
                sortStr &= "Column" & sortColumns(i).ToString
            Next
            dt.DefaultView.Sort = sortStr
        End If
        TableToCSV(dt.DefaultView.ToTable, destinationFile)
    End Sub

    'Parses a csv into a datatable  
    Private Function CsvToTable(ByVal filePathName As String) As DataTable
        Dim result As New DataTable
        If IO.File.Exists(filePathName) Then
            Dim parser As New FileIO.TextFieldParser(filePathName)
            parser.Delimiters = New String() {","}
            parser.HasFieldsEnclosedInQuotes = True 'use if data may contain delimiters  
            parser.TextFieldType = FileIO.FieldType.Delimited
            parser.TrimWhiteSpace = True

            While Not parser.EndOfData
                AddValuesToTable(parser.ReadFields, result)
            End While

            parser.Close()
        End If
        Return result
    End Function

    'Writes a datatable back into a csv  
    Private Sub TableToCSV(ByVal sourceTable As DataTable, ByVal filePathName As String)
        Dim sb As New Text.StringBuilder
        For Each dr As DataRow In sourceTable.Rows
            sb.AppendLine(String.Join(",", Array.ConvertAll(dr.ItemArray,
            Function(o As Object) If(o.ToString.Contains(","),
            ControlChars.Quote & o.ToString & ControlChars.Quote, o.ToString))))
        Next
        IO.File.WriteAllText(filePathName, sb.ToString)
    End Sub

    'Ensures a datatable can hold an array of values and then adds a new row  
    Private Sub AddValuesToTable(ByVal source() As String, ByVal destination As DataTable)
        Dim existing As Integer = destination.Columns.Count
        For i As Integer = 0 To source.Length - existing - 1
            destination.Columns.Add("Column" & (existing + 1 + i).ToString, GetType(String))
        Next
        destination.Rows.Add(source)
    End Sub
End Class
