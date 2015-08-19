Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Esri_Trisoft_TopicStatus_Tool
    Public GUID As New ArrayList
    Public VersionNum As New ArrayList
    Public Resolution As New ArrayList
    Public pubName As New ArrayList
    Public langName As New ArrayList
    Private f_Properties As New Hashtable

    Private Sub Browse1_Click(sender As Object, e As EventArgs) Handles Browse1.Click
        OpenFileDialog.Title = "Select csv file as input file"
        OpenFileDialog.FileName = ""
        OpenFileDialog.Filter = "csv files | *.csv"
        OpenFileDialog.InitialDirectory = "C:\"
        If OpenFileDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = OpenFileDialog.FileName
        End If
    End Sub

    Private Sub Browse2_Click(sender As Object, e As EventArgs) Handles Browse2.Click
        If FolderBrowserDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox2.Text = FolderBrowserDialog.SelectedPath
        End If
    End Sub

    Private Sub Run_Click(sender As Object, e As EventArgs) Handles Run.Click
        Dim usrDir As String = Directory.GetCurrentDirectory
        Dim toolPropertyFile As String = usrDir & "\tool.properities"
        If Not File.Exists(toolPropertyFile) Then
            MsgBox("Missing property file in user directory")
            Return
        End If
        Dim wholeProperty As String = My.Computer.FileSystem.ReadAllText(toolPropertyFile)
        Dim lineDataProperty() As String = Split(wholeProperty, vbNewLine)

        For Each lineTextProperty As String In lineDataProperty
            f_Properties.Add(lineTextProperty.Split("=")(0), lineTextProperty.Split("=")(1))
        Next

        Dim UserName As String = f_Properties.Item("USERNAME")
        Dim Password As String = f_Properties.Item("PASSWORD")
        Dim URL As String = f_Properties.Item("HOMEURL")

        Dim PubType As String = ""
        If CheckedListBox1.SelectedIndex <> -1 Then
            For k As Integer = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemCheckState(k, CheckState.Unchecked)
            Next
            CheckedListBox1.SetItemCheckState(CheckedListBox1.SelectedIndex, CheckState.Checked)
            PubType = CheckedListBox1.SelectedItem.ToString()
        Else
            MsgBox("You have to select one item in the checkbox.")
            Return
        End If

        Dim filePath As String
        Dim savePath As String
        Dim wholeFile As String
        Dim lineData() As String
        Dim fieldData() As String
        filePath = TextBox1.Text
        savePath = TextBox2.Text
        'filePath = "C:\Users\wei7771\Desktop\TopicStatus_Tool\test.csv"
        'savePath = "C:\Users\wei7771\Desktop\TopicStatus_Tool"
        Dim excelFileName = filePath.Substring(filePath.LastIndexOf("\") + 1)

        Dim Context As String = ""
        Dim testobj As New IshObjs(UserName, Password, URL)
        testobj.ISHAppObj.Login("InfoShareAuthor", UserName, Password, Context)
        Dim test As New IshPubOutput(UserName, Password, URL)
        Dim outputFileLang As String = ""
        Dim i As Integer
        If Not String.IsNullOrEmpty(filePath) And Not String.IsNullOrEmpty(savePath) Then
            wholeFile = My.Computer.FileSystem.ReadAllText(filePath)
            lineData = Split(wholeFile, vbNewLine)
            For Each lineOfText As String In lineData
                fieldData = lineOfText.Split(",")
                For i = 0 To fieldData.Length - 1
                    Select Case i
                        Case 0
                            GUID.Add(fieldData(0))
                        Case 1
                            VersionNum.Add(fieldData(1))
                        Case 2
                            Resolution.Add(fieldData(2))
                        Case 3
                            langName.Add(fieldData(3))
                        Case 4
                            pubName.Add(fieldData(4))
                    End Select
                Next
            Next lineOfText
        Else
            MessageBox.Show("The file path is empty.")
        End If

        GUID.RemoveAt(GUID.Count - 1)

        Dim Log_Path = savePath + "\log.txt"
        If Not File.Exists(Log_Path) Then
            File.Create(Log_Path).Dispose()
        End If

        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oSheet As Excel.Worksheet
        'Start a new workbook in Excel.
        oExcel = CreateObject("Excel.Application")
        oBook = oExcel.Workbooks.Add
        'Add data to cells of the first worksheet in the new workbook.
        oSheet = oBook.ActiveSheet
        'oSheet = oBook.Worksheets(1)
        oSheet.Cells(1, 1).Value = "Pub GUID"
        oSheet.Cells(1, 2).Value = "Pub Name"
        oSheet.Cells(1, 3).Value = "Pub Version"
        oSheet.Cells(1, 4).Value = "Resolution"
        oSheet.Cells(1, 5).Value = "Topic GUID"
        oSheet.Cells(1, 6).Value = "Topic Type"
        oSheet.Cells(1, 7).Value = "Topic Name"
        oSheet.Cells(1, 8).Value = "Topic Version"
        oSheet.Cells(1, 9).Value = "Topic Status"
        oSheet.Cells(1, 10).Value = "Language"
        oSheet.Cells(1, 11).Value = "EN Release Date"
        oSheet.Cells(1, 12).Value = "Author"
        oSheet.Cells(1, 13).Value = "Enable L10N"
        oSheet.Cells(1, 14).Value = "Translated Date"
        oSheet.Cells(1, 15).Value = "In Translation Date"
        oSheet.Cells(1, 16).Value = "Comments"
        Dim excelRow As Integer = 2

        Dim j As Integer
        For j = 0 To GUID.Count - 1
            Dim id As String = GUID(j).ToString()
            Dim version As String = VersionNum(j).ToString()
            Dim reso As String = Resolution(j).ToString()
            Dim lang As String = langName(j).ToString()
            Dim name As String = pubName(j).ToString()
            outputFileLang = lang
            test.GetStatus(id, name, version, reso, lang, PubType, savePath, Log_Path, excelRow, oSheet)
        Next
        oSheet.Range("K1", "K" & excelRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        oSheet.Range("N1", "N" & excelRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        oSheet.Range("O1", "O" & excelRow).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        excelFileName = excelFileName.Replace(".csv", "_TopicStatus_" & outputFileLang.ToUpper() & ".xlsx")
        oSheet.Range("A1:P1").EntireColumn.AutoFit()
        oBook.SaveAs(savePath & "\" & excelFileName)
        oSheet = Nothing
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing
        GC.Collect()
        MsgBox("The Task is Done")
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Dim item As Integer = CheckedListBox1.SelectedIndex
        CheckedListBox1.SetItemCheckState(item, CheckState.Checked)
        If CheckedListBox1.CheckedItems.Count > 1 Then
            For k As Integer = 0 To CheckedListBox1.Items.Count - 1
                If k <> item And CheckedListBox1.GetItemChecked(k) = True Then
                    CheckedListBox1.SetItemCheckState(k, CheckState.Unchecked)
                End If
            Next
        End If
    End Sub
End Class
