Imports Microsoft.Office.Interop
Imports System.IO

Public Class Form1

    Private Sub BTN_ChooseInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ChooseInput.Click
        FolderBrowserDialog_Input.ShowDialog()
        TB_Input.Text = FolderBrowserDialog_Input.SelectedPath
        inputPath = FolderBrowserDialog_Input.SelectedPath
    End Sub
    Private Sub BTN_ChooseOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ChooseOutput.Click
        FolderBrowserDialog_Output.ShowDialog()
        TB_Output.Text = FolderBrowserDialog_Output.SelectedPath
        outputPath = FolderBrowserDialog_Input.SelectedPath
    End Sub
    Dim inputPath As String
    Dim outputPath As String
    Dim outputFileName As String = "\new.xlsx"
    Dim totalCount As Int16
    Dim currentCount As Int16 = 1
    Private Sub BTN_Merge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Merge.Click
        BTN_Open.Enabled = False
        Label_Count.Text = "准备开始..."
        If TB_OutputFileName.Text <> Nothing Then
            outputFileName = "\" + TB_OutputFileName.Text + ".xlsx"
        End If
        xlRows = 2

        If (inputPath = Nothing Or outputPath = Nothing) Then
            MessageBox.Show("请先选择输入输出路径！")
            Return
        End If
        createNewExcel()
        Dim dir As System.IO.DirectoryInfo = New DirectoryInfo(inputPath)
        totalCount = dir.GetFiles("*.xlsx", SearchOption.AllDirectories).Length
        For Each item As System.IO.FileInfo In dir.GetFiles("*.xlsx", SearchOption.AllDirectories)
            If (item.Attributes And FileAttributes.Hidden) <> FileAttributes.Hidden Then '除去隐藏文件
                ''进度展示
                Label_Count.Text = "正在进行..." + CStr(currentCount) + "/" + CStr(totalCount)
                currentCount += 1
                ''
                Dim myExcel As New Excel.Application()
                Dim workBook As Excel.Workbook '定义工作簿
                Dim sheet As Excel.Worksheet '定义工作表
                Dim columns As Int16
                Dim rows As Int16
                myExcel.Visible = False
                ''
                workBook = myExcel.Workbooks.Open(item.FullName)
                sheet = workBook.Sheets(1)
                columns = sheet.UsedRange.Cells.Columns.Count
                rows = sheet.UsedRange.Cells.Rows.Count
                For i = 2 To rows
                    xlColumns = 1
                    For j = 1 To columns
                        xlSheet.Cells(xlRows, xlColumns).value = sheet.Cells(i, j).value '往新文件中写入数据
                        xlColumns += 1
                    Next
                    xlRows += 1
                Next
                sheet = Nothing
                workBook.Close(True)
                myExcel.Quit()
                workBook = Nothing
                myExcel = Nothing
                ''
            End If
        Next

        ''''''''''''''''''''''''''''''补充表头
        Dim di As IO.DirectoryInfo = New IO.DirectoryInfo(inputPath)
        ''取所有文件
        Dim fii() As IO.FileInfo = di.GetFiles("*.xlsx", IO.SearchOption.AllDirectories)
        createHeader(fii)
        '''''''''''''''''''''''''''''''''''''''

        closeNewExcel()
        Label_Count.Text = "完成！"

        MessageBox.Show("合并完成！")
        BTN_Open.Enabled = True
        ProcessKill()
        GC.Collect()

    End Sub
    Dim xlApp As New Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim xlColumns As Int16 = 1
    Dim xlRows As Int16 = 2
    Private Sub createNewExcel()
        '新建excel
        xlApp = CType(CreateObject("Excel.Application"), Excel.Application) '创建EXCEL对象 
        xlBook = xlApp.Workbooks.Add()
        xlSheet = CType(xlBook.Worksheets("sheet1"), Excel.Worksheet) '设置活动工作表 
    End Sub
    Private Sub createHeader(ByVal fii() As IO.FileInfo)
        Dim myExcel As New Excel.Application()
        Dim workBook As Excel.Workbook '定义工作簿
        Dim sheet As Excel.Worksheet '定义工作表
        Dim columns As Int16
        Dim rows As Int16
        myExcel.Visible = False
        ''
        workBook = myExcel.Workbooks.Open(fii(0).FullName)
        sheet = workBook.Sheets(1)
        columns = sheet.UsedRange.Cells.Columns.Count
        rows = sheet.UsedRange.Cells.Rows.Count
        xlColumns = 1
        For j = 1 To columns
            xlSheet.Cells(1, xlColumns).value = sheet.Cells(1, j).value '往新文件中写入数据
            xlColumns += 1
        Next
        sheet = Nothing
        workBook.Close(True)
        myExcel.Quit()
        workBook = Nothing
        myExcel = Nothing
    End Sub
    Private Sub closeNewExcel()
        ''关闭新文件
        xlBook.SaveAs(outputPath + outputFileName)
        xlSheet = Nothing
        xlBook.Close(True)
        xlApp.Quit()
        xlBook = Nothing
        xlApp = Nothing
    End Sub
    Private Sub ProcessKill()
        Dim p As System.Diagnostics.Process
        p = New System.Diagnostics.Process
        For Each p In System.Diagnostics.Process.GetProcesses()
            If p.ProcessName.ToUpper() = "EXCEL" Then
                p.Kill()
            End If
        Next
    End Sub


    Private Sub BTN_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Open.Click
        Dim myExcel As New Excel.Application()
        Dim workBook As Excel.Workbook '定义工作簿
        myExcel.Visible = True
        ''
        workBook = myExcel.Workbooks.Open(outputPath + outputFileName)
        BTN_Open.Enabled = False
    End Sub
End Class
