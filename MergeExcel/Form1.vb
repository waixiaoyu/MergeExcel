Imports Microsoft.Office.Interop

Public Class Form1





    Private Sub BTN_ChooseOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_ChooseOutput.Click
        FolderBrowserDialog_Output.ShowDialog()
        TB_Output.Text = FolderBrowserDialog_Output.SelectedPath
    End Sub

    Private Sub BTN_Merge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_Merge.Click
        Dim myExcel As New Excel.Application()
        Dim workBook As Excel.Workbook '定义工作簿
        Dim sheet As Excel.Worksheet '定义工作表
        myExcel.Visible = False
        workBook = myExcel.Workbooks.Open("C:\Users\Administrator\Desktop\1.xlsx")
        sheet = workBook.Sheets(1)


        MessageBox.Show(sheet.Cells(1, 2).value())



        myExcel.Quit()



    End Sub
End Class
