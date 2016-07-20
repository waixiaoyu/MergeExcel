<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BTN_Merge = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.BTN_ChooseInput = New System.Windows.Forms.Button()
        Me.TB_Input = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.BTN_ChooseOutput = New System.Windows.Forms.Button()
        Me.TB_Output = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FolderBrowserDialog_Input = New System.Windows.Forms.FolderBrowserDialog()
        Me.FolderBrowserDialog_Output = New System.Windows.Forms.FolderBrowserDialog()
        Me.TB_OutputFileName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label_Count = New System.Windows.Forms.Label()
        Me.BTN_Open = New System.Windows.Forms.Button()
        Me.BTN_OpenDir = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'BTN_Merge
        '
        Me.BTN_Merge.Location = New System.Drawing.Point(161, 126)
        Me.BTN_Merge.Name = "BTN_Merge"
        Me.BTN_Merge.Size = New System.Drawing.Size(154, 55)
        Me.BTN_Merge.TabIndex = 1
        Me.BTN_Merge.Text = "开始合并"
        Me.BTN_Merge.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 37)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(143, 12)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "选择存储Excel的文件夹："
        '
        'BTN_ChooseInput
        '
        Me.BTN_ChooseInput.Location = New System.Drawing.Point(366, 32)
        Me.BTN_ChooseInput.Name = "BTN_ChooseInput"
        Me.BTN_ChooseInput.Size = New System.Drawing.Size(62, 23)
        Me.BTN_ChooseInput.TabIndex = 10
        Me.BTN_ChooseInput.Text = "浏览"
        Me.BTN_ChooseInput.UseVisualStyleBackColor = True
        '
        'TB_Input
        '
        Me.TB_Input.Location = New System.Drawing.Point(161, 34)
        Me.TB_Input.Name = "TB_Input"
        Me.TB_Input.ReadOnly = True
        Me.TB_Input.Size = New System.Drawing.Size(195, 21)
        Me.TB_Input.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(89, 12)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "选择输出路径："
        '
        'BTN_ChooseOutput
        '
        Me.BTN_ChooseOutput.Location = New System.Drawing.Point(366, 66)
        Me.BTN_ChooseOutput.Name = "BTN_ChooseOutput"
        Me.BTN_ChooseOutput.Size = New System.Drawing.Size(62, 23)
        Me.BTN_ChooseOutput.TabIndex = 13
        Me.BTN_ChooseOutput.Text = "浏览"
        Me.BTN_ChooseOutput.UseVisualStyleBackColor = True
        '
        'TB_Output
        '
        Me.TB_Output.Location = New System.Drawing.Point(107, 68)
        Me.TB_Output.Name = "TB_Output"
        Me.TB_Output.ReadOnly = True
        Me.TB_Output.Size = New System.Drawing.Size(249, 21)
        Me.TB_Output.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(227, 12)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "将需要合并的Excel放在同一个目录下即可"
        '
        'TB_OutputFileName
        '
        Me.TB_OutputFileName.Location = New System.Drawing.Point(143, 99)
        Me.TB_OutputFileName.Name = "TB_OutputFileName"
        Me.TB_OutputFileName.Size = New System.Drawing.Size(172, 21)
        Me.TB_OutputFileName.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 102)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(125, 12)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "输出文件名（可选）："
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(315, 102)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(35, 12)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = ".xlsx"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(354, 102)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(113, 12)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "(默认为：new.xlsx)"
        '
        'Label_Count
        '
        Me.Label_Count.AutoSize = True
        Me.Label_Count.Location = New System.Drawing.Point(334, 141)
        Me.Label_Count.Name = "Label_Count"
        Me.Label_Count.Size = New System.Drawing.Size(0, 12)
        Me.Label_Count.TabIndex = 20
        '
        'BTN_Open
        '
        Me.BTN_Open.Enabled = False
        Me.BTN_Open.Location = New System.Drawing.Point(37, 126)
        Me.BTN_Open.Name = "BTN_Open"
        Me.BTN_Open.Size = New System.Drawing.Size(90, 23)
        Me.BTN_Open.TabIndex = 21
        Me.BTN_Open.Text = "打开合并结果"
        Me.BTN_Open.UseVisualStyleBackColor = True
        '
        'BTN_OpenDir
        '
        Me.BTN_OpenDir.Location = New System.Drawing.Point(37, 158)
        Me.BTN_OpenDir.Name = "BTN_OpenDir"
        Me.BTN_OpenDir.Size = New System.Drawing.Size(90, 23)
        Me.BTN_OpenDir.TabIndex = 22
        Me.BTN_OpenDir.Text = "打开文件夹"
        Me.BTN_OpenDir.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 193)
        Me.Controls.Add(Me.BTN_OpenDir)
        Me.Controls.Add(Me.BTN_Open)
        Me.Controls.Add(Me.Label_Count)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TB_OutputFileName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.BTN_ChooseOutput)
        Me.Controls.Add(Me.TB_Output)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.BTN_ChooseInput)
        Me.Controls.Add(Me.TB_Input)
        Me.Controls.Add(Me.BTN_Merge)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "MergeExcel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BTN_Merge As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BTN_ChooseInput As System.Windows.Forms.Button
    Friend WithEvents TB_Input As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BTN_ChooseOutput As System.Windows.Forms.Button
    Friend WithEvents TB_Output As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog_Input As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents FolderBrowserDialog_Output As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents TB_OutputFileName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label_Count As System.Windows.Forms.Label
    Friend WithEvents BTN_Open As System.Windows.Forms.Button
    Friend WithEvents BTN_OpenDir As System.Windows.Forms.Button

End Class
