<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Esri_Trisoft_TopicStatus_Tool
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Browse1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Browse2 = New System.Windows.Forms.Button()
        Me.Run = New System.Windows.Forms.Button()
        Me.Open_Folder = New System.Windows.Forms.Button()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(58, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(315, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Please select the xlsx or csv file that contains a list of publications"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(59, 88)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(431, 20)
        Me.TextBox1.TabIndex = 6
        '
        'Browse1
        '
        Me.Browse1.Location = New System.Drawing.Point(575, 72)
        Me.Browse1.Name = "Browse1"
        Me.Browse1.Size = New System.Drawing.Size(109, 36)
        Me.Browse1.TabIndex = 7
        Me.Browse1.Text = "Browse..."
        Me.Browse1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(57, 143)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(203, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Please select the folder to store the report"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(59, 187)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(429, 20)
        Me.TextBox2.TabIndex = 9
        '
        'Browse2
        '
        Me.Browse2.Location = New System.Drawing.Point(575, 171)
        Me.Browse2.Name = "Browse2"
        Me.Browse2.Size = New System.Drawing.Size(110, 36)
        Me.Browse2.TabIndex = 10
        Me.Browse2.Text = "Browse..."
        Me.Browse2.UseVisualStyleBackColor = True
        '
        'Run
        '
        Me.Run.Location = New System.Drawing.Point(61, 365)
        Me.Run.Name = "Run"
        Me.Run.Size = New System.Drawing.Size(134, 38)
        Me.Run.TabIndex = 11
        Me.Run.Text = "Run"
        Me.Run.UseVisualStyleBackColor = True
        '
        'Open_Folder
        '
        Me.Open_Folder.Location = New System.Drawing.Point(375, 365)
        Me.Open_Folder.Name = "Open_Folder"
        Me.Open_Folder.Size = New System.Drawing.Size(143, 38)
        Me.Open_Folder.TabIndex = 12
        Me.Open_Folder.Text = "Open Folder"
        Me.Open_Folder.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Items.AddRange(New Object() {"Both Image and Topic", "Image Only", "Topic Only"})
        Me.CheckedListBox1.Location = New System.Drawing.Point(204, 275)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(155, 64)
        Me.CheckedListBox1.TabIndex = 13
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(58, 242)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(276, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Please select the topic type you want for the topic status:"
        '
        'Esri_Trisoft_TopicStatus_Tool
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(769, 469)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CheckedListBox1)
        Me.Controls.Add(Me.Open_Folder)
        Me.Controls.Add(Me.Run)
        Me.Controls.Add(Me.Browse2)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Browse1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Esri_Trisoft_TopicStatus_Tool"
        Me.Text = "Esri_Trisoft_Topic_Status_Tool"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Browse1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Browse2 As System.Windows.Forms.Button
    Friend WithEvents Run As System.Windows.Forms.Button
    Friend WithEvents Open_Folder As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
