<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Converter
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StandardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ScientificToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProgrammerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GraphingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FunctionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ConverterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(4, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(524, 24)
        Me.MenuStrip1.TabIndex = 27
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MenuToolStripMenuItem
        '
        Me.MenuToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StandardToolStripMenuItem, Me.ScientificToolStripMenuItem, Me.ProgrammerToolStripMenuItem, Me.GraphingToolStripMenuItem, Me.FunctionsToolStripMenuItem, Me.ConverterToolStripMenuItem})
        Me.MenuToolStripMenuItem.Name = "MenuToolStripMenuItem"
        Me.MenuToolStripMenuItem.Size = New System.Drawing.Size(50, 20)
        Me.MenuToolStripMenuItem.Text = "Menu"
        '
        'StandardToolStripMenuItem
        '
        Me.StandardToolStripMenuItem.Name = "StandardToolStripMenuItem"
        Me.StandardToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.StandardToolStripMenuItem.Text = "Standard"
        '
        'ScientificToolStripMenuItem
        '
        Me.ScientificToolStripMenuItem.Name = "ScientificToolStripMenuItem"
        Me.ScientificToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ScientificToolStripMenuItem.Text = "Scientific"
        '
        'ProgrammerToolStripMenuItem
        '
        Me.ProgrammerToolStripMenuItem.Name = "ProgrammerToolStripMenuItem"
        Me.ProgrammerToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ProgrammerToolStripMenuItem.Text = "Programmer"
        '
        'GraphingToolStripMenuItem
        '
        Me.GraphingToolStripMenuItem.Name = "GraphingToolStripMenuItem"
        Me.GraphingToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.GraphingToolStripMenuItem.Text = "Graphing"
        '
        'FunctionsToolStripMenuItem
        '
        Me.FunctionsToolStripMenuItem.Name = "FunctionsToolStripMenuItem"
        Me.FunctionsToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.FunctionsToolStripMenuItem.Text = "Functions"
        '
        'ConverterToolStripMenuItem
        '
        Me.ConverterToolStripMenuItem.Name = "ConverterToolStripMenuItem"
        Me.ConverterToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ConverterToolStripMenuItem.Text = "Converter"
        '
        'Converter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(524, 588)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "Converter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Calculator"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents MenuToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StandardToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ScientificToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProgrammerToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GraphingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FunctionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ConverterToolStripMenuItem As ToolStripMenuItem
End Class
