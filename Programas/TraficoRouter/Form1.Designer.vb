<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.ResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.TamañoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.IntervaloToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.GrosorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.VentanaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TransparenciaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem5 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem3 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem4 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem6 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem7 = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripMenuItem8 = New System.Windows.Forms.ToolStripMenuItem
        Me.SiempreEncimaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Chart1
        '
        Me.Chart1.Dock = System.Windows.Forms.DockStyle.Fill
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(0, 24)
        Me.Chart1.Name = "Chart1"
        Me.Chart1.Size = New System.Drawing.Size(292, 242)
        Me.Chart1.TabIndex = 0
        Me.Chart1.Text = "Chart1"
        '
        'Timer1
        '
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.VentanaToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(292, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ResetToolStripMenuItem, Me.ToolStripSeparator1, Me.TamañoToolStripMenuItem, Me.IntervaloToolStripMenuItem, Me.GrosorToolStripMenuItem})
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(53, 20)
        Me.ToolStripMenuItem1.Text = "Grafica"
        '
        'ResetToolStripMenuItem
        '
        Me.ResetToolStripMenuItem.Name = "ResetToolStripMenuItem"
        Me.ResetToolStripMenuItem.Size = New System.Drawing.Size(129, 22)
        Me.ResetToolStripMenuItem.Text = "Reset"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(126, 6)
        '
        'TamañoToolStripMenuItem
        '
        Me.TamañoToolStripMenuItem.Name = "TamañoToolStripMenuItem"
        Me.TamañoToolStripMenuItem.Size = New System.Drawing.Size(129, 22)
        Me.TamañoToolStripMenuItem.Text = "Tamaño"
        '
        'IntervaloToolStripMenuItem
        '
        Me.IntervaloToolStripMenuItem.Name = "IntervaloToolStripMenuItem"
        Me.IntervaloToolStripMenuItem.Size = New System.Drawing.Size(129, 22)
        Me.IntervaloToolStripMenuItem.Text = "Intervalo"
        '
        'GrosorToolStripMenuItem
        '
        Me.GrosorToolStripMenuItem.Name = "GrosorToolStripMenuItem"
        Me.GrosorToolStripMenuItem.Size = New System.Drawing.Size(129, 22)
        Me.GrosorToolStripMenuItem.Text = "Grosor"
        '
        'VentanaToolStripMenuItem
        '
        Me.VentanaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TransparenciaToolStripMenuItem, Me.SiempreEncimaToolStripMenuItem})
        Me.VentanaToolStripMenuItem.Name = "VentanaToolStripMenuItem"
        Me.VentanaToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.VentanaToolStripMenuItem.Text = "Ventana"
        '
        'TransparenciaToolStripMenuItem
        '
        Me.TransparenciaToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem5, Me.ToolStripMenuItem3, Me.ToolStripMenuItem4, Me.ToolStripMenuItem6, Me.ToolStripMenuItem7, Me.ToolStripMenuItem8})
        Me.TransparenciaToolStripMenuItem.Name = "TransparenciaToolStripMenuItem"
        Me.TransparenciaToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.TransparenciaToolStripMenuItem.Text = "Transparencia"
        '
        'ToolStripMenuItem5
        '
        Me.ToolStripMenuItem5.Name = "ToolStripMenuItem5"
        Me.ToolStripMenuItem5.Size = New System.Drawing.Size(108, 22)
        Me.ToolStripMenuItem5.Text = "0%"
        '
        'ToolStripMenuItem3
        '
        Me.ToolStripMenuItem3.Name = "ToolStripMenuItem3"
        Me.ToolStripMenuItem3.Size = New System.Drawing.Size(108, 22)
        Me.ToolStripMenuItem3.Text = "5%"
        '
        'ToolStripMenuItem4
        '
        Me.ToolStripMenuItem4.Name = "ToolStripMenuItem4"
        Me.ToolStripMenuItem4.Size = New System.Drawing.Size(108, 22)
        Me.ToolStripMenuItem4.Text = "10%"
        '
        'ToolStripMenuItem6
        '
        Me.ToolStripMenuItem6.Name = "ToolStripMenuItem6"
        Me.ToolStripMenuItem6.Size = New System.Drawing.Size(108, 22)
        Me.ToolStripMenuItem6.Text = "20%"
        '
        'ToolStripMenuItem7
        '
        Me.ToolStripMenuItem7.Name = "ToolStripMenuItem7"
        Me.ToolStripMenuItem7.Size = New System.Drawing.Size(108, 22)
        Me.ToolStripMenuItem7.Text = "30%"
        '
        'ToolStripMenuItem8
        '
        Me.ToolStripMenuItem8.Name = "ToolStripMenuItem8"
        Me.ToolStripMenuItem8.Size = New System.Drawing.Size(108, 22)
        Me.ToolStripMenuItem8.Text = "40%"
        '
        'SiempreEncimaToolStripMenuItem
        '
        Me.SiempreEncimaToolStripMenuItem.Name = "SiempreEncimaToolStripMenuItem"
        Me.SiempreEncimaToolStripMenuItem.Size = New System.Drawing.Size(159, 22)
        Me.SiempreEncimaToolStripMenuItem.Text = "Siempre encima"
        '
        'BackgroundWorker1
        '
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "Internet"
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TamañoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VentanaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransparenciaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem5 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem6 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem7 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem8 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SiempreEncimaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IntervaloToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GrosorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
