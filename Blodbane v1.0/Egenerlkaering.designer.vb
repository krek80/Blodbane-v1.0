<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Egenerlkaering
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
        Me.components = New System.ComponentModel.Container()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.rdbtnJa = New System.Windows.Forms.RadioButton()
        Me.rdbtnNei = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnForrige = New System.Windows.Forms.Button()
        Me.btnNeste = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.GroupBox1.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Label1"
        '
        'rdbtnJa
        '
        Me.rdbtnJa.AutoSize = True
        Me.rdbtnJa.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.rdbtnJa.Location = New System.Drawing.Point(32, 153)
        Me.rdbtnJa.Name = "rdbtnJa"
        Me.rdbtnJa.Size = New System.Drawing.Size(47, 22)
        Me.rdbtnJa.TabIndex = 1
        Me.rdbtnJa.TabStop = True
        Me.rdbtnJa.Text = "Ja"
        Me.rdbtnJa.UseVisualStyleBackColor = True
        '
        'rdbtnNei
        '
        Me.rdbtnNei.AutoCheck = False
        Me.rdbtnNei.AutoSize = True
        Me.rdbtnNei.Location = New System.Drawing.Point(32, 181)
        Me.rdbtnNei.Name = "rdbtnNei"
        Me.rdbtnNei.Size = New System.Drawing.Size(50, 21)
        Me.rdbtnNei.TabIndex = 2
        Me.rdbtnNei.TabStop = True
        Me.rdbtnNei.Text = "Nei"
        Me.rdbtnNei.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(324, 266)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Label 2"
        '
        'btnForrige
        '
        Me.btnForrige.Location = New System.Drawing.Point(15, 228)
        Me.btnForrige.Name = "btnForrige"
        Me.btnForrige.Size = New System.Drawing.Size(166, 23)
        Me.btnForrige.TabIndex = 4
        Me.btnForrige.Text = "<< Forrige spørsmål"
        Me.btnForrige.UseVisualStyleBackColor = True
        '
        'btnNeste
        '
        Me.btnNeste.Location = New System.Drawing.Point(204, 228)
        Me.btnNeste.Name = "btnNeste"
        Me.btnNeste.Size = New System.Drawing.Size(175, 23)
        Me.btnNeste.TabIndex = 5
        Me.btnNeste.Text = "Neste spørsmål >>"
        Me.btnNeste.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.btnNeste)
        Me.GroupBox1.Controls.Add(Me.rdbtnJa)
        Me.GroupBox1.Controls.Add(Me.btnForrige)
        Me.GroupBox1.Controls.Add(Me.rdbtnNei)
        Me.GroupBox1.Location = New System.Drawing.Point(104, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(409, 297)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Egenerklæring"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(674, 479)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents rdbtnJa As RadioButton
    Friend WithEvents rdbtnNei As RadioButton
    Friend WithEvents Label2 As Label
    Friend WithEvents btnForrige As Button
    Friend WithEvents btnNeste As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents BindingSource1 As BindingSource
End Class
