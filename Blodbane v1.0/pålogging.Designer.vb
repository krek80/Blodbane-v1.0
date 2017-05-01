<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class pålogging
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
        Me.btnAnsattPålogg = New System.Windows.Forms.Button()
        Me.txtAnsattBrNavn = New System.Windows.Forms.TextBox()
        Me.txtAnsattPassord = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnAnsattPålogg
        '
        Me.btnAnsattPålogg.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAnsattPålogg.Location = New System.Drawing.Point(36, 197)
        Me.btnAnsattPålogg.Margin = New System.Windows.Forms.Padding(4)
        Me.btnAnsattPålogg.Name = "btnAnsattPålogg"
        Me.btnAnsattPålogg.Size = New System.Drawing.Size(374, 58)
        Me.btnAnsattPålogg.TabIndex = 0
        Me.btnAnsattPålogg.Text = "LOGG PÅ"
        Me.btnAnsattPålogg.UseVisualStyleBackColor = True
        '
        'txtAnsattBrNavn
        '
        Me.txtAnsattBrNavn.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAnsattBrNavn.Location = New System.Drawing.Point(152, 121)
        Me.txtAnsattBrNavn.Margin = New System.Windows.Forms.Padding(4)
        Me.txtAnsattBrNavn.Name = "txtAnsattBrNavn"
        Me.txtAnsattBrNavn.Size = New System.Drawing.Size(258, 27)
        Me.txtAnsattBrNavn.TabIndex = 1
        '
        'txtAnsattPassord
        '
        Me.txtAnsattPassord.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAnsattPassord.Location = New System.Drawing.Point(152, 154)
        Me.txtAnsattPassord.Margin = New System.Windows.Forms.Padding(4)
        Me.txtAnsattPassord.Name = "txtAnsattPassord"
        Me.txtAnsattPassord.Size = New System.Drawing.Size(258, 27)
        Me.txtAnsattPassord.TabIndex = 2
        Me.txtAnsattPassord.UseSystemPasswordChar = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(32, 123)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Brukernavn"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(32, 157)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Passord"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 11)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(347, 94)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Pålogging for ansatte i blodbanken"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pålogging
        '
        Me.AcceptButton = Me.btnAnsattPålogg
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(445, 292)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtAnsattPassord)
        Me.Controls.Add(Me.txtAnsattBrNavn)
        Me.Controls.Add(Me.btnAnsattPålogg)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "pålogging"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Logg på"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnAnsattPålogg As Button
    Friend WithEvents txtAnsattBrNavn As TextBox
    Friend WithEvents txtAnsattPassord As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
End Class
