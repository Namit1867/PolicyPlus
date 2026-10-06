' Raw data loaded from ADMX files
Public Class AdmxProduct
    Public ID As String
    Public DisplayCode As String
    Public Type As AdmxProductType
    Public Version As Integer
    Public Parent As AdmxProduct
    Public DefinedIn As AdmxFile
End Class
Public Enum AdmxProductType
    Product
    MajorRevision
    MinorRevision
End Enum
Public Class AdmxSupportDefinition
    Public ID As String
    Public DisplayCode As String
    Public Logic As AdmxSupportLogicType
    Public Entries As List(Of AdmxSupportEntry)
    Public DefinedIn As AdmxFile
End Class
Public Enum AdmxSupportLogicType
    Blank
    AllOf
    AnyOf
End Enum
Public Class AdmxSupportEntry
    Public ProductID As String
    Public IsRange As Boolean
    Public MinVersion As Integer?
    Public MaxVersion As Integer?
End Class
Public Class AdmxCategory
    Public ID As String
    Public DisplayCode As String
    Public ExplainCode As String
    Public ParentID As String
    Public DefinedIn As AdmxFile
End Class
Public Class AdmxPolicy
    Public ID As String
    Public Section As AdmxPolicySection
    Public CategoryID As String
    Public DisplayCode As String
    Public ExplainCode As String
    Public SupportedCode As String
    Public PresentationID As String
    Public ClientExtension As String
    Public RegistryKey As String
    Public RegistryValue As String
    Public AffectedValues As PolicyRegistryList
    Public Elements As List(Of PolicyElement)
    Public DefinedIn As AdmxFile
End Class
Public Enum AdmxPolicySection
    Machine = 1
    User = 2
    Both = 3
End Enum


    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim NameLabel As System.Windows.Forms.Label
        Dim IdLabel As System.Windows.Forms.Label
        Dim DefinedLabel As System.Windows.Forms.Label
        Dim KindLabel As System.Windows.Forms.Label
        Dim VersionLabel As System.Windows.Forms.Label
        Dim ParentLabel As System.Windows.Forms.Label
        Dim ChildrenLabel As System.Windows.Forms.Label
        Dim DisplayCodeLabel As System.Windows.Forms.Label
        Me.NameTextbox = New System.Windows.Forms.TextBox()
        Me.IdTextbox = New System.Windows.Forms.TextBox()
        Me.DefinedTextbox = New System.Windows.Forms.TextBox()
        Me.DisplayCodeTextbox = New System.Windows.Forms.TextBox()
        Me.KindTextbox = New System.Windows.Forms.TextBox()
        Me.ParentButton = New System.Windows.Forms.Button()
        Me.ParentTextbox = New System.Windows.Forms.TextBox()
        Me.ChildrenListview = New System.Windows.Forms.ListView()
        Me.ChVersion = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ChName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CloseButton = New System.Windows.Forms.Button()
        Me.VersionTextbox = New System.Windows.Forms.TextBox()
        NameLabel = New System.Windows.Forms.Label()
        IdLabel = New System.Windows.Forms.Label()
        DefinedLabel = New System.Windows.Forms.Label()
        KindLabel = New System.Windows.Forms.Label()
        VersionLabel = New System.Windows.Forms.Label()
        ParentLabel = New System.Windows.Forms.Label()
        ChildrenLabel = New System.Windows.Forms.Label()
        DisplayCodeLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'NameLabel
        '
        NameLabel.AutoSize = True
        NameLabel.Location = New System.Drawing.Point(12, 15)
        NameLabel.Name = "NameLabel"
        NameLabel.Size = New System.Drawing.Size(35, 13)
        NameLabel.TabIndex = 1
        NameLabel.Text = "Name"
        '
        'IdLabel
        '
        IdLabel.AutoSize = True
        IdLabel.Location = New System.Drawing.Point(12, 41)
        IdLabel.Name = "IdLabel"
        IdLabel.Size = New System.Drawing.Size(55, 13)
        IdLabel.TabIndex = 6
        IdLabel.Text = "Unique ID"
        '
        'DefinedLabel
        '
        DefinedLabel.AutoSize = True
        DefinedLabel.Location = New System.Drawing.Point(12, 67)
        DefinedLabel.Name = "DefinedLabel"
        DefinedLabel.Size = New System.Drawing.Size(55, 13)
        DefinedLabel.TabIndex = 7
        DefinedLabel.Text = "Defined in"
        '
        'KindLabel
        '
        KindLabel.AutoSize = True
        KindLabel.Location = New System.Drawing.Point(12, 119)
        KindLabel.Name = "KindLabel"
        KindLabel.Size = New System.Drawing.Size(28, 13)
        KindLabel.TabIndex = 8
        KindLabel.Text = "Kind"
        '
        'VersionLabel
        '
        VersionLabel.AutoSize = True
        VersionLabel.Location = New System.Drawing.Point(12, 145)
        VersionLabel.Name = "VersionLabel"
        VersionLabel.Size = New System.Drawing.Size(80, 13)
        VersionLabel.TabIndex = 9
        VersionLabel.Text = "Version number"
        '
        'ParentLabel
        '
        ParentLabel.AutoSize = True
        ParentLabel.Location = New System.Drawing.Point(12, 171)
        ParentLabel.Name = "ParentLabel"
        ParentLabel.Size = New System.Drawing.Size(38, 13)
        ParentLabel.TabIndex = 12
        ParentLabel.Text = "Parent"
        '
        'ChildrenLabel
        '
        ChildrenLabel.AutoSize = True
        ChildrenLabel.ForeColor = System.Drawing.SystemColors.ControlText
        ChildrenLabel.Location = New System.Drawing.Point(12, 197)
        ChildrenLabel.Name = "ChildrenLabel"
        ChildrenLabel.Size = New System.Drawing.Size(67, 13)
        ChildrenLabel.TabIndex = 14
        ChildrenLabel.Text = "Subproducts"
        '
        'DisplayCodeLabel
        '
        DisplayCodeLabel.AutoSize = True
        DisplayCodeLabel.Location = New System.Drawing.Point(12, 93)
        DisplayCodeLabel.Name = "DisplayCodeLabel"
        DisplayCodeLabel.Size = New System.Drawing.Size(68, 13)
        DisplayCodeLabel.TabIndex = 16
        DisplayCodeLabel.Text = "Display code"
        '
        'NameTextbox
        '
        Me.NameTextbox.Location = New System.Drawing.Point(98, 12)
        Me.NameTextbox.Name = "NameTextbox"
        Me.NameTextbox.ReadOnly = True
        Me.NameTextbox.Size = New System.Drawing.Size(256, 20)
        Me.NameTextbox.TabIndex = 0
        '
        'IdTextbox
        '
        Me.IdTextbox.Location = New System.Drawing.Point(98, 38)
        Me.IdTextbox.Name = "IdTextbox"
        Me.IdTextbox.ReadOnly = True
        Me.IdTextbox.Size = New System.Drawing.Size(256, 20)
        Me.IdTextbox.TabIndex = 2
        '
        'DefinedTextbox
        '
        Me.DefinedTextbox.Location = New System.Drawing.Point(98, 64)
        Me.DefinedTextbox.Name = "DefinedTextbox"
        Me.DefinedTextbox.ReadOnly = True
        Me.DefinedTextbox.Size = New System.Drawing.Size(256, 20)
        Me.DefinedTextbox.TabIndex = 3
        '
        'DisplayCodeTextbox
        '
        Me.DisplayCodeTextbox.Location = New System.Drawing.Point(98, 90)
        Me.DisplayCodeTextbox.Name = "DisplayCodeTextbox"
        Me.DisplayCodeTextbox.ReadOnly = True
        Me.DisplayCodeTextbox.Size = New System.Drawing.Size(256, 20)
        Me.DisplayCodeTextbox.TabIndex = 4
        '
        'KindTextbox
        '
        Me.KindTextbox.Location = New System.Drawing.Point(98, 116)
        Me.KindTextbox.Name = "KindTextbox"
        Me.KindTextbox.ReadOnly = True
        Me.KindTextbox.Size = New System.Drawing.Size(256, 20)
        Me.KindTextbox.TabIndex = 5
        '
        'ParentButton
        '
        Me.ParentButton.Location = New System.Drawing.Point(279, 166)
        Me.ParentButton.Name = "ParentButton"
        Me.ParentButton.Size = New System.Drawing.Size(75, 23)
        Me.ParentButton.TabIndex = 10
        Me.ParentButton.Text = "Details"
        Me.ParentButton.UseVisualStyleBackColor = True
        '
        'ParentTextbox
        '
        Me.ParentTextbox.Location = New System.Drawing.Point(98, 168)
        Me.ParentTextbox.Name = "ParentTextbox"
        Me.ParentTextbox.ReadOnly = True
        Me.ParentTextbox.Size = New System.Drawing.Size(175, 20)
        Me.ParentTextbox.TabIndex = 7
        '
        'ChildrenListview
        '
        Me.ChildrenListview.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ChVersion, Me.ChName})
        Me.ChildrenListview.FullRowSelect = True
        Me.ChildrenListview.HideSelection = False
        Me.ChildrenListview.Location = New System.Drawing.Point(98, 194)
        Me.ChildrenListview.MultiSelect = False
        Me.ChildrenListview.Name = "ChildrenListview"
        Me.ChildrenListview.Size = New System.Drawing.Size(256, 110)
        Me.ChildrenListview.TabIndex = 13
        Me.ChildrenListview.UseCompatibleStateImageBehavior = False
        Me.ChildrenListview.View = System.Windows.Forms.View.Details
        '
        'ChVersion
        '
        Me.ChVersion.Text = "Version"
        Me.ChVersion.Width = 51
        '
        'ChName
        '
        Me.ChName.Text = "Name"
        Me.ChName.Width = 176
        '
        'CloseButton
        '
        Me.CloseButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.CloseButton.Location = New System.Drawing.Point(279, 310)
        Me.CloseButton.Name = "CloseButton"
        Me.CloseButton.Size = New System.Drawing.Size(75, 23)
        Me.CloseButton.TabIndex = 15
        Me.CloseButton.Text = "Close"
        Me.CloseButton.UseVisualStyleBackColor = True
        '
        'VersionTextbox
        '
        Me.VersionTextbox.Location = New System.Drawing.Point(98, 142)
        Me.VersionTextbox.Name = "VersionTextbox"
        Me.VersionTextbox.ReadOnly = True
        Me.VersionTextbox.Size = New System.Drawing.Size(256, 20)
        Me.VersionTextbox.TabIndex = 6
        '
        'DetailProduct
        '
        Me.AcceptButton = Me.CloseButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.CloseButton
        Me.ClientSize = New System.Drawing.Size(366, 345)
        Me.Controls.Add(Me.VersionTextbox)
        Me.Controls.Add(DisplayCodeLabel)
        Me.Controls.Add(Me.CloseButton)
        Me.Controls.Add(ChildrenLabel)
        Me.Controls.Add(Me.ChildrenListview)
        Me.Controls.Add(ParentLabel)
        Me.Controls.Add(Me.ParentTextbox)
        Me.Controls.Add(Me.ParentButton)
        Me.Controls.Add(VersionLabel)
        Me.Controls.Add(KindLabel)
        Me.Controls.Add(DefinedLabel)
        Me.Controls.Add(IdLabel)
        Me.Controls.Add(Me.KindTextbox)
        Me.Controls.Add(Me.DisplayCodeTextbox)
        Me.Controls.Add(Me.DefinedTextbox)
        Me.Controls.Add(Me.IdTextbox)
        Me.Controls.Add(NameLabel)
        Me.Controls.Add(Me.NameTextbox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DetailProduct"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Product Details"
        Me.ResumeLayout(False)
        Me.PerformLayout()