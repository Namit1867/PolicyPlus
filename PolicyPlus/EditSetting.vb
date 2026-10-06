Public Class EditSetting
    Dim CurrentSetting As PolicyPlusPolicy
    Dim CurrentSection As AdmxPolicySection
    Dim AdmxWorkspace As AdmxBundle
    Dim CompPolSource, UserPolSource As IPolicySource
    Dim CompPolLoader, UserPolLoader As PolicyLoader
    Dim CompComments, UserComments As Dictionary(Of String, String)
    ' Above: passed in; below: internal state
    Dim ElementControls As Dictionary(Of String, Control)
    Dim ResizableControls As List(Of Control)
    Dim CurrentSource As IPolicySource
    Dim CurrentLoader As PolicyLoader
    Dim CurrentComments As Dictionary(Of String, String)
    Dim ChangesMade As Boolean ' To either side
    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        If ChangesMade Then DialogResult = DialogResult.OK Else DialogResult = DialogResult.Cancel
    End Sub
    Private Sub EditSetting_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        SettingNameLabel.Text = CurrentSetting.DisplayName
        If CurrentSetting.SupportedOn Is Nothing Then SupportedTextbox.Text = "" Else SupportedTextbox.Text = CurrentSetting.SupportedOn.DisplayName
        HelpTextbox.Text = Main.PrettifyDescription(CurrentSetting.DisplayExplanation)
        If CurrentSetting.RawPolicy.Section = AdmxPolicySection.Both Then
            SectionDropdown.Enabled = True
            CurrentSection = If(CurrentSection = AdmxPolicySection.Both, AdmxPolicySection.Machine, CurrentSection)
        Else
            SectionDropdown.Enabled = False
            CurrentSection = CurrentSetting.RawPolicy.Section
        End If
        ExtraOptionsPanel.HorizontalScroll.Maximum = 0
        ExtraOptionsPanel.VerticalScroll.Visible = True
        ExtraOptionsPanel.AutoScroll = True
        PreparePolicyElements()
        SectionDropdown.Text = If(CurrentSection = AdmxPolicySection.Machine, "Computer", "User")
        SectionDropdown_SelectedIndexChanged(Nothing, Nothing) ' Force an update of the current source
        PreparePolicyState()
        StateRadiosChanged(Nothing, Nothing)
    End Sub
    Sub PreparePolicyElements()
        For n = ExtraOptionsTable.RowCount - 1 To 0 Step -1 ' Go backwards because Dispose changes the indexes
            Dim ctl = ExtraOptionsTable.GetControlFromPosition(0, n)
            If ctl IsNot Nothing Then ctl.Dispose()
        Next
        ExtraOptionsTable.Controls.Clear()
        ExtraOptionsTable.RowCount = 0
        Dim curTabIndex = 10
        Dim addControl = Sub(ID As String, Control As Control, Label As String)
                             ExtraOptionsTable.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                             If Label = "" Then ' Just a single control
                                 If Control.AutoSize Then ResizableControls.Add(Control)
                                 ExtraOptionsTable.Controls.Add(Control, 0, ExtraOptionsTable.RowStyles.Count - 1)
                             Else ' Has a label attached
                                 Dim flowPanel As New FlowLayoutPanel With {.WrapContents = True, .AutoSize = True, .AutoSizeMode = AutoSizeMode.GrowAndShrink}
                                 flowPanel.Margin = New Padding(0)
                                 ExtraOptionsTable.Controls.Add(flowPanel, 0, ExtraOptionsTable.RowStyles.Count - 1)
                                 Dim labelControl As New Label With {.AutoSize = True, .Text = Label}
                                 labelControl.Anchor = AnchorStyles.Left
                                 Control.Anchor = AnchorStyles.Left
                                 flowPanel.Controls.Add(labelControl)
                                 flowPanel.Controls.Add(Control)
                                 ResizableControls.Add(flowPanel)
                             End If
                             If ID <> "" Then
                                 ElementControls.Add(ID, Control)
                                 Control.TabStop = True
                                 Control.TabIndex = curTabIndex
                                 curTabIndex += 1
                             Else
                                 Control.TabStop = False
                             End If
                         End Sub
        ExtraOptionsTable.RowStyles.Clear()
        ElementControls = New Dictionary(Of String, Control)
        ResizableControls = New List(Of Control)
        ' Create the Windows Forms elements
        If CurrentSetting.RawPolicy.Elements IsNot Nothing And CurrentSetting.Presentation IsNot Nothing Then
            Dim elemDict = CurrentSetting.RawPolicy.Elements.ToDictionary(Function(e) e.ID)
            For Each pres In CurrentSetting.Presentation.Elements
                Select Case pres.ElementType
                    Case "text" ' A plain label
                        Dim textPres As LabelPresentationElement = pres
                        Dim label As New Label With {.Text = textPres.Text, .AutoSize = True}
                        label.Margin = New Padding(3, 6, 3, 6)
                        addControl(textPres.ID, label, "")
                    Case "decimalTextBox" ' Numeric spin box or a plain text box restricted to numbers
                        Dim decimalTextPres As NumericBoxPresentationElement = pres
                        Dim numeric As DecimalPolicyElement = elemDict(pres.ID)
                        Dim newControl As Control
                        If decimalTextPres.HasSpinner Then
                            newControl = New NumericUpDown With {
                                .Minimum = numeric.Minimum,
                                .Maximum = numeric.Maximum,
                                .Increment = decimalTextPres.SpinnerIncrement,
                                .Value = decimalTextPres.DefaultValue
                            }
                        Else
                            Dim text As New TextBox
                            AddHandler text.TextChanged, Sub()
                                                             If Not Integer.TryParse(text.Text, 0) Then text.Text = decimalTextPres.DefaultValue
                                                             Dim curNum As Integer = text.Text
                                                             If curNum > numeric.Maximum Then text.Text = numeric.Maximum
                                                             If curNum < numeric.Minimum Then text.Text = numeric.Minimum
                                                         End Sub
                            AddHandler text.KeyPress, Sub(Sender As Object, EventArgs As KeyPressEventArgs)
                                                          If Not (Char.IsControl(EventArgs.KeyChar) Or Char.IsDigit(EventArgs.KeyChar)) Then EventArgs.Handled = True
                                                      End Sub
                            text.Text = decimalTextPres.DefaultValue
                            newControl = text
                        End If
                        addControl(pres.ID, newControl, decimalTextPres.Label)
                    Case "textBox" ' Simple text box
                        Dim textboxPres As TextBoxPresentationElement = pres
                        Dim text As TextPolicyElement = elemDict(pres.ID)
                        Dim textbox As New TextBox With {
                            .Width = ExtraOptionsTable.Width * 0.75,
                            .Text = textboxPres.DefaultValue,
                            .MaxLength = text.MaxLength
                        }
                        addControl(pres.ID, textbox, textboxPres.Label)
                    Case "checkBox" ' Check box
                        Dim checkPres As CheckBoxPresentationElement = pres
                        Dim checkbox As New CheckBox With {.TextAlign = ContentAlignment.MiddleLeft}
                        checkbox.Text = checkPres.Text
                        checkbox.Width = ExtraOptionsTable.ClientSize.Width
                        Using g = checkbox.CreateGraphics ' Figure out how tall it should be
                            Dim size = g.MeasureString(checkbox.Text, checkbox.Font, checkbox.Width)
                            checkbox.Height = (size.Height + checkbox.Padding.Vertical + checkbox.Margin.Vertical)
                        End Using
                        checkbox.Checked = checkPres.DefaultState
                        addControl(pres.ID, checkbox, "")
                    Case "comboBox" ' Text box with suggestions, not tested because it's not used in any default ADML
                        Dim comboPres As ComboBoxPresentationElement = pres
                        Dim text As TextPolicyElement = elemDict(pres.ID)
                        Dim combobox As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDown}
                        combobox.MaxLength = text.MaxLength
                        combobox.Width = ExtraOptionsTable.Width * 0.75
                        combobox.Text = comboPres.DefaultText
                        combobox.Sorted = Not comboPres.NoSort
                        For Each suggestion In comboPres.Suggestions
                            combobox.Items.Add(suggestion)
                        Next
                        addControl(pres.ID, combobox, comboPres.Label)
                    Case "dropdownList" ' Dropdown list of options
                        Dim dropdownPres As DropDownPresentationElement = pres
                        Dim combobox As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList}
                        combobox.Sorted = Not dropdownPres.NoSort
                        Dim enumElem As EnumPolicyElement = elemDict(pres.ID)
                        Dim itemId As Integer = 0
                        Using g = combobox.CreateGraphics ' Figure out how wide it should be, and add entries
                            Dim maxWidth = combobox.Width
                            For Each entry In enumElem.Items
                                Dim map = New DropdownPresentationMap With {.ID = itemId, .DisplayName = AdmxWorkspace.ResolveString(entry.DisplayCode, CurrentSetting.RawPolicy.DefinedIn)}
                                Dim width = g.MeasureString(map.DisplayName, combobox.Font).Width + 25 ' A little extra margin
                                If width > maxWidth Then maxWidth = width
                                combobox.Items.Add(map)
                                If itemId = dropdownPres.DefaultItemID.GetValueOrDefault(-1) Then combobox.SelectedItem = map
                                itemId += 1
                            Next
                            combobox.Width = maxWidth
                        End Using
                        addControl(pres.ID, combobox, dropdownPres.Label)
                    Case "listBox" ' Button to launch a grid view editor
                        Dim listPres As ListPresentationElement = pres
                        Dim list As ListPolicyElement = elemDict(pres.ID)
                        Dim button As New Button With {
                            .UseVisualStyleBackColor = True,
                            .Text = "Edit..."
                        }
                        AddHandler button.Click, Sub()
                                                     If ListEditor.PresentDialog(listPres.Label, button.Tag, list.UserProvidesNames) = DialogResult.OK Then button.Tag = ListEditor.FinalData
                                                 End Sub
                        addControl(pres.ID, button, listPres.Label)
                    Case "multiTextBox" ' Multiline text box
                        Dim multiTextPres As MultiTextPresentationElement = pres
                        Dim bigTextbox As New TextBox With {
                            .AutoSize = False,
                            .Width = ExtraOptionsPanel.Width * 0.8,
                            .Multiline = True,
                            .ScrollBars = ScrollBars.Both,
                            .WordWrap = False,
                            .AcceptsReturn = True
                        }
                        bigTextbox.Height *= 4
                        addControl(pres.ID, bigTextbox, multiTextPres.Label)
                End Select
            Next
            OptionsTableResized()
        End If
    End Sub
    Sub PreparePolicyState()
        ' Set the value of the UI elements depending on the current policy state
        Select Case PolicyProcessing.GetPolicyState(CurrentSource, CurrentSetting)
            Case PolicyState.Disabled
                DisabledOption.Checked = True
            Case PolicyState.Enabled
                EnabledOption.Checked = True
                Dim optionStates = PolicyProcessing.GetPolicyOptionStates(CurrentSource, CurrentSetting)
                For Each kv In optionStates
                    Dim uiControl As Control = ElementControls(kv.Key)
                    If TypeOf kv.Value Is UInteger Then ' Numeric box
                        If TypeOf uiControl Is TextBox Then
                            CType(uiControl, TextBox).Text = kv.Value.ToString
                        Else
                            CType(uiControl, NumericUpDown).Value = kv.Value
                        End If
                    ElseIf TypeOf kv.Value Is String Then ' Text box or combo box
                        If TypeOf uiControl Is ComboBox Then
                            CType(uiControl, ComboBox).Text = kv.Value
                        Else
                            CType(uiControl, TextBox).Text = kv.Value
                        End If
                    ElseIf TypeOf kv.Value Is Integer Then ' Dropdown list
                        Dim combobox As ComboBox = uiControl
                        Dim matchingItem = combobox.Items.OfType(Of DropdownPresentationMap).FirstOrDefault(Function(i) i.ID = CInt(kv.Value))
                        If matchingItem IsNot Nothing Then combobox.SelectedItem = matchingItem
                    ElseIf TypeOf kv.Value Is Boolean Then ' Check box
                        CType(uiControl, CheckBox).Checked = kv.Value
                    ElseIf TypeOf kv.Value Is String() Then ' Multiline text box
                        CType(uiControl, TextBox).Lines = kv.Value
                    Else ' List box (pop-out button)
                        uiControl.Tag = kv.Value
                    End If
                Next
            Case Else
                NotConfiguredOption.Checked = True
        End Select
        Dim canWrite = (CurrentLoader.GetWritability <> PolicySourceWritability.NoWriting)
        ApplyButton.Enabled = canWrite
        OkButton.Enabled = canWrite
        If CurrentComments Is Nothing Then
            CommentTextbox.Enabled = False
            CommentTextbox.Text = "Comments unavailable for this policy source"
        ElseIf CurrentComments.ContainsKey(CurrentSetting.UniqueID) Then
            CommentTextbox.Enabled = True
            CommentTextbox.Text = CurrentComments(CurrentSetting.UniqueID)
        Else
            CommentTextbox.Enabled = True
            CommentTextbox.Text = ""
        End If
    End Sub
    Sub ApplyToPolicySource()
        ' Write the new state to the policy source object
        PolicyProcessing.ForgetPolicy(CurrentSource, CurrentSetting)
        If EnabledOption.Checked Then
            Dim options As New Dictionary(Of String, Object)
            If CurrentSetting.RawPolicy.Elements IsNot Nothing Then
                For Each elem In CurrentSetting.RawPolicy.Elements
                    Dim uiControl As Control = ElementControls(elem.ID)
                    Select Case elem.ElementType
                        Case "decimal"
                            If TypeOf uiControl Is TextBox Then
                                options.Add(elem.ID, CUInt(CType(uiControl, TextBox).Text))
                            Else
                                options.Add(elem.ID, CUInt(CType(uiControl, NumericUpDown).Value))
                            End If
                        Case "text"
                            If TypeOf uiControl Is ComboBox Then
                                options.Add(elem.ID, CType(uiControl, ComboBox).Text)
                            Else
                                options.Add(elem.ID, CType(uiControl, TextBox).Text)
                            End If
                        Case "boolean"
                            options.Add(elem.ID, CType(uiControl, CheckBox).Checked)
                        Case "enum"
                            options.Add(elem.ID, CType(CType(uiControl, ComboBox).SelectedItem, DropdownPresentationMap).ID)
                        Case "list"
                            options.Add(elem.ID, uiControl.Tag)
                        Case "multiText"
                            options.Add(elem.ID, CType(uiControl, TextBox).Lines)
                    End Select
                Next
            End If
            PolicyProcessing.SetPolicyState(CurrentSource, CurrentSetting, PolicyState.Enabled, options)
        ElseIf DisabledOption.Checked Then
            PolicyProcessing.SetPolicyState(CurrentSource, CurrentSetting, PolicyState.Disabled, Nothing)
        End If
        ' Update the comment for this policy
        If CurrentComments IsNot Nothing Then
            If CommentTextbox.Text = "" Then
                If CurrentComments.ContainsKey(CurrentSetting.UniqueID) Then CurrentComments.Remove(CurrentSetting.UniqueID)
            Else
                If CurrentComments.ContainsKey(CurrentSetting.UniqueID) Then CurrentComments(CurrentSetting.UniqueID) = CommentTextbox.Text Else CurrentComments.Add(CurrentSetting.UniqueID, CommentTextbox.Text)
            End If
        End If
    End Sub
    Public Function PresentDialog(Policy As PolicyPlusPolicy, Section As AdmxPolicySection, Workspace As AdmxBundle, CompPolSource As IPolicySource, UserPolSource As IPolicySource, CompPolLoader As PolicyLoader, UserPolLoader As PolicyLoader, CompComments As Dictionary(Of String, String), UserComments As Dictionary(Of String, String)) As DialogResult
        CurrentSetting = Policy
        CurrentSection = Section
        AdmxWorkspace = Workspace
        Me.CompPolSource = CompPolSource
        Me.UserPolSource = UserPolSource
        Me.CompPolLoader = CompPolLoader
        Me.UserPolLoader = UserPolLoader
        Me.CompComments = CompComments
        Me.UserComments = UserComments
        ChangesMade = False
        Return ShowDialog()
    End Function
    Private Sub StateRadiosChanged(sender As Object, e As EventArgs) Handles DisabledOption.CheckedChanged, EnabledOption.CheckedChanged, NotConfiguredOption.CheckedChanged
        If ElementControls Is Nothing Then Exit Sub ' A change to the tab order causes a spurious CheckedChanged
        Dim allowOptions = EnabledOption.Checked
        For Each kv In ElementControls
            kv.Value.Enabled = allowOptions
        Next
    End Sub
    Private Sub SectionDropdown_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SectionDropdown.SelectedIndexChanged
        Dim isUser = (SectionDropdown.Text = "User")
        CurrentSource = If(isUser, UserPolSource, CompPolSource)
        CurrentLoader = If(isUser, UserPolLoader, CompPolLoader)
        CurrentComments = If(isUser, UserComments, CompComments)
        PreparePolicyState()
    End Sub
    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles OkButton.Click
        ApplyToPolicySource()
        DialogResult = DialogResult.OK
    End Sub
    Private Sub ApplyButton_Click(sender As Object, e As EventArgs) Handles ApplyButton.Click
        ApplyToPolicySource()
        ChangesMade = True
    End Sub
    Private Sub OptionsTableResized()
        ' Update the width limit on the extra options controls (in case the vertical scrollbar appeared or disappeared)
        If ResizableControls Is Nothing Then Exit Sub
        ExtraOptionsTable.MaximumSize = New Size(ExtraOptionsPanel.ClientSize.Width, 0)
        ExtraOptionsTable.MinimumSize = ExtraOptionsTable.MaximumSize
        For Each ctl In ResizableControls
            ctl.MaximumSize = New Size(ExtraOptionsPanel.ClientSize.Width, 0)
        Next
    End Sub
    Private Sub EditSetting_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' Share the extra width between the two halves of the form
        Dim extraWidth = Width - 654
        ExtraOptionsPanel.Width = 299 + (extraWidth / 2)
        HelpTextbox.Width = 309 + (extraWidth / 2)
        HelpTextbox.Left = ExtraOptionsPanel.Left + ExtraOptionsPanel.Width + 6
        CommentTextbox.Left = HelpTextbox.Left
        CommentTextbox.Width = HelpTextbox.Width
        SupportedTextbox.Left = HelpTextbox.Left
        SupportedTextbox.Width = HelpTextbox.Width
        CommentLabel.Left = CommentTextbox.Left - 57
        SupportedLabel.Left = SupportedTextbox.Left - 77
        OptionsTableResized()
    End Sub
    Private Sub EditSetting_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If ChangesMade Then DialogResult = DialogResult.OK
    End Sub
    Private Class DropdownPresentationMap ' Used for keeping the ID with an option in dropdown boxes
        Public ID As Integer
        Public DisplayName As String
        Public Overrides Function ToString() As String
            Return DisplayName
        End Function
    End Class
End Class


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

    End Sub

    Friend WithEvents NameTextbox As TextBox
    Friend WithEvents IdTextbox As TextBox
    Friend WithEvents DefinedTextbox As TextBox
    Friend WithEvents DisplayCodeTextbox As TextBox
    Friend WithEvents KindTextbox As TextBox
    Friend WithEvents ParentButton As Button
    Friend WithEvents ParentTextbox As TextBox
    Friend WithEvents ChildrenListview As ListView
    Friend WithEvents ChVersion As ColumnHeader
    Friend WithEvents ChName As ColumnHeader
    Friend WithEvents CloseButton As Button
    Friend WithEvents VersionTextbox As TextBox
End Class
