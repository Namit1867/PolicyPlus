Public Class PolicyProcessing
    Public Shared Function GetPolicyState(PolicySource As IPolicySource, Policy As PolicyPlusPolicy) As PolicyState
        ' Determine the basic state of a policy
        Dim enabledEvidence As Decimal = 0
        Dim disabledEvidence As Decimal = 0
        Dim rawpol = Policy.RawPolicy
        Dim checkOneVal = Sub(Value As PolicyRegistryValue, Key As String, ValueName As String, ByRef EvidenceVar As Decimal)
                              If Value Is Nothing Then Exit Sub
                              If ValuePresent(Value, PolicySource, Key, ValueName) Then EvidenceVar += 1
                          End Sub
        Dim checkValList = Sub(ValList As PolicyRegistrySingleList, DefaultKey As String, ByRef EvidenceVar As Decimal)
                               If ValList Is Nothing Then Exit Sub
                               Dim listKey = If(ValList.DefaultRegistryKey = "", DefaultKey, ValList.DefaultRegistryKey)
                               For Each regVal In ValList.AffectedValues
                                   Dim entryKey = If(regVal.RegistryKey = "", listKey, regVal.RegistryKey)
                                   checkOneVal(regVal.Value, entryKey, regVal.RegistryValue, EvidenceVar)
                               Next
                           End Sub
        ' Check the policy's standard Registry values
        If rawpol.RegistryValue <> "" Then
            If rawpol.AffectedValues.OnValue Is Nothing Then
                checkOneVal(New PolicyRegistryValue With {.NumberValue = 1UI, .RegistryType = PolicyRegistryValueType.Numeric}, rawpol.RegistryKey, rawpol.RegistryValue, enabledEvidence)
            Else
                checkOneVal(rawpol.AffectedValues.OnValue, rawpol.RegistryKey, rawpol.RegistryValue, enabledEvidence)
            End If
            If rawpol.AffectedValues.OffValue Is Nothing Then
                checkOneVal(New PolicyRegistryValue With {.RegistryType = PolicyRegistryValueType.Delete}, rawpol.RegistryKey, rawpol.RegistryValue, disabledEvidence)
            Else
                checkOneVal(rawpol.AffectedValues.OffValue, rawpol.RegistryKey, rawpol.RegistryValue, disabledEvidence)
            End If
        End If
        checkValList(rawpol.AffectedValues.OnValueList, rawpol.RegistryKey, enabledEvidence)
        checkValList(rawpol.AffectedValues.OffValueList, rawpol.RegistryKey, disabledEvidence)
        ' Check the policy's elements
        If rawpol.Elements IsNot Nothing Then
            Dim deletedElements As Decimal = 0
            Dim presentElements As Decimal = 0
            For Each elem In rawpol.Elements
                Dim elemKey = If(elem.RegistryKey = "", rawpol.RegistryKey, elem.RegistryKey)
                If elem.ElementType = "list" Then
                    Dim neededValues = 0
                    If PolicySource.WillDeleteValue(elemKey, "") Then
                        deletedElements += 1
                        neededValues = 1
                    End If
                    If PolicySource.GetValueNames(elemKey).Count > 0 Then
                        deletedElements -= neededValues
                        presentElements += 1
                    End If
                ElseIf elem.ElementType = "boolean" Then
                    Dim booleanElem As BooleanPolicyElement = elem
                    If PolicySource.WillDeleteValue(elemKey, elem.RegistryValue) Then
                        deletedElements += 1 ' Implicit checkboxes are deleted when the policy is disabled
                    Else
                        Dim checkboxDisabled As Decimal = 0
                        checkOneVal(booleanElem.AffectedRegistry.OffValue, elemKey, elem.RegistryValue, checkboxDisabled)
                        checkValList(booleanElem.AffectedRegistry.OffValueList, elemKey, checkboxDisabled)
                        deletedElements += checkboxDisabled * 0.1D ' Checkboxes in the off state are weak evidence for the policy being disabled
                        checkOneVal(booleanElem.AffectedRegistry.OnValue, elemKey, elem.RegistryValue, presentElements)
                        checkValList(booleanElem.AffectedRegistry.OnValueList, elemKey, presentElements)
                    End If
                Else
                    If PolicySource.WillDeleteValue(elemKey, elem.RegistryValue) Then
                        deletedElements += 1
                    ElseIf PolicySource.ContainsValue(elemKey, elem.RegistryValue) Then
                        presentElements += 1
                    End If
                End If
            Next
            If presentElements > 0 Then
                enabledEvidence += presentElements
            ElseIf deletedElements > 0 Then
                disabledEvidence += deletedElements
            End If
        End If
        ' Judge the evidence collected
        If enabledEvidence > disabledEvidence Then
            Return PolicyState.Enabled
        ElseIf disabledEvidence > enabledEvidence Then
            Return PolicyState.Disabled
        ElseIf enabledEvidence = 0 Then ' No evidence for either side
            Return PolicyState.NotConfigured
        Else
            Return PolicyState.Unknown
        End If
    End Function
    Private Shared Function ValuePresent(Value As PolicyRegistryValue, Source As IPolicySource, Key As String, ValueName As String) As Boolean
        ' Determine whether the given value is found in the Registry
        Select Case Value.RegistryType
            Case PolicyRegistryValueType.Delete
                Return Source.WillDeleteValue(Key, ValueName)
            Case PolicyRegistryValueType.Numeric
                If Not Source.ContainsValue(Key, ValueName) Then Return False
                Dim sourceVal = Source.GetValue(Key, ValueName)
                If TypeOf sourceVal IsNot UInteger And TypeOf sourceVal IsNot Integer Then Return False
                Return CLng(sourceVal) = Value.NumberValue
            Case PolicyRegistryValueType.Text
                If Not Source.ContainsValue(Key, ValueName) Then Return False
                Dim sourceVal = Source.GetValue(Key, ValueName)
                If TypeOf sourceVal IsNot String Then Return False
                Return CStr(sourceVal) = Value.StringValue
            Case Else
                Throw New InvalidOperationException("Illegal value type")
        End Select
    End Function
    Private Shared Function ValueListPresent(ValueList As PolicyRegistrySingleList, Source As IPolicySource, Key As String, ValueName As String) As Boolean
        ' Determine whether all the values in a value list are in the Registry data
        Dim sublistKey = If(ValueList.DefaultRegistryKey = "", Key, ValueList.DefaultRegistryKey)
        Return ValueList.AffectedValues.All(Function(e)
                                                Dim entryKey = If(e.RegistryKey = "", sublistKey, e.RegistryKey)
                                                Return ValuePresent(e.Value, Source, entryKey, e.RegistryValue)
                                            End Function)
    End Function
    Public Shared Function DeduplicatePolicies(Workspace As AdmxBundle) As Integer
        ' Merge otherwise-identical pairs of user and computer policies into both-section policies
        Dim dedupeCount = 0
        For Each cat In Workspace.Policies.GroupBy(Function(c) c.Value.Category)
            For Each namegroup In cat.GroupBy(Function(p) p.Value.DisplayName).Select(Function(x) x.ToList).ToList
                If namegroup.Count <> 2 Then Continue For
                Dim a = namegroup(0).Value
                Dim b = namegroup(1).Value
                If a.RawPolicy.Section + b.RawPolicy.Section <> AdmxPolicySection.Both Then Continue For
                If a.DisplayExplanation <> b.DisplayExplanation Then Continue For
                If a.RawPolicy.RegistryKey <> b.RawPolicy.RegistryKey Then Continue For
                a.Category.Policies.Remove(a)
                Workspace.Policies.Remove(a.UniqueID)
                b.RawPolicy.Section = AdmxPolicySection.Both
                dedupeCount += 1
            Next
        Next
        Return dedupeCount
    End Function
    Public Shared Function GetPolicyOptionStates(PolicySource As IPolicySource, Policy As PolicyPlusPolicy) As Dictionary(Of String, Object)
        ' Get the element states of an enabled policy
        Dim state As New Dictionary(Of String, Object)
        If Policy.RawPolicy.Elements Is Nothing Then Return state
        For Each elem In Policy.RawPolicy.Elements
            Dim elemKey = If(elem.RegistryKey = "", Policy.RawPolicy.RegistryKey, elem.RegistryKey)
            Select Case elem.ElementType
                Case "decimal"
                    state.Add(elem.ID, CUInt(PolicySource.GetValue(elemKey, elem.RegistryValue)))
                Case "boolean"
                    Dim booleanElem As BooleanPolicyElement = elem
                    state.Add(elem.ID, GetRegistryListState(PolicySource, booleanElem.AffectedRegistry, elemKey, elem.RegistryValue))
                Case "text"
                    state.Add(elem.ID, PolicySource.GetValue(elemKey, elem.RegistryValue))
                Case "list"
                    Dim listElem As ListPolicyElement = elem
                    If listElem.UserProvidesNames Then ' Keys matter, use a dictionary
                        Dim entries As New Dictionary(Of String, String)
                        For Each value In PolicySource.GetValueNames(elemKey)
                            entries.Add(value, PolicySource.GetValue(elemKey, value))
                        Next
                        state.Add(elem.ID, entries)
                    Else ' Keys don't matter, use a list
                        Dim entries As New List(Of String)
                        If listElem.HasPrefix Then
                            Dim n As Integer = 1
                            Do While PolicySource.ContainsValue(elemKey, elem.RegistryValue & n)
                                entries.Add(PolicySource.GetValue(elemKey, elem.RegistryValue & n))
                                n += 1
                            Loop
                        Else
                            For Each value In PolicySource.GetValueNames(elemKey)
                                entries.Add(value)
                            Next
                        End If
                        state.Add(elem.ID, entries)
                    End If
                Case "enum"
                    ' Determine which option has results that match the Registry
                    Dim enumElem As EnumPolicyElement = elem
                    Dim selectedIndex As Integer = -1
                    For n = 0 To enumElem.Items.Count - 1
                        Dim enumItem = enumElem.Items(n)
                        If ValuePresent(enumItem.Value, PolicySource, elemKey, elem.RegistryValue) Then
                            If enumItem.ValueList Is Nothing OrElse ValueListPresent(enumItem.ValueList, PolicySource, elemKey, elem.RegistryValue) Then
                                selectedIndex = n
                                Exit For
                            End If
                        End If
                    Next
                    state.Add(elem.ID, selectedIndex)
                Case "multiText"
                    state.Add(elem.ID, PolicySource.GetValue(elemKey, elem.RegistryValue))
            End Select
        Next
        Return state
    End Function
    Private Shared Function GetRegistryListState(PolicySource As IPolicySource, RegList As PolicyRegistryList, DefaultKey As String, DefaultValueName As String) As Boolean
        ' Whether a list of Registry values is present
        Dim isListAllPresent = Function(l As PolicyRegistrySingleList) ValueListPresent(l, PolicySource, DefaultKey, DefaultValueName)
        If RegList.OnValue IsNot Nothing Then
            If ValuePresent(RegList.OnValue, PolicySource, DefaultKey, DefaultValueName) Then Return True
        ElseIf RegList.OnValueList IsNot Nothing Then
            If isListAllPresent(RegList.OnValueList) Then Return True
        Else
            If CUInt(PolicySource.GetValue(DefaultKey, DefaultValueName)) = 1UI Then Return True
        End If
        If RegList.OffValue IsNot Nothing Then
            If ValuePresent(RegList.OffValue, PolicySource, DefaultKey, DefaultValueName) Then Return False
        ElseIf RegList.OffValueList IsNot Nothing Then
            If isListAllPresent(RegList.OffValueList) Then Return False
        End If
        Return False
    End Function
    Public Shared Function GetReferencedRegistryValues(Policy As PolicyPlusPolicy) As List(Of RegistryKeyValuePair)
        Return WalkPolicyRegistry(Nothing, Policy, False)
    End Function
    Public Shared Sub ForgetPolicy(PolicySource As IPolicySource, Policy As PolicyPlusPolicy)
        WalkPolicyRegistry(PolicySource, Policy, True)
    End Sub
    Private Shared Function WalkPolicyRegistry(PolicySource As IPolicySource, Policy As PolicyPlusPolicy, Forget As Boolean) As List(Of RegistryKeyValuePair)
        ' This function handles both GetReferencedRegistryValues and ForgetPolicy because they require searching through the same things
        Dim entries As New List(Of RegistryKeyValuePair)
        Dim addReg = Sub(Key As String, Value As String)
                         Dim rkvp As New RegistryKeyValuePair With {.Key = Key, .Value = Value}
                         If Not entries.Contains(rkvp) Then entries.Add(rkvp)
                     End Sub
        ' Get all Registry values affected by this policy
        Dim rawpol = Policy.RawPolicy
        If rawpol.RegistryValue <> "" Then addReg(rawpol.RegistryKey, rawpol.RegistryValue)
        Dim addSingleList = Sub(SingleList As PolicyRegistrySingleList, OverrideKey As String)
                                If SingleList Is Nothing Then Exit Sub
                                Dim defaultKey = If(OverrideKey = "", rawpol.RegistryKey, OverrideKey)
                                Dim listKey = If(SingleList.DefaultRegistryKey = "", defaultKey, SingleList.DefaultRegistryKey)
                                For Each e In SingleList.AffectedValues
                                    Dim entryKey = If(e.RegistryKey = "", listKey, e.RegistryKey)
                                    addReg(entryKey, e.RegistryValue)
                                Next
                            End Sub
        addSingleList(rawpol.AffectedValues.OnValueList, "")
        addSingleList(rawpol.AffectedValues.OffValueList, "")
        If rawpol.Elements IsNot Nothing Then
            For Each elem In rawpol.Elements
                Dim elemKey As String = If(elem.RegistryKey = "", rawpol.RegistryKey, elem.RegistryKey)
                If elem.ElementType <> "list" Then addReg(elemKey, elem.RegistryValue)
                Select Case elem.ElementType
                    Case "boolean"
                        Dim booleanElem As BooleanPolicyElement = elem
                        addSingleList(booleanElem.AffectedRegistry.OnValueList, elemKey)
                        addSingleList(booleanElem.AffectedRegistry.OffValueList, elemKey)
                    Case "enum"
                        Dim enumElem As EnumPolicyElement = elem
                        For Each e In enumElem.Items
                            addSingleList(e.ValueList, elemKey)
                        Next
                    Case "list"
                        If Forget Then
                            PolicySource.ClearKey(elemKey) ' Delete all the values
                            PolicySource.ForgetKeyClearance(elemKey)
                        Else
                            addReg(elemKey, "")
                        End If
                End Select
            Next
        End If
        If Forget Then ' Remove them if forgetting
            For Each e In entries
                PolicySource.ForgetValue(e.Key, e.Value)
            Next
        End If
        Return entries
    End Function
    Public Shared Sub SetPolicyState(PolicySource As IPolicySource, Policy As PolicyPlusPolicy, State As PolicyState, Options As Dictionary(Of String, Object))
        ' Write a full policy state to the policy source
        Dim setValue = Sub(Key As String, ValueName As String, Value As PolicyRegistryValue)
                           If Value Is Nothing Then Exit Sub
                           Select Case Value.RegistryType
                               Case PolicyRegistryValueType.Delete
                                   PolicySource.DeleteValue(Key, ValueName)
                               Case PolicyRegistryValueType.Numeric
                                   PolicySource.SetValue(Key, ValueName, Value.NumberValue, Microsoft.Win32.RegistryValueKind.DWord)
                               Case PolicyRegistryValueType.Text
                                   PolicySource.SetValue(Key, ValueName, Value.StringValue, Microsoft.Win32.RegistryValueKind.String)
                           End Select
                       End Sub
        Dim setSingleList = Sub(SingleList As PolicyRegistrySingleList, DefaultKey As String)
                                If SingleList Is Nothing Then Exit Sub
                                Dim listKey As String = If(SingleList.DefaultRegistryKey = "", DefaultKey, SingleList.DefaultRegistryKey)
                                For Each e In SingleList.AffectedValues
                                    Dim itemKey As String = If(e.RegistryKey = "", listKey, e.RegistryKey)
                                    setValue(itemKey, e.RegistryValue, e.Value)
                                Next
                            End Sub
        Dim setList = Sub(List As PolicyRegistryList, DefaultKey As String, DefaultValue As String, IsOn As Boolean)
                          If List Is Nothing Then Exit Sub
                          If IsOn Then
                              setValue(DefaultKey, DefaultValue, List.OnValue)
                              setSingleList(List.OnValueList, DefaultKey)
                          Else
                              setValue(DefaultKey, DefaultValue, List.OffValue)
                              setSingleList(List.OffValueList, DefaultKey)
                          End If