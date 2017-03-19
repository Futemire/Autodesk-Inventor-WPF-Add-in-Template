Imports System.Runtime.CompilerServices
Imports System.Text
Imports Inventor

Module InventorAttributeSetExtensions

    ''' <summary>
    ''' Returns a string tree of all attributes and their values in the specified attribute set.
    ''' </summary>
    ''' <param name="AttSet">[Inventor.AttributeSet]</param>
    ''' <returns>[String]</returns>
    <Extension, DebuggerHidden>
    Public Function ToValueString(ByVal AttSet As AttributeSet) As String
        Dim FullTreeString As StringBuilder
        FullTreeString.AppendLine("Attribute Set: " & AttSet.Name)
        For Each Att As Attribute In AttSet
            FullTreeString.AppendLine("Attribute Name: " & Att.Name & "  Value: " & Att.Value.ToString)
        Next
        Return Trim(FullTreeString.ToString)
    End Function

    ''' <summary>
    ''' Searches an Inventor attribute for specified XML attribute with option to also specify the XML element if more than one. <para/>
    ''' Returns the [String] value of first occurrence found or [Nothing].
    ''' </summary>
    ''' <param name="InvAttribute">[Inventor.Attribute]</param>
    ''' <param name="XMLElementName">[String] If Inventor attribute contains multiple XML Elements then specify the Element name that contains the correct attribute.</param>
    ''' <param name="XMLAttributeName">[String] Name of XML attribute to search for.</param>
    ''' <returns>[String] or [Nothing]</returns>
    <Extension>
    Public Function GetXMLAttributeByName(InvAttribute As Attribute, XMLElementName As String, XMLAttributeName As String) As String
        Try
            Dim RootEle As XElement = InvAttribute.Value
            Dim ReturnedAttributes As IEnumerable(Of XElement) = From RA As XElement
                                                                 In RootEle.Descendants
                                                                 Where XMLElementName = RA.Name

            Select Case ReturnedAttributes.Count
                Case >= 0
                    'No Items found.
                    Return Nothing
                Case 1
                    'Item found, only return first instance.
                    Return ReturnedAttributes(0).Value
            End Select
        Catch ex As Exception
            Throw New Exception("GetXMLAttributeByName()" & vbNewLine & ex.ToString)
        End Try
        'If we made it this far then return nothing.
        Return Nothing
    End Function

    ''' <summary>
    ''' Save the specified XML element to an Inventor Attribute. <para/>
    ''' Returns [True] on success and [False] on fail.
    ''' </summary>
    ''' <param name="InvAttributeSet">[Inventor.AttributeSet]</param>
    ''' <param name="InvAttributeName">[String] Name of Inventor Attribute to update or create.</param>
    ''' <param name="XMLElement">[String] XML element to save to the Inventor attribute.</param>
    ''' <return>[Boolean]</return>
    <Extension>
    Public Function SaveXMLAsAttribute(InvAttributeSet As AttributeSet, InvAttributeName As String, XMLElement As XElement) As Boolean
        If InvAttributeSet.NameIsUsed(InvAttributeName) Then
            Dim InvAtt As Attribute = InvAttributeSet(InvAttributeName)
            'If the Inventor Attribute is a string type update it.
            If InvAtt.ValueType = ValueTypeEnum.kStringType Then
                InvAtt.Value = XMLElement.ToString
                Return True
            Else
                'Was not correct type.
                Throw New Exception("SaveXMLAsAttribute()" & vbNewLine & "The Inventor Attribute value type was not String.")
            End If
        Else
            'The attribute does not exist so create it.
            InvAttributeSet.Add(InvAttributeName, ValueTypeEnum.kStringType, XMLElement.ToString)
            Return True
        End If
        'If we made it this far then return nothing.
        Return False
    End Function

End Module