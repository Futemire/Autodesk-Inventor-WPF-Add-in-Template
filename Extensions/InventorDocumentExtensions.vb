Imports System.Runtime.CompilerServices
Imports Inventor

Module InventorDocumentExtensions
    Public Enum PropertySetEnum
        All
        SummaryInfo
        DesignTracking
        Custom
    End Enum

    ''' <summary>
    ''' Inventor Document InventorDocumentSubType
    ''' </summary>
    <Flags>
    Public Enum InventorDocumentSubType
        UnKnown
        SolidPart
        SheetMetal
        GenericProxy
        CompatabilityProxy
        CatalogProxy
        Assembly
        Drawing
        DesignElement
        Presentation
    End Enum

    ''' <summary>
    ''' Gets the Document SubType of Inventor Document.
    ''' </summary>
    ''' <param name="InvDoc">[Inventor.Document] Document to return the SubType on.</param>
    ''' <returns>[InventorDocumentSubType] Document SubType Enum</returns>
    ''' <remarks></remarks>
    <Extension, DebuggerHidden>
    Public Function GetDocumentSubType(ByRef InvDoc As Inventor.Document) As InventorDocumentSubType
        Select Case InvDoc.DocumentSubType.DocumentSubTypeID
            Case "{4D29B490-49B2-11D0-93C3-7E0706000000}"
                Return InventorDocumentSubType.SolidPart
            Case "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
                Return InventorDocumentSubType.SheetMetal
            Case "{92055419-B3FA-11D3-A479-00C04F6B9531}"
                Return InventorDocumentSubType.GenericProxy
            Case "{9C464204-9BAE-11D3-8BAD-0060B0CE6BB4}"
                Return InventorDocumentSubType.CompatabilityProxy
            Case "{9C88D3AF-C3EB-11D3-B79E-0060B0F159EF}"
                Return InventorDocumentSubType.CatalogProxy
            Case "{E60F81E1-49B3-11D0-93C3-7E0706000000}"
                Return InventorDocumentSubType.Assembly
            Case "{BBF9FDF1-52DC-11D0-8C04-0800090BE8EC}"
                Return InventorDocumentSubType.Drawing
            Case "{62FBB030-24C7-11D3-B78D-0060B0F159EF}"
                Return InventorDocumentSubType.DesignElement
            Case "{76283A80-50DD-11D3-A7E3-00C04F79D7BC}"
                Return InventorDocumentSubType.Presentation
        End Select
        Return InventorDocumentSubType.UnKnown
    End Function

    ''' <summary>
    ''' Function returns [Parameter] with specified name Or [Nothing].
    ''' </summary>
    ''' <param name="ParameterName">[String] Name of parameter to return.</param>
    ''' <returns>[Inventor.Parameter] Or [Nothing]</returns>
    <Extension>
    Public Function GetParameter(InvDoc As Inventor.PartDocument, ParameterName As String) As Inventor.Parameter
        Try
            Return InvDoc.ComponentDefinition.Parameters.Item(ParameterName)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Function returns [Parameter] with specified name Or [Nothing].
    ''' </summary>
    ''' <param name="ParameterName">[String] Name of parameter to return.</param>
    ''' <returns>[Inventor.Parameter] Or [Nothing]</returns>
    <Extension>
    Public Function GetParameter(InvDoc As Inventor.AssemblyDocument, ParameterName As String) As Inventor.Parameter
        Try
            Return InvDoc.ComponentDefinition.Parameters.Item(ParameterName)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns the specified iProperty from the specified PropertySet or Nothing.
    ''' </summary>
    ''' <param name="InventorDocument">[Inventor.Document]</param>
    ''' <param name="PropertyName">[String] Name of property to search for.</param>
    ''' <param name="PropertySet">[PropertySetEnum] PropertySet to search. Default=All</param>
    ''' <param name="CreateProperty">[Boolean] If TRUE then create non existing property.</param>
    ''' <returns></returns>
    <Extension>
    Public Function GetIpropertyFromPropertySet(InventorDocument As Document, PropertyName As String,
                                                Optional PropertySet As PropertySetEnum = PropertySetEnum.All,
                                                Optional CreateProperty As Boolean = False,
                                                Optional PropertyValue As String = "") As [Property]
        If PropertySet = PropertySetEnum.All Then
            'Search all properties.
            For Each PropSet As PropertySet In InventorDocument.PropertySets
                For Each Prop As [Property] In PropSet
                    If Prop.Name = PropertyName Then Return Prop
                Next
            Next
        Else
            'Search a specific PropertySet.
            Dim SetName As String = Nothing
            Select Case PropertySet
                Case PropertySetEnum.SummaryInfo
                    SetName = "Inventor Summary Information"
                Case PropertySetEnum.DesignTracking
                    SetName = "Design Tracking Properties"
                Case PropertySetEnum.Custom
                    SetName = "Inventor User Defined Properties"
            End Select
            'Try to return a value.
            Try
                Return InventorDocument.PropertySets(SetName)(PropertyName)
            Catch ex As Exception

                If CreateProperty Then
                    'Create the non-existing property as a custom iProperty.
                    Return InventorDocument.PropertySets("Inventor User Defined Properties").Add(PropertyValue, PropertyName)
                End If
            End Try
        End If
        'Was not found and was not created.
        Return Nothing
    End Function

End Module