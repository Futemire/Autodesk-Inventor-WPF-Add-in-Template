Imports Inventor
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' The ControlDefBaseClass handles the repetitive work of creating/accessing Tabs and Panels as well as adding Control Definitions to specified Panels.  <para/>
''' This class must be inherited.
''' </summary>
Public MustInherit Class ControlDefBaseClass : Inherits AxHost

    ''' <summary>
    ''' Inventor Tool Ribbon Types.
    ''' <para />RibbonType: All • Assembly • Drawing • Part • ZeroDoc
    ''' </summary>
    Public Enum RibbonType
        ''' <summary>
        ''' Adds the control to all Tool Ribbons.
        ''' </summary>
        All
        ''' <summary>
        ''' Add the control to the Assembly Tool Ribbon.
        ''' </summary>
        Assembly
        ''' <summary>
        ''' Add the control to the Drawing Tool Ribbon.
        ''' </summary>
        Drawing
        ''' <summary>
        ''' Add the control to the Solid Part and Sheet Metal Part Tool Ribbons.
        ''' </summary>
        Part
        ''' <summary>
        ''' Add the control to the No Documents Open Tool Ribbon.
        ''' </summary>
        ZeroDoc
    End Enum

    ''' <summary>
    ''' AddinBaseClass object to be used referenced by this class.
    ''' </summary>
    Private _Addin As AddinBaseClass

    ''' <summary>
    ''' Returns a collection of Inventor ControlDefinitions.
    ''' </summary>
    ''' <returns>[Inventor.ControlDefinitions]</returns>
    Public ReadOnly Property ControlDefs() As ControlDefinitions
        Get
            Return _Addin.InvApp.CommandManager.ControlDefinitions
        End Get
    End Property

    ''' <summary>
    ''' Returns the active CommandControl object from the current panel of the active environment.
    ''' </summary>
    ''' <returns>[CommandControl]</returns>
    Public ReadOnly Property ButtonControl As CommandControl
        Get
            Return _Addin.InvApp.UserInterfaceManager.ActiveEnvironment.Ribbon.RibbonTabs.Item(TabInternalName) _
                    .RibbonPanels.Item(PanelInternalName).CommandControls.Item(ButtonInternalName)
        End Get
    End Property

    ''' <summary>
    ''' [Required] The display name of the Button.
    ''' </summary>
    Public Property ButtonName As String
    ''' <summary>
    ''' [Required] A unique identification name for the button.
    ''' <para />
    ''' Best Practice: Use legible name with no spaces. <para/>
    ''' Format example: MyCompany_MyAddin_ButtonName
    ''' </summary>
    Public ButtonInternalName As String
    ''' <summary>
    ''' [Optional] Text to be displayed in tool tip when hovering over the button.
    ''' </summary>
    Public ToolTip As String
    ''' <summary>
    ''' [Optional] Small icon to be displayed on button. Size 16px X 16px
    ''' <para/>
    ''' Accepted Image Types: BMP • JPG • PNG
    ''' <para/>
    ''' Note: ButtonName is always displayed if an icon is not set.
    ''' </summary>
    Public SmallIcon As Bitmap
    ''' <summary>
    ''' [Optional] Large icon to be displayed on button. Size 32px X 32px
    ''' <para />
    ''' Accepted Image Types: BMP • JPG • PNG
    ''' <para />
    ''' Note: ButtonName is always displayed if an icon is not set.
    ''' </summary>
    Public LargeIcon As Bitmap
    ''' <summary>
    ''' [True] Shows the button name and icon. (Default)
    ''' <para />
    ''' [False] Hides the button name and only displays the icon.
    ''' <para />
    ''' Note: ButtonName is always displayed if an icon is not set.
    ''' </summary>
    Public ShowButtonText As Boolean = True
    ''' <summary>
    ''' [True](Default) Displays button using the LargeIcon. 
    ''' <para />
    ''' [False] Displays button using the SmallIcon.
    ''' <para />
    ''' Note: If LargeIcon is not set then the SmallIcon is scaled.
    ''' </summary>
    Public LargeButtonFormat As Boolean = True
    ''' <summary>
    ''' [Required] The display name of the ribbon panel that the button will be added to.
    ''' <para/>
    ''' This name is used only if the ribbon panel needs to be created.
    ''' </summary>
    Public PanelName As String
    ''' <summary>
    ''' [Required] The unique identification name of the panel that the button will be added to.
    ''' <para />
    ''' This will be used as the internal name of the ribbon panel if it is to be created.
    ''' <para />
    ''' Best Practice: Use legible name with no spaces. <para/>
    ''' Format example: MyCompany_MyAddin_PanelName
    ''' </summary>
    Public PanelInternalName As String
    ''' <summary>
    ''' [Required] The display name of the ribbon tab that includes the panel that holds the button.
    ''' <para />
    ''' This name is used if the ribbon tab needs to be created.
    ''' </summary>
    Public TabName As String
    ''' <summary>
    ''' [Required] The unique identification name of the ribbon tab that includes the panel that holds the button.
    ''' <para />
    ''' This will be used as the internal name of the ribbon tab if it is to be created.
    ''' <para />
    ''' Best Practice: Use legible name with no spaces. <para/>
    ''' Format example: MyCompany_MyAddin_TabName
    ''' </summary>
    Public TabInternalName As String
    ''' <summary>
    ''' [True] Displays an expanding tool-tip that can hold a short summary of the command and an image or video.
    ''' <para />
    ''' [False](Default) Displays the standard tool-tip and does not load the progressive tool-tip objects. 
    ''' </summary>
    Public Progressive As Boolean
    ''' <summary>
    ''' [Optional] Title displayed at the top of the progressive tool-tip.
    ''' <para />
    ''' Ignored if Progressive is set to [False].
    ''' </summary>
    Public ProgressiveTitle As String
    ''' <summary>
    ''' Summary or Description text presented in the progressive tool-tip window.
    ''' <para />
    ''' Ignored if Progressive is set to [False].
    ''' </summary>
    Public ProgressiveDesc As String
    ''' <summary>
    ''' An Image presented in the progressive tool-tip window.
    ''' <para/>
    ''' Accepted Image Types: BMP • JPG • PNG
    ''' <para />
    ''' Ignored if Progressive is set to [False].
    ''' </summary>
    Public ProgressiveImage As Image
    ''' <summary>
    ''' An Video presented in the progressive tool-tip window.
    ''' <para/>
    ''' [String] Video stream path or URL.
    ''' <para />
    ''' Ignored if Progressive is set to [False].
    ''' </summary>
    Public ProgressiveVideo As String
    ''' <summary>
    ''' [Optional] Holds a list of RibbonType that define which Tool Ribbons the control will be added too.
    ''' <para />
    ''' Note: If no RibbonTypes are added to the list then the ControlDefinition will be created but not added to the UI.
    ''' </summary>
    Public IncludeInRibbons As List(Of RibbonType)
    ''' <summary>
    ''' The [ButtonDefinition] object for the current [ControlDefinition].
    ''' </summary>
    Public ButtonDefinition As ButtonDefinition
    ''' <summary>
    ''' The [CommandControl] object for the current [ControlDefinition].
    ''' </summary>
    Public ButtonCommand As CommandControl

    ''' <summary>
    ''' Creates a new instance of the AddinControlDefinitionBaseClass.
    ''' </summary>
    ''' <param name="Addin">[AddinBaseClass] AddinBaseClass object to be used referenced by this class.</param>
    Sub New(ByVal Addin As AddinBaseClass)
        'Needed to Initialize the AxHost...
        MyBase.New(CLID)

        'Pass the Addin Object.
        _Addin = Addin
    End Sub

    ''' <summary>
    ''' Creates a [ButtonDefinitionObject] for Autodesk Inventor command buttons.
    ''' </summary>
    ''' <returns>[ButtonDefinitionObject]</returns>
    Public Function CreateControlDefButton() As ButtonDefinitionObject

        'Convert the bitmaps to IPictureDisp
        Dim ICO16 As IPictureDisp = Nothing, ICO32 As IPictureDisp = Nothing

        If Not SmallIcon Is Nothing Then ICO16 = DirectCast(GetIPictureDispFromPicture(SmallIcon), IPictureDisp)
        If Not LargeIcon Is Nothing Then ICO32 = DirectCast(GetIPictureDispFromPicture(LargeIcon), IPictureDisp)

        'Create The Button Definition.
        ButtonDefinition = ControlDefs.AddButtonDefinition(ButtonName, ButtonInternalName, CommandTypesEnum.kQueryOnlyCmdType, CLID, ToolTip, ToolTip, ICO16, ICO32)

        'If we have not set the control to have a progressive tool tip then go ahead and return the ButtonDefinition.
        If Not Progressive Then Return ButtonDefinition

        'Enable the progressive tool tip.
        ButtonDefinition.ProgressiveToolTip.IsProgressive = True

        'If a title was defined then add it.
        If Not ProgressiveTitle Is Nothing Then ButtonDefinition.ProgressiveToolTip.Title = ProgressiveTitle

        'If there is an extended description then add it.
        If Not ProgressiveDesc Is Nothing Then ButtonDefinition.ProgressiveToolTip.ExpandedDescription = ProgressiveDesc

        On Error Resume Next
        'If an image was passed then include it.
        If Not ProgressiveImage Is Nothing Then ButtonDefinition.ProgressiveToolTip.Image = DirectCast(GetIPictureDispFromPicture(ProgressiveImage), IPictureDisp)

        'If an image was passed then include it as well.
        If Not ProgressiveVideo Is Nothing Then ButtonDefinition.ProgressiveToolTip.Video = ProgressiveVideo

        'Return the definition.
        Return ButtonDefinition
    End Function

    ''' <summary>
    ''' Adds the Button to the Inventor UI based on specified Panel, Tab, and Ribbon.
    ''' </summary>
    Public Function AddButtonToUI() As CommandControl
        'If no tool ribbons were specified then we need to log the error and exit.
        If IncludeInRibbons.Count = 0 Then
            _Addin.Logging.WriteToErrorLog("Could not add " & ButtonName & " to the Inventor UI because no Tool Ribbons were specified.", _Addin)
            Return Nothing
        End If

        Dim ASY_RIB, DWG_RIB, PRT_RIB, NODOC_RIB As Ribbon
        Try
            'Add the Button to all Tool Ribbons.
            If IncludeInRibbons.Contains(RibbonType.All) Then
                ASY_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("Assembly")
                DWG_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("Drawing")
                PRT_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("Part")
                NODOC_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("ZeroDoc")
                AddButtonDefinitionToPanel(ASY_RIB)
                AddButtonDefinitionToPanel(DWG_RIB)
                AddButtonDefinitionToPanel(PRT_RIB)
                AddButtonDefinitionToPanel(NODOC_RIB)
            Else
                'Add the Button to only the specified Tool Ribbons.
                For Each SelectedRibbon In IncludeInRibbons
                    Select Case SelectedRibbon
                        Case RibbonType.Assembly
                            'Assembly Ribbon
                            ASY_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("Assembly") : AddButtonDefinitionToPanel(ASY_RIB)
                        Case RibbonType.Drawing
                            'Drawing Ribbon
                            DWG_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("Drawing") : AddButtonDefinitionToPanel(DWG_RIB)
                        Case RibbonType.Part
                            'Part Ribbon
                            PRT_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("Part") : AddButtonDefinitionToPanel(PRT_RIB)
                        Case RibbonType.ZeroDoc
                            'Zero Doc Ribbon
                            NODOC_RIB = _Addin.InvApp.UserInterfaceManager.Ribbons.Item("ZeroDoc") : AddButtonDefinitionToPanel(NODOC_RIB)
                        Case Else
                            'Do Nothing
                    End Select
                Next
            End If
        Catch ex As Exception
            _Addin.Logging.WriteToErrorLog("Could not add " & ButtonName & " to the Inventor UI because no Tool Ribbons were specified.", _Addin)
            Return Nothing
        Finally
            ASY_RIB = Nothing
            DWG_RIB = Nothing
            PRT_RIB = Nothing
            NODOC_RIB = Nothing
        End Try
        Return ButtonCommand
    End Function

    ''' <summary>
    ''' Find a Ribbon Tab using the TabInteranlName specified in the ControlDefinition.
    ''' Returns the RibbonTab or Nothing if it is not found.
    ''' </summary>
    ''' <param name="Ribbon">[Inventor.Ribbon] Ribbon to search.</param>
    ''' <param name="Create">[Boolean](Optional | Default = True) Creates a new Ribbon Tab if not found.</param>
    ''' <returns>[Inventor.RibbonTab] or Nothing</returns>
    Private Function FindRibbonTab(ByVal Ribbon As Ribbon, Optional Create As Boolean = True) As RibbonTab
        Try
            Return Ribbon.RibbonTabs.Item(TabInternalName)
        Catch ex As Exception
            If Create Then
                Return CreateRibbonTab(Ribbon)
            Else
                Return Nothing
            End If
        End Try
    End Function

    ''' <summary>
    ''' Creates a new Ribbon Tab using the TabName specified in the ControlDefinition.
    ''' Returns the RibbonTab or Nothing if it could not be created.
    ''' </summary>
    ''' <param name="Ribbon">[Inventor.Ribbon] The ribbon object to add the new Tab to.</param>
    ''' <returns>[Inventor.Ribbon] or Nothing</returns>
    Private Function CreateRibbonTab(ByVal Ribbon As Ribbon) As RibbonTab
        Return Ribbon.RibbonTabs.Add(TabName, TabInternalName, CLID, "id_TabTools")
    End Function

    ''' <summary>
    ''' Find a Ribbon Panel using the PanelInteranlName specified in the ControlDefinition.
    ''' Returns the RibbonPanel or Nothing if it is not found.
    ''' </summary>
    ''' <param name="RibbonTab">[Inventor.RibbonTab] RibbonTab to search.</param>
    ''' <param name="Create">[Boolean](Optional | Default = True) Creates a new Ribbon Panel if not found.</param>
    ''' <returns>[Inventor.RibbonPanel] or Nothing</returns>
    Private Function FindRibbonPanel(ByVal RibbonTab As RibbonTab, Optional Create As Boolean = True) As RibbonPanel
        Try
            Return RibbonTab.RibbonPanels.Item(PanelInternalName)
        Catch ex As Exception
            If Create Then
                Return CreateRibbonPanel(RibbonTab)
            Else
                Return Nothing
            End If
        End Try
    End Function

    ''' <summary>
    ''' Creates a new Ribbon Panel using the PanelName specified in the ControlDefinition.
    ''' Returns the RibbonPanel or Nothing if it could not be created.
    ''' </summary>
    ''' <param name="RibbonTab">[Inventor.RibbonTab] RibbonTab to search.</param>
    ''' <returns>[Inventor.RibbonPanel] or Nothing</returns>
    Private Function CreateRibbonPanel(ByVal RibbonTab As RibbonTab) As RibbonPanel
        Return RibbonTab.RibbonPanels.Add(PanelName, PanelInternalName, CLID)
    End Function

    ''' <summary>
    ''' Creates a CommandControl and adds it to the specified Ribbon Panel.
    ''' Returns the CommandControl or Nothing if it could not be created.
    ''' </summary>
    ''' <param name="RibbonPanel">[Inventor.RibbonPanel] Ribbon Panel that gets the button definition.</param>
    ''' <returns>[Inventor.CommandControl] or Nothing</returns>
    Private Function CreatCommandControl(ByVal RibbonPanel As RibbonPanel) As CommandControl
        Try
            Dim ControlButton As CommandControl = RibbonPanel.CommandControls.AddButton(ButtonDefinition)
            With ControlButton
                .UseLargeIcon = LargeButtonFormat
                .ShowText = ShowButtonText
            End With
            Return ControlButton
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Deletes the Command Control and ControlDefinition for the current control.
    ''' </summary>
    Public Sub RemoveCommand()
        On Error Resume Next
        _Addin.InvApp.CommandManager.ControlDefinitions.Item(ButtonInternalName).Delete()
    End Sub

    ''' <summary>
    ''' Adds the ButtonDefinition to the specified RibbonPanel.
    ''' This sub will try to find or create the specified RibbonTab and RibbonPanel defined in the ControlDefinition.
    ''' </summary>
    ''' <param name="Ribbon">[Inventor.Ribbon] The ribbon object to add the new items to.</param>
    Private Sub AddButtonDefinitionToPanel(ByVal Ribbon As Ribbon)

        Dim rTab As RibbonTab, rPanel As RibbonPanel
        Try
            'First validate that we have the needed parameters, else throw a specific error for each one.
            If TabName = Nothing OrElse TabName = "" Then Throw New ArgumentException("TabName is required when adding a command button to the UI!", ButtonName)
            If TabInternalName = Nothing OrElse TabInternalName = "" Then Throw New ArgumentException("TabInternalName is required when adding a command button to the UI!", ButtonName)
            If PanelName = Nothing OrElse PanelName = "" Then Throw New ArgumentException("PanelName is required when adding a command button to the UI!", ButtonName)
            If PanelInternalName = Nothing OrElse PanelInternalName = "" Then Throw New ArgumentException("PanelInternalName is required when adding a command button to the UI!", ButtonName)

            rTab = FindRibbonTab(Ribbon, True)
            rPanel = FindRibbonPanel(rTab, True)

            'If the rTab & rPanel have been found then add the command.
            If Not rTab Is Nothing AndAlso Not rPanel Is Nothing Then ButtonCommand = CreatCommandControl(rPanel)

        Catch ex As ArgumentException
            Dim Message As String = "A value for the parameter " & ex.ParamName &
                " was not specified during the creation of the """ & ButtonName &
                """ ButtonDefinition." & vbNewLine & "The control could not be added to the UI."
            _Addin.Logging.WriteToErrorLog(Message, _Addin, Ex:=ex)
        Catch ex As Exception
            _Addin.Logging.WriteToErrorLog("An error occurred while trying to add the " & ButtonName & " to the " & PanelName & " panel.", _Addin, Ex:=ex)
        Finally
            rTab = Nothing
            rPanel = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Turns the visibility of the Command Button On or Off.
    ''' </summary>
    ''' <param name="Visibile">[Boolean]</param>
    ''' <param name="InternalName">[String](Optional | Default = Calling Command Button) Specifies a specific button to turn on/off visibility. <para/>
    ''' If not specified the defaults to internal name of command calling the method.</param>
    Public Sub ChangeButtonVisibility(Visibile As Boolean, Optional InternalName As String = Nothing)
        Dim cPanel As RibbonPanel = _Addin.InvApp.UserInterfaceManager.ActiveEnvironment.Ribbon.RibbonTabs.Item(TabInternalName).RibbonPanels.Item(PanelInternalName)
        Dim inName As String = Nothing
        If Not InternalName Is Nothing Then
            inName = InternalName
        Else
            inName = ButtonInternalName
        End If

        Select Case Visibile
            Case True
                cPanel.CommandControls.Item(inName).Visible = True
            Case False
                cPanel.CommandControls.Item(inName).Visible = False
        End Select
    End Sub

End Class