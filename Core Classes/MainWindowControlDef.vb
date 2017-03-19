Imports System.Reflection
Imports System.Security.Principal
Imports System.Text
Imports System.Text.RegularExpressions
Imports $safeprojectname$.My
Imports Inventor

''' <summary>
''' This is a sample class that demonstrates creating a Control Definition by inheriting from the ControlDefBaseClass. <para/>
''' The ControlDefBaseClass handles the repetitive work of creating/accessing Tabs and Panels as well as adding Control Definitions to specified Panels. 
''' You can modify this to fit your needs, or delete it.
''' </summary>
Public Class MainWindowControlDef : Inherits ControlDefBaseClass

    'Inventor's button definition for this ControlDef.
    Public WithEvents ButtonDef As ButtonDefinition

    'The window that this control definition will display/interact with.
    Public WithEvents MainWin As MainWindow

    'Use Inventor's Transaction manager to set undo checkpoints.
    Public TransMan As TransactionManager
    Public AddinTransAction As Transaction

    'Events to be used in the current ControlDef.
    Public WithEvents AppEvents As ApplicationEvents

    ''' <summary>
    ''' [Private] Active Inventor Document reference object.
    ''' </summary>
    Private _ActDoc As Document

    ''' <summary>
    ''' [Private] Document sub type such as "SolidPart" or "Sheetmetal".
    ''' </summary>
    Private _ComponentType As InventorDocumentSubType

    ''' <summary>
    ''' [Private] AddinBaseClass reference object.
    ''' </summary>
    Private _Addin As AddinBaseClass


    ''' <summary>
    ''' Creates the ComponentDefinition.
    ''' </summary>
    ''' <param name="Addin"></param>
    Public Sub New(ByVal Addin As AddinBaseClass)

        'Initialize the base class.
        'This MUST be called before anything else.
        MyBase.New(Addin)

        'Populate to the private object.
        _Addin = Addin

        'Create the balloon tips that will be displayed to the user.
        'This is less intrusive than constantly displaying message boxes.
        _Addin.InvApp.UserInterfaceManager.BalloonTips.Add("$companysafename$_$addinsafename$_HeyNowBalloonTip", '<---- Company Name, Addin Name, Balloon Tip Name, No Spaces
                                                           "Hey Now", '<---- Balloon Tip Display Name, Spaces OK
                                                           "Hey Now, This is a balloon tip!"
                                                           )


        'Hookup the event objects needed for this control.
        AppEvents = _Addin.InvApp.ApplicationEvents

        'Define the button name & internal name. The Button Name can be anything but the Internal Name must be unique and cannot contain spaces.
        'Because there are nearly 2000 button definition already in Inventor I typically create Internal Names by using this format...
        'MyCompanyName_MyAddinSafeName_MyButtonName

        ButtonName = "$addinname$ Button 1"
        LargeIcon = My.Resources.Settings
        SmallIcon = My.Resources.Settings_32x32

        ButtonInternalName = "$companysafename$_$addinsafename$_MainWindow"
        ShowButtonText = True


        'General hover text to be displayed when user hovers over the button.
        'Note that this text will ignored if you decide to use the progressive tool tip.
        ToolTip = "Launch $addinname$."

        'This feature is optional. If set to True then your button will display a more elaborate tool tip with a
        'title, space for a longer button description, and the option to add an image or video string.
        Progressive = False
        ProgressiveTitle = Nothing
        ProgressiveDesc = Nothing
        ProgressiveImage = Nothing
        ProgressiveVideo = Nothing

        'Create a new list of the ribbons you want to add the button to.
        IncludeInRibbons = New List(Of RibbonType)({RibbonType.Part, RibbonType.Assembly})

        'Specify the Internal Name of the Ribbon Tab that the button should be included in.
        'If Left blank then a control definition will be created but NOT added to UI.
        'If a Ribbon Tab with the specified InternalName is found then the existing object is used.
        'If a Ribbon Tab with the specified InternalName is not fount then it is created.
        'Use the same nomenclature as explained for the ButtonInternalName above.
        '<Required if adding button to UI.>
        TabInternalName = "$companysafename$_$addinsafename$_Tab"

        'Friendly Name of Tab to create if InteranlName is not found. This name can contain spaces.
        TabName = "$companyname$"

        'Specify the Internal Name of the Ribbon Panel that the button should be included in.
        'If Left blank then a control definition will be created but NOT added to UI.
        'If a Ribbon Panel with the specified InternalName is found then the existing object is used.
        'If a Ribbon Panel with the specified InternalName is not fount then it is created.
        'Use the same nomenclature as explained for the ButtonInternalName above.
        '<Required if adding button to UI.>
        PanelInternalName = "$companysafename$_$addinsafename$_Panel1"

        'Friendly Name of Tab to create if InteranlName is not found. This name can contain spaces.
        PanelName = "$addinname$"

        'Create the control button definition and link it to your Global Private WithEvents button def above.
        ButtonDef = CreateControlDefButton()

        'Make the call to add the button to the User Interface.
        Call AddButtonToUI()

    End Sub

    ''' <summary>
    ''' Launch the main window for the $addinname$ add-in.
    ''' </summary>
    Private Sub ButtonDefinition_OnExecute(Context As NameValueMap) Handles ButtonDef.OnExecute
        Try

            'Don't allow multiple instances of the window.
            If Not MainWin Is Nothing Then Exit Sub

            'Set the Active Document & Active RangeBox
            _ActDoc = _Addin.InvApp.ActiveDocument

            'Get the active document type.
            _ComponentType = _Addin.InvApp.ActiveDocument.GetDocumentSubType

            'Get the Active Document.
            Select Case True
                Case _ComponentType = InventorDocumentSubType.Assembly
                    _ActDoc = DirectCast(_ActDoc, AssemblyDocument)
                Case _ComponentType = InventorDocumentSubType.SolidPart
                    _ActDoc = DirectCast(_ActDoc, PartDocument)
                Case _ComponentType = InventorDocumentSubType.SheetMetal
                    _ActDoc = DirectCast(_ActDoc, PartDocument)
                Case Else
                    'We don't have a valid document type.
                    Exit Sub
            End Select

            'Create the transaction manager and start a transaction.
            TransMan = _Addin.InvApp.TransactionManager
            AddinTransAction = TransMan.StartTransaction(_ActDoc, "$addinname$")

            'Create the window object.
            MainWin = New MainWindow(_Addin)

            'You could create events in your MainWindow and add handlers here to handle user interactions.
            'It would look something like this...
            AddHandler MainWin.MyEvent, AddressOf SomeMethodInThisClass

            'Now show the window.
            MainWin.ShowDialog()

            'End the transaction so it is committed.
            'We won't get here until the MainWin is closed.
            If Not AddinTransAction Is Nothing Then
                AddinTransAction.End()
                AddinTransAction = Nothing
            End If

        Catch ex As Exception
            'On error we need to abort the transaction if it is uncommitted.
            If Not AddinTransAction Is Nothing Then
                AddinTransAction.Abort()
                AddinTransAction = Nothing
            End If

            'Log the error.
            _Addin.Logging.WriteToErrorLog("There was an error opening the form!", _Addin, ex)
        End Try
    End Sub


#Region "Inventor Events"
    'An example of where you can handle events that this control def has hooked into.
    Private Sub AppEvents_OnActivateDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum,
                                             Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum) Handles AppEvents.OnActivateDocument
        Select Case BeforeOrAfter
            Case EventTimingEnum.kBefore

            Case EventTimingEnum.kAfter

            Case EventTimingEnum.kAbort
        End Select

    End Sub

    Private Sub AppEvents_OnSaveDocument(DocumentObject As _Document, BeforeOrAfter As EventTimingEnum,
                                         Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum) Handles AppEvents.OnSaveDocument
        Select Case BeforeOrAfter
            Case EventTimingEnum.kBefore

            Case EventTimingEnum.kAfter

            Case EventTimingEnum.kAbort

        End Select
    End Sub
#End Region

#Region "Sample Event Handler"
    ''' <summary>
    ''' Just a sample event handler to show how to handle events from your MainWindow. <para/>
    ''' You can modify this to fit your needs, or delete it.
    ''' </summary>
    Public Sub SomeMethodInThisClass()
        MsgBox("My Main Window Fired An Event!")
    End Sub
#End Region
End Class