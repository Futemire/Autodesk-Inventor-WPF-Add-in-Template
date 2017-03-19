Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Inventor

''' <summary>
''' For global add-in variables and methods only that rarely change.
''' Best practice is to wrap variables and properties in classes, so add here sparingly and only if it doesn't need to be managed.
''' </summary>
Public Module AddinInfo

    ''' <summary>
    ''' Returns the Name attribute of the calling assembly.
    ''' </summary>
    ''' <returns>[String]</returns>
    Public ReadOnly Property AddinName As String
        Get
            Dim ProductName As AssemblyProductAttribute = Assembly.GetCallingAssembly.GetCustomAttribute(GetType(AssemblyProductAttribute))
            Return ProductName.Product
        End Get
    End Property

    ''' <summary>
    ''' Safe name of the add-in containing no spaces and beginning with an underscore or letter.
    ''' </summary>
    ''' <returns>[String]</returns>
    Public ReadOnly Property SafeName As String
        Get
            'Replace Spaces with underscore.
            Dim Name = AddinName.Replace(" ", "_")
            'Add underscore if first character is a number.
            If IsNumeric(Name.Chars(0)) Then Name = "_" & Name
            Return Name
        End Get
    End Property

    ''' <summary>
    ''' Returns the Description attribute of the calling assembly.
    ''' </summary>
    ''' <returns>[String]</returns>
    Public ReadOnly Property AddinDescription As String
        Get
            Dim ProductDescription As AssemblyDescriptionAttribute = Assembly.GetCallingAssembly.GetCustomAttribute(GetType(AssemblyDescriptionAttribute))
            Return ProductDescription.Description
        End Get
    End Property

    ''' <summary>
    ''' Returns the Version attribute of the calling assembly.
    ''' </summary>
    ''' <returns>[String]</returns>
    Public ReadOnly Property AddinVersion As String
        Get
            Dim AssemblyVersion As AssemblyFileVersionAttribute = Assembly.GetCallingAssembly.GetCustomAttribute(GetType(AssemblyFileVersionAttribute))
            Return AssemblyVersion.Version
        End Get
    End Property

    ''' <summary>
    ''' Client ID of the add-in.
    ''' </summary>
    ''' <returns>[String]</returns>
    Public ReadOnly Property CLID As String
        Get
            Return "$guid1$"
        End Get
    End Property

    ''' <summary>
    ''' Gets or creates the AppData path for the current add-in. <para/>
    ''' </summary>
    ''' <param name="SubDirectoryName">[String] Subdirectory to append to AppData path.</param>
    ''' <returns>[String]</returns>
    Public ReadOnly Property AddinAppDataPath(Optional SubDirectoryName As String = Nothing) As String
        Get
            Dim AppDataDirectory As String
            Select Case True
                Case SubDirectoryName Is Nothing
                    AppDataDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Autodesk\Inventor Addins\"
                Case SubDirectoryName = ""
                    AppDataDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Autodesk\Inventor Addins\"
                Case Else
                    SubDirectoryName.Replace("\", "")
                    AppDataDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) & "\Autodesk\Inventor Addins\" & SubDirectoryName & "\"
            End Select
            'Create the AppData directory if it did not exist.
            If Not My.Computer.FileSystem.DirectoryExists(AppDataDirectory) Then
                My.Computer.FileSystem.CreateDirectory(AppDataDirectory)
            End If
            Return AppDataDirectory
        End Get
    End Property

End Module

<ProgId("$safeprojectname$.InventorAddin"), Guid("$guid1$"), ComVisible(True)>
Public Class AddinBaseClass
    Implements ApplicationAddInServer

    ''' <summary>
    ''' The Inventor Application object referenced by this add-in.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property InvApp As Inventor.Application
        Get
            Return _InvApp
        End Get
    End Property
    Private _InvApp As Inventor.Application

    ''' <summary>
    ''' The error logging class for the add-in.
    ''' </summary>
    Friend WithEvents Logging As ErrorLogging

    ''' <summary>
    '''  Enables or Disables this add-in for watching for log changes via the FileSystemWatcher.
    ''' </summary>
    Friend EnableLogWatching As Boolean = False

    ''' <summary>
    ''' Location of Application Data folder on current machine.
    ''' Windows best practice is to use the AppData folders to save application data from you app.
    ''' Saving in ProgramData or other locations you can start to run into issues with having to have elevated permissions.
    ''' Appdata Roaming is usefull to save app settings or config info for your app that can be used across multiple machines
    ''' that a user logs into with the same Active Directlry account. This is great in an enterprise environment.
    ''' </summary>
    Public AppDataDirectory As String

    ''' <summary>
    ''' An example Button Control Definition.
    ''' You dont need to use WithEvents if your ControlDef does not raise events.
    ''' See Setting Up Control Definitions in this project.
    ''' </summary>
    Public WithEvents MainWindowCDef As MainWindowControlDef

    ''' <summary>
    ''' Inventor Application Events
    ''' </summary>
    Public WithEvents AppEvents As ApplicationEvents

    ''' <summary>
    ''' Method required by Autodesk Inventor API Interface.
    ''' This property is provided to allow the Add-in to expose an API.
    ''' </summary>
    Public ReadOnly Property Automation As Object Implements ApplicationAddInServer.Automation
        Get
            Return Me
        End Get
    End Property

    ''' <summary>
    ''' This method is called by Inventor when it loads the Add-in.
    ''' </summary>
    ''' <param name="addInSiteObject">[Inventor.AddInSiteObject]  Provides access to the Inventor Application object.</param>
    ''' <param name="firstTime">[Boolean] Indicates if the Add-in is loaded for the first time in the current Inventor session.</param>
    Public Sub Activate(AddInSiteObject As ApplicationAddInSite, FirstTime As Boolean) Implements ApplicationAddInServer.Activate
        Try
            If FirstTime Then
                'Setup the Unhandled Exception Handler.
                Dim curDomain = AppDomain.CurrentDomain : AddHandler curDomain.UnhandledException, AddressOf AddinUnhandledExceptionHandler

                'Load the private variables...
                _InvApp = AddInSiteObject.Application

                'Hook into events...
                AppEvents = _InvApp.ApplicationEvents

                'Instantiate the error logging class.
                Logging = New ErrorLogging(My.Settings.ErrorLogDirectory)
                '                                ▲
                '                                └─── This is not the best solution and should be changed.
                '                                     For the sake of simplicity in this template I removed my custom configuration 
                '                                     file that can be administered without modifying or rebuilding code.
                '                                     Consider implementing your own to avoid hard coded paths or at lease 
                '                                     revalidate the saved path on each startup and give the user the ability 
                '                                     to update at will..


                'The control definitions are defined above so that they can be accessed anywhere the AddinBaseClass is passed.
                'Create an instance of the control definitions. 
                'The MainWindowCDef creates a button that opens the MainWindow in this project when clicked. 
                MainWindowCDef = New MainWindowControlDef(Me)

            End If

        Catch ex As Exception
            'Write the error to the XML Error Log.
            Logging.WriteToErrorLog("The $addinname$ Add-in was not loaded!", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Method required by Autodesk Inventor API Interface.
    ''' </summary>
    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
        'Only release if we have an application object.
        If Not _InvApp Is Nothing Then Marshal.ReleaseComObject(_InvApp)

        'Call Garbage Collection
        GC.WaitForPendingFinalizers()
        GC.Collect()
    End Sub

    ''' <summary>
    ''' Method required by Autodesk Inventor API Interface.
    ''' <para/>
    ''' Depreciated. This method is required to be included for legacy purposes but needs to remain empty.
    ''' </summary>
    Public Sub ExecuteCommand(CommandID As Integer) Implements ApplicationAddInServer.ExecuteCommand
        'Do Nothing
    End Sub

    ''' <summary>
    ''' Sets the Inventor window as the parent/owner of the specified form/window.
    ''' </summary>
    ''' <param name="CurrWind">[Window] The window to set as the child of the Inventor Window.</param>
    ''' <returns>[Boolean] Returns TRUE on Success.</returns>
    ''' <remarks>Setting an add-in window/dialog/form as a child will make sure it is hidden when the main Inventor window is minimized.</remarks>
    <DebuggerHidden>
    Public Function SetInventorAsOwnerWindow(CurrWind As Window) As Boolean

        Try
            'Find the Inventor Window Handle.
            Dim InvWndHnd As IntPtr = IntPtr.Zero

            'Search the process list for the matching Inventor application.
            For Each pList As Process In Process.GetProcesses()
                If pList.MainWindowTitle.Contains(InvApp.Caption) Then
                    InvWndHnd = pList.MainWindowHandle
                    Exit For
                End If
            Next

            Dim InvWndIrp = New Interop.WindowInteropHelper(CurrWind)
            InvWndIrp.EnsureHandle()
            InvWndIrp.Owner = InvWndHnd
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Handles all "Unhandled" exceptions caught in the add-in.
    ''' </summary>
    Sub AddinUnhandledExceptionHandler(sender As Object, args As UnhandledExceptionEventArgs)
        MsgBox("Unhandled Exception was caught!" & vbNewLine & DirectCast(args.ExceptionObject, Exception).ToString)
    End Sub

    ''' <summary>
    ''' Allows user to select a config file for the add-in.
    ''' <para/>
    ''' Returns TRUE on success, FALSE if canceled or failed.
    ''' </summary>
    ''' <returns>[Boolean]</returns>
    Private Function SetConfigFile() As Boolean
        'Create the file open dialog.
        Dim FOD As New OpenFileDialog
        With FOD
            .AddExtension = True
            .CheckFileExists = True
            .CheckPathExists = True
            .Filter = "Config File (*.Config) | * .Config"
            .Multiselect = False
            .Title = "Select the Build StockNumber configuration file..."
        End With

RETRY_LOAD_CONFIG:
        'If a file was selected then check the path and set it.
        Select Case FOD.ShowDialog
            Case DialogResult.OK
                'First make sure the path exists.
                If Not IsValidPath(FOD.FileName) Then GoTo RETRY_LOAD_CONFIG
                'Next make sure it is a valid config file for the add-in.
                Dim ConfigText = IO.File.ReadAllText(FOD.FileName)
                If Not ConfigText.Contains("addin_name=""$safeprojectname$""") Then
                    MsgBox("The selected file is not a valid config file for the Build StockNumber add-in!" & vbNewLine & "Please make another selection...")
                    GoTo RETRY_LOAD_CONFIG
                End If
                My.Settings.ConfigDocPath = FOD.FileName
                My.Settings.Save()
                Return True
        End Select

        'Either a path couldn't be found or the user canceled.
        Return False
    End Function

    Protected Function ValidateConfigFile()
REVALIDATE:
        'First check to see if the saved config file is valid, if not return false.
        If Not CheckPath(My.Settings.ConfigDocPath) Then
            MsgBox("The config file for the Build StockNumber add-in could not be loaded!" & vbNewLine & "Please select a valid config file...")
            If Not SetConfigFile() Then Return False Else GoTo REVALIDATE
        End If

        Dim ConfigText = IO.File.ReadAllText(My.Settings.ConfigDocPath)
        If Not ConfigText.Contains("addin_name=""$safeprojectname$""") Then
            MsgBox("The selected file is not a valid config file for the Build StockNumber add-in!" & vbNewLine & "Please select a valid config file...")
            If Not SetConfigFile() Then Return False Else GoTo REVALIDATE
        End If

        'Next see if the config file has an updated path element.
        Dim Config As XDocument = XDocument.Load(My.Settings.ConfigDocPath)
        Dim xUpdatedPath As XElement = Config.Root.Descendants("UpdatedConfigPath")(0)

        If Not xUpdatedPath Is Nothing Then
            'We found an UpdatedConfigPath element so lets check to make sure it is valid.
            Dim NewPath As XAttribute = xUpdatedPath.Attribute("path")
            If Not NewPath Is Nothing Then

                'Check the NewPath.
                If IsValidPath(NewPath.Value) Then
                    'The new path exists so next make sure it is a valid config file for the add-in.
                    Dim NewConfigText = IO.File.ReadAllText(NewPath.Value)
                    If Not NewConfigText.Contains("addin_name=""$safeprojectname$""") Then
                        MsgBox("The Updated Config file path listed in the current config file for the Build StockNumber add-in is invalid!" & vbNewLine & "Please make another selection...")
                        If Not SetConfigFile() Then Return False Else GoTo REVALIDATE
                    End If
                    My.Settings.ConfigDocPath = NewPath.Value
                    My.Settings.Save()
                End If
            End If
        End If

        Return True
    End Function

End Class