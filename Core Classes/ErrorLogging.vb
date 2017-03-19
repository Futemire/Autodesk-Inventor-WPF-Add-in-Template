Imports System.DirectoryServices.AccountManagement
Imports System.Reflection
Imports System.Runtime.CompilerServices

''' <summary>
''' Notes about the ErrorLogging class. <para/>
''' This class will allow your application to write its errors to a specific XML document for each user. <para/>
''' The file name will be [WindowsAccountUserName].xml in the directory specified when creating the class. <para/>
''' If a directory is not supplied when creating the class then the default directory will be the user's AppData directory for the add-in.
''' </summary>
Public Class ErrorLogging

    ''' <summary>
    ''' Gets/Sets a validated path for the Error Log Directory from the application settings.
    ''' </summary>
    ''' <returns>[String] A valid directory to store error logs.</returns>
    Public Property CurrentErrorLogDirectory() As String
        Set(value As String)
            If value.EnsureDirectory Then
                My.Settings.ErrorLogDirectory = value
                My.Settings.Save()
            Else
                Throw New System.IO.FileNotFoundException("The specified error log could not be found or created!", value)
            End If
        End Set
        Get
            Return My.Settings.ErrorLogDirectory
        End Get
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ErrorLogDirectory">[String](Optional | Default = Nothing) The path where the error log should be created. <para/>
    ''' Be sure it is in a location that the current application user can write to!
    ''' </param>
    Public Sub New(Optional ErrorLogDirectory As String = Nothing)

        Try
            'Pass the optional parameter to the CurrentErrorLogDirectory property, even if it is nothing.
            CurrentErrorLogDirectory = ErrorLogDirectory

        Catch ex As Exception
            'If we made it here then either the optional ErrorLogDirectory parameter was nothing or invalid.
            'So create an Error Logs folder under the current user's AppData directory for the add-in.
            CurrentErrorLogDirectory = AddinAppDataPath & "Error Logs\"
        End Try

    End Sub

    ''' <summary>
    ''' Creates an XML Style sheet for the error log.
    ''' </summary>
    Public Sub CreateErrorStyleSheet()
        Dim ErrorStyles As XDocument = Nothing
        ErrorStyles =
        <?xml version="1.0" encoding="utf-8"?>
        <xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
            <xsl:preserve-space elements="*"/>
            <xsl:template match="/">
                <html>
                    <head>
                        <script src="https://code.jquery.com/jquery-1.12.0.min.js"/>
                        <script src="https://code.jquery.com/ui/1.11.4/jquery-ui.min.js"/>
                        <link rel="stylesheet" type="text/css" href="https://code.jquery.com/ui/1.11.4/themes/black-tie/jquery-ui.css"/>
                        <style>
                                                                body { background-color:lightgrey; }
                                                                strong { font-variant: small-caps; }
                                                                h1, h2, h3, h4, h5 { margin-top: -5px; margin-bottom: -5px; }
                                                                .Summary { background-color: #000; color: #FFF; margin: 25px 25px 0px 25px; padding: 10px; border-radius: 5px 5px 0 0; position: relative; }
                                                                .Exception { background-color: #FF1A3D; border: Double 6px #000; color: #FFF; display: inline-block; margin -5px 20px 20px 20px; padding: 10px; line-height: 1.4; }
                                                                .ui-accordion { margin: 0 25px; }
                                                                .ui-widget { font-family: Ebrima; font-size: 16px; }
                                                                .Error { box-shadow: #666 2px 0 3px inset;  padding: 10px; position: relative; clear: b }
                                                                 #Errors { margin-top: -5px; }
                                    </style>
                        <script>
                                                                $(document).ready(function(){ $( "#Errors" ).accordion({ heightStyle: "content", collapsible: true, active: false }); });
                                    </script>
                    </head>
                    <body>
                        <div class="Summary">
                            <h1>Error Summary | <xsl:value-of select="ErrorLog/Errors/Error/Addin/@name"/></h1>
                            <hr/>
                            <strong>Name:
                                        </strong>
                            <xsl:value-of select="ErrorLog/ErrorLogSummary/@full_name"/>
                            <br/>
                            <strong>Login:
                                        </strong>
                            <xsl:value-of select="ErrorLog/ErrorLogSummary/@user_name"/>
                            <br/>
                            <strong>Error Count:
                                        </strong>
                            <xsl:value-of select="ErrorLog/ErrorLogSummary/@error_count"/>
                            <br/>
                            <strong>Last Project:
                                        </strong>
                            <xsl:value-of select="ErrorLog/ErrorLogSummary/@last_project"/>
                            <br/>
                            <strong>Last Error:
                                        </strong>
                            <xsl:value-of select="ErrorLog/ErrorLogSummary/@last_error_occurred"/>
                        </div>
                        <div id="Errors">
                            <xsl:for-each select="ErrorLog/Errors/Error">
                                <xsl:sort select="position()" data-type="number" order="descending"/>
                                <h3>
                                    <strong>Error No. <xsl:value-of select="@error_number"/> | <xsl:value-of select="Message"/></strong>
                                </h3>
                                <p class="Error">
                                    <strong>Time:
                                                </strong>
                                    <xsl:value-of select="@time"/>
                                    <strong> | Inventor Version:
                                                </strong>
                                    <xsl:value-of select="Inventor/@version"/>
                                    <strong> | Project:
                                                </strong>
                                    <xsl:value-of select="Inventor/@project"/>
                                    <strong> | Addin Version:
                                                </strong>
                                    <xsl:value-of select="Addin/@version"/>
                                    <br/>
                                    <strong>Addin Config:
                                                </strong>
                                    <xsl:value-of select="CurrentConfig/@path"/>
                                    <br/>
                                    <br/>
                                    <span class="Exception">
                                        <strong>Error Message:</strong>
                                        <br/>
                                        <xsl:value-of select="Message"/>
                                        <br/>
                                        <br/>
                                        <strong>Exception Message:</strong>
                                        <br/>
                                        <xsl:value-of select="ExceptionMessage"/>
                                        <br/>
                                        <br/>
                                        <strong>Stack Trace:</strong>
                                        <br/>
                                        <xsl:value-of select="ExceptionStack"/>
                                    </span>
                                </p>
                            </xsl:for-each>
                        </div>
                    </body>
                </html>
            </xsl:template>
        </xsl:stylesheet>

        'Save the style sheet.
        ErrorStyles.Save(CurrentErrorLogDirectory & "ErrorStyles.xsl")
    End Sub

    ''' <summary>
    ''' Writes the error to a log specific to the current user.
    ''' </summary>
    ''' <param name="Message">[String] User friendly message.</param>
    ''' <param name="Ex">[Exception] Exception to be included in log.</param>
    Public Sub WriteToErrorLog(ByVal Message As String,
                               Optional Ex As Exception = Nothing,
                               Optional AdditionalElements As List(Of XElement) = Nothing)

        On Error Resume Next
        'Exit if the file path is invalid or can not be created.
        If Not CurrentErrorLogDirectory.EnsureDirectory Then Exit Sub

        Dim LogDirectoryPath = CurrentErrorLogDirectory & Environment.UserName & ".xml"

        Dim xErrorLog As New XDocument

        'Write to the log.
        If Not LogDirectoryPath.IsValidPath Then
            'The file did not exist so create the file and header with the first entry.

            xErrorLog =
                        <?xml version="1.0" encoding="utf-8"?>
                        <?xml-stylesheet type='text/xsl' href='ErrorStyles.xsl'?>
                        <!--User Error Log generated by Inventor Add-in Monitor -->
                        <ErrorLog>
                            <ErrorLogSummary
                                full_name=<%= UserPrincipal.Current.Name %>
                                user_name=<%= Environment.UserName %>
                                email=<%= UserPrincipal.Current.EmailAddress %>
                                error_count="1"
                                last_error_occurred=<%= Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString %>
                            />
                            <Errors>
                                <Error error_number="1" time=<%= Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString %>>
                                    <Addin name=<%= AddinName %> version=<%= AddinVersion %>/>
                                    <CurrentConfig path=<%= My.Settings.ConfigDocPath %>/>
                                    <%= If(IsNothing(AdditionalElements), Nothing, From ele In AdditionalElements) %>
                                    <Message><%= New XCData(Message) %></Message>
                                    <%= If(IsNothing(AdditionalElements), Nothing, From ele In AdditionalElements) %>
                                    <%= If(IsNothing(Ex), Nothing, <ExceptionMessage><%= New XCData(Ex.Message) %></ExceptionMessage>) %>
                                    <%= If(IsNothing(Ex), Nothing, <ExceptionStack><%= New XCData(Ex.StackTrace) %></ExceptionStack>) %>
                                </Error>
                            </Errors>
                        </ErrorLog>

        Else
            'Load the existing error log.
            xErrorLog = XDocument.Load(LogDirectoryPath)

            'Get current error number.
            Dim CurrCount As Integer = CInt(xErrorLog.Root.Element("ErrorLogSummary").Attribute("error_count").Value) + 1

            'Create a new Error element.
            Dim xError =
                        <Error error_number=<%= CurrCount %> time=<%= Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString %>>
                            <Addin name=<%= AddinName %> version=<%= AddinVersion %>/>
                            <CurrentConfig path=<%= My.Settings.ConfigDocPath %>/>
                            <%= If(IsNothing(AdditionalElements), Nothing, From ele In AdditionalElements) %>
                            <Message><%= New XCData(Message) %></Message>
                            <%= If(IsNothing(AdditionalElements), Nothing, From ele In AdditionalElements) %>
                            <%= If(IsNothing(Ex), Nothing, <ExceptionMessage><%= New XCData(Ex.Message) %></ExceptionMessage>) %>
                            <%= If(IsNothing(Ex), Nothing, <ExceptionStack><%= New XCData(Ex.StackTrace) %></ExceptionStack>) %>
                        </Error>

            'Add the error element to the document.
            xErrorLog.Root.Element("Errors").Add(xError)

            'Update the ErrorLogSummary
            'Error Count
            xErrorLog.Root.Element("ErrorLogSummary").Attribute("error_count").Value = CurrCount
            'Last Error Time
            xErrorLog.Root.Element("ErrorLogSummary").Attribute("last_error_occurred").Value = Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString

        End If

        'Save the error log.
        xErrorLog.Save(LogDirectoryPath)
        xErrorLog = Nothing

        'Lets make sure that an ErrorStyles sheet exists in the saved path.
        If Not CurrentErrorLogDirectory & "ErrorStyles.xsl".IsValidPath Then
            'One did not exist so lets create it.
            Call CreateErrorStyleSheet()
        End If

    End Sub

    ''' <summary>
    ''' Writes the error to a log specific to the current user.
    ''' </summary>
    ''' <param name="Message">[String] User friendly message.</param>
    ''' <param name="Addin">[AddinBaseClass] Base class for the add-in.</param>
    ''' <param name="Ex">[Exception] Exception to be included in log.</param>
    Public Sub WriteToErrorLog(ByVal Message As String,
                               ByVal Addin As AddinBaseClass,
                               Optional Ex As Exception = Nothing,
                               Optional AdditionalElements As List(Of XElement) = Nothing)

        On Error Resume Next
        'Exit if the file path is invalid or can not be created.
        'If My.Computer.FileSystem.EnsureDirectory(ErrorLogDirectory) Is Nothing Then Exit Sub

        Dim LogDirectoryPath = CurrentErrorLogDirectory & Environment.UserName & ".xml"

        Dim InventorProject As String = Addin.InvApp.DesignProjectManager.ActiveDesignProject.Name
        Dim InventorVersion As String = Addin.InvApp.SoftwareVersion.DisplayVersion
        Dim xErrorLog As New XDocument

        'Write to the log.
        If Not LogDirectoryPath.IsValidPath Then
            'The file did not exist so create the file and header with the first entry.

            xErrorLog =
                        <?xml version="1.0" encoding="utf-8"?>
                        <?xml-stylesheet type='text/xsl' href='ErrorStyles.xsl'?>
                        <!--User Error Log generated by Inventor Add-in Monitor -->
                        <ErrorLog>
                            <ErrorLogSummary
                                full_name=<%= UserPrincipal.Current.Name %>
                                user_name=<%= Environment.UserName %>
                                email=<%= UserPrincipal.Current.EmailAddress %>
                                error_count="1"
                                last_project=<%= InventorProject %>
                                last_error_occurred=<%= Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString %>
                            />
                            <Errors>
                                <Error error_number="1" time=<%= Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString %>>
                                    <Addin name=<%= AddinName %> version=<%= AddinVersion %>/>
                                    <CurrentConfig path=<%= My.Settings.ConfigDocPath %>/>
                                    <Inventor project=<%= InventorProject %> version=<%= InventorVersion %>/>
                                    <Message><%= New XCData(Message) %></Message>
                                    <%= If(IsNothing(AdditionalElements), Nothing, From ele In AdditionalElements) %>
                                    <%= If(IsNothing(Ex), Nothing, <ExceptionMessage><%= New XCData(Ex.Message) %></ExceptionMessage>) %>
                                    <%= If(IsNothing(Ex), Nothing, <ExceptionStack><%= New XCData(Ex.StackTrace) %></ExceptionStack>) %>
                                </Error>
                            </Errors>
                        </ErrorLog>

        Else
            'Load the existing error log.
            xErrorLog = XDocument.Load(LogDirectoryPath)

            'Get current error number.
            Dim CurrCount As Integer = CInt(xErrorLog.Root.Element("ErrorLogSummary").Attribute("error_count").Value) + 1

            'Create a new Error element.
            Dim xError =
                          <Error error_number=<%= CurrCount %> time=<%= Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString %>>
                              <Addin name=<%= AddinName %> version=<%= AddinVersion %>/>
                              <CurrentConfig path=<%= My.Settings.ConfigDocPath %>/>
                              <Inventor project=<%= InventorProject %> version=<%= InventorVersion %>/>
                              <Message><%= New XCData(Message) %></Message>
                              <%= If(IsNothing(AdditionalElements), Nothing, From ele In AdditionalElements) %>
                              <%= If(IsNothing(Ex), Nothing, <ExceptionMessage><%= New XCData(Ex.Message) %></ExceptionMessage>) %>
                              <%= If(IsNothing(Ex), Nothing, <ExceptionStack><%= New XCData(Ex.StackTrace) %></ExceptionStack>) %>
                          </Error>

            'Add the error element to the document.
            xErrorLog.Root.Element("Errors").Add(xError)

            'Update the ErrorLogSummary
            'Error Count
            xErrorLog.Root.Element("ErrorLogSummary").Attribute("error_count").Value = CurrCount
            'Last Error Time
            xErrorLog.Root.Element("ErrorLogSummary").Attribute("last_error_occurred").Value = Date.Now.ToShortDateString & " @ " & Date.Now.ToShortTimeString
            'Last Project
            xErrorLog.Root.Element("ErrorLogSummary").Attribute("last_project").Value = InventorProject

        End If

        'Save the error log.
        xErrorLog.Save(LogDirectoryPath)
        xErrorLog = Nothing

        'Lets make sure that an ErrorStyles sheet exists in the saved path.
        If Not CurrentErrorLogDirectory & "ErrorStyles.xsl".IsValidPath Then
            'One did not exist so lets create it.
            Call CreateErrorStyleSheet()
        End If
    End Sub

End Class
