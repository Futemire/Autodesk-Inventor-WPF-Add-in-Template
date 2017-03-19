Imports System.Runtime.CompilerServices

Module FileSystemExtensions

    ''' <summary>
    ''' Checks the specified path and returns the last write time of the file.
    ''' </summary>
    ''' <param name="FilePath">[String] Full file path.</param>
    <Extension>
    Public Function GetLastWriteTime(FilePath As String) As Date
        Return My.Computer.FileSystem.GetFileInfo(FilePath).LastWriteTime
    End Function

    ''' <summary>
    ''' Checks the path of files or directories and returns [TRUE] if it exists.
    ''' </summary>
    ''' <param name="Path">[Sting] Path of file or directory to check.</param>
    ''' <returns>[Boolean]</returns>
    <Extension, DebuggerHidden>
    Public Function IsValidPath(ByVal Path As String) As Boolean
        'First check to see if the path exists as a file.
        If My.Computer.FileSystem.FileExists(Path) Then Return True
        'Then check to see if the path exists as a directory.
        If My.Computer.FileSystem.DirectoryExists(Path) Then Return True
        'If we made it here then the path does not exist at all...
        Return False
    End Function

    ''' <summary>
    ''' Ensures that a directory path exists, if not then tries to create it. <para/>
    ''' This method works on File and Directory paths.<para/>
    ''' Returns [True] if the path exists or was created, [False] is failed to create.
    ''' </summary>
    ''' <param name="Path">[String] A file or Directory path to check.</param>
    ''' <return>[True] or [False]</return>
    <Extension>
    Public Function EnsureDirectory(Path As String) As Boolean
        Try
            Dim DirPath = IO.Path.GetDirectoryName(Path)
            If Not My.Computer.FileSystem.DirectoryExists(DirPath) Then
                My.Computer.FileSystem.CreateDirectory(DirPath)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

End Module