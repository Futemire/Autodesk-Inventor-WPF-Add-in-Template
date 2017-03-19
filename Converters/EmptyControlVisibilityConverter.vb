Imports System.Globalization

Public Class EmptyControlVisibilityConverter : Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        If Not TypeOf (value) Is Integer Then Return Binding.DoNothing
        Select Case True
            Case value > 0
                Return Visibility.Visible
            Case Else
                Return Visibility.Collapsed
        End Select
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        If Not TypeOf (value) Is Visibility Then Return Binding.DoNothing
        Select Case True
            Case Visibility.Collapsed
                Return 0
            Case Visibility.Hidden
                Return 0
            Case Else
                Return 1
        End Select
    End Function
End Class