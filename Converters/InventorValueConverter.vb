Imports System.Globalization
Imports Inventor
Public Class InventorValueConverter : Implements IMultiValueConverter

    ''' <summary>
    ''' Converts Inventor internal parameter value to correct units in lists. <para/>
    ''' values(0) must be the Inventor Parameter. <para/>
    ''' values(1) is the internal value.
    ''' </summary>
    ''' <param name="values">[Array] Inventor Parameter and Parameter Internal Value to be converted.</param>
    ''' <returns>[String]</returns>
    Public Function Convert(values() As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IMultiValueConverter.Convert
        Try
            Dim InvDoc As Document = values(0)
            Dim UOM As UnitsOfMeasure = InvDoc.UnitsOfMeasure
            Dim ConvertedValue As String = UOM.GetStringFromValue(values(1), UOM.LengthUnits)
            Return ConvertedValue
        Catch ex As Exception
            Return Binding.DoNothing
        End Try
    End Function

    ''' <summary>
    ''' ConvertBack not implemented.
    ''' </summary>
    Public Function ConvertBack(value As Object, targetTypes() As Type, parameter As Object, culture As CultureInfo) As Object() Implements IMultiValueConverter.ConvertBack
        Return Binding.DoNothing
    End Function
End Class