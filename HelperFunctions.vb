Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Module HelperFunctions
    <Extension()>
    Public Function GetEnumDescription(Of T)(ByVal e As T) As String
        If e.GetType().IsEnum Then
            Dim type As Type = e.GetType()
            Dim values As Array = [Enum].GetValues(type)

            For Each val As Integer In values
                If val = Convert.ToInt32(e) Then
                    Dim memInfo = type.GetMember(type.GetEnumName(val))
                    Dim descriptionAttribute As DescriptionAttribute = memInfo(0).GetCustomAttributes((New DescriptionAttribute).GetType(), False).FirstOrDefault()

                    If descriptionAttribute IsNot Nothing Then
                        Return descriptionAttribute.Description
                    End If
                End If
            Next
        End If

        Return String.Empty
    End Function
End Module
