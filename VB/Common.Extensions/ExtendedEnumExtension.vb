Imports System.Reflection
Imports System.Runtime.CompilerServices

<AttributeUsage(AttributeTargets.Field)>
Public Class ExtendedEnumItem
    Inherits Attribute

    Public ReadOnly Value As String
    Public ReadOnly Description As String

    Public Sub New(value As String, description As String)

        Me.Value = value
        Me.Description = description

    End Sub

End Class

Public Module ExtendedEnumExtension

    <Extension>
    Function Value(Of T As {Structure, IConvertible})(enumItem As T) As String

        Return GetExtendedEnumItemAttribute(enumItem).Value

    End Function

    <Extension>
    Function Description(Of T As {Structure, IConvertible})(enumItem As T) As String

        Return GetExtendedEnumItemAttribute(enumItem).Description

    End Function

    Private Function GetExtendedEnumItemAttribute(Of T As {Structure, IConvertible})(enumItem As T) As ExtendedEnumItem

        Dim fieldInfo As FieldInfo = enumItem.GetType().GetField(Convert.ToString(enumItem))
        If fieldInfo Is Nothing Then Return Nothing
        Dim attribute = CType(fieldInfo.GetCustomAttribute(GetType(ExtendedEnumItem)), ExtendedEnumItem)

        Return attribute

    End Function

End Module
