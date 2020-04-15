Imports System.Configuration

Public MustInherit Class ConfigurationCollectionBase(Of T As ConfigurationElement)
    Inherits ConfigurationElementCollection

    Public Overrides Function IsReadOnly() As Boolean
        Return False
    End Function

    Default Public Overloads Property Item(ByVal index As Integer) As T
        Get
            Return CType(BaseGet(index), T)
        End Get
        Set(ByVal value As T)

            If BaseGet(index) IsNot Nothing Then
                BaseRemoveAt(index)
            End If

            LockItem = False
            BaseAdd(index, value)
        End Set
    End Property

    Public Sub Add(ByVal item As T)
        LockItem = False
        BaseAdd(item)
    End Sub

    Public Sub Clear()
        BaseClear()
    End Sub

    Protected Overrides Function CreateNewElement() As ConfigurationElement
        Return Activator.CreateInstance(Of T)()
    End Function

    Public Sub RemoveAt(ByVal index As Integer)
        BaseRemoveAt(index)
    End Sub

    Public Sub Remove(ByVal name As String)
        BaseRemove(name)
    End Sub

End Class