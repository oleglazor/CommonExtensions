Imports System.Configuration
Imports Common.Extensions

<ConfigurationCollection(GetType(Sender), AddItemName:="add", CollectionType:=ConfigurationElementCollectionType.BasicMap)>
Public Class SendersCollection
    Inherits ConfigurationCollectionBase(Of Sender)

    Protected Overrides Function GetElementKey(ByVal element As ConfigurationElement) As Object
        Return (CType(element, Sender)).Email
    End Function

    Public Overloads Sub Remove(ByVal serviceConfig As Sender)
        BaseRemove(serviceConfig.Email)
    End Sub

End Class