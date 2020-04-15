Imports System.Configuration
Imports Common.Extensions

<ConfigurationCollection(GetType(RecipientsCollection), AddItemName:="add", CollectionType:=ConfigurationElementCollectionType.BasicMap)>
Public Class RecipientsCollection
    Inherits ConfigurationCollectionBase(Of Recipient)

    Protected Overrides Function GetElementKey(ByVal element As ConfigurationElement) As Object
        Return (CType(element, Recipient)).Email
    End Function

    Public Overloads Sub Remove(ByVal serviceConfig As Recipient)
        BaseRemove(serviceConfig.Email)
    End Sub

End Class