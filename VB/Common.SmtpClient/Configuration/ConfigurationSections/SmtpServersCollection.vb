Imports System.Configuration
Imports Common.Extensions

<ConfigurationCollection(GetType(SmtpServer), AddItemName:="add", CollectionType:=ConfigurationElementCollectionType.BasicMap)>
Public Class SmtpServersCollection
    Inherits ConfigurationCollectionBase(Of SmtpServer)

    Protected Overrides Function GetElementKey(ByVal element As ConfigurationElement) As Object
        Return (CType(element, SmtpServer)).Host
    End Function

    Public Overloads Sub Remove(ByVal serviceConfig As SmtpServer)
        BaseRemove(serviceConfig.Host)
    End Sub

End Class