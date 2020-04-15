Imports System.Runtime.Caching

Public Interface IMemoryCache
    Function [Get](Of T)(key As String) As T
    Function [Get](key As String) As Object
    Sub [Set](key As String, value As Object)
    Sub Remove(key As String)
End Interface

Public Class MemoryCacheSingleton
    Implements IMemoryCache

    Private Shared ReadOnly Cache As MemoryCache = MemoryCache.Default

    Public Function [Get](Of T)(key As String) As T Implements IMemoryCache.Get
        Return Cache.[Get](key)
    End Function

    Public Function [Get](key As String) As Object Implements IMemoryCache.Get
        Return Cache.[Get](key)
    End Function

    Public Sub [Set](key As String, value As Object) Implements IMemoryCache.Set
        Cache.AddOrGetExisting(key, value, DateTimeOffset.Now.AddHours(2))
    End Sub

    Public Sub Remove(key As String) Implements IMemoryCache.Remove
        Cache.Remove(key)
    End Sub

End Class
