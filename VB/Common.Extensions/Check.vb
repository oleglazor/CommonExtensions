Imports System.IO
Imports System.Text.RegularExpressions

Public Module Check
    Sub ThrowIfNull(ByVal arg As Object, ByVal Optional messageOverride As String = "")
        If arg Is Nothing Then
            Throw New ArgumentNullException(messageOverride)
        End If
    End Sub

    Function IfNull(ByVal arg As Object) As Boolean
        Return arg Is Nothing
    End Function

    Sub ThrowIfAnyIsFalse(ParamArray checks As Func(Of Boolean)())
        checks.ToList().ForEach(Sub(check)
                                    If Not check() Then
                                        Throw New InvalidDataException($"Invalid data or condition was meet.")
                                    End If
                                End Sub)
    End Sub

    Function IsNullOrEmpty(ByVal arg As Object) As Boolean
        If arg Is Nothing Then
            Return True
        End If

        If TypeOf arg Is String Then

            If String.IsNullOrEmpty((TryCast(arg, String)).Trim()) Then
                Return True
            End If
        End If

        If arg.[GetType]().GetInterface(NameOf(ICollection)) IsNot Nothing Then

            If (TryCast(arg, ICollection)).Count = 0 Then
                Return True
            End If
        End If

        Return False
    End Function

    Sub ThrowIfNullOrEmpty(ByVal arg As Object, ByVal Optional argName As String = "", ByVal Optional messageOverride As String = "")
        If IsNullOrEmpty(arg) Then
            Throw New ArgumentNullException(argName, messageOverride)
        End If
    End Sub

    Sub ThrowIfAllIsNullOrEmpty(Of T)(ByVal array As T(), ByVal Optional messageOverride As String = "")
        If array Is Nothing Then
            Throw New ArgumentNullException(GetType(T).Name, messageOverride)
        End If

        If GetType(T) = GetType(String) AndAlso array.All(Function(a) String.IsNullOrEmpty(a?.ToString().Trim())) Then
            Throw New ArgumentNullException(GetType(T).Name, If(String.IsNullOrEmpty(messageOverride), "At least one of the argument must be not empty.", messageOverride))
        End If

        If array.All(Function(a) a Is Nothing) Then
            Throw New ArgumentNullException(GetType(T).Name, "At least one of the argument must be not null.")
        End If
    End Sub

    Function IfAllIsNullOrEmpty(Of T)(ByVal array As T()) As Boolean
        If array Is Nothing Then
            Return True
        End If

        If GetType(T) = GetType(String) AndAlso array.All(Function(a) String.IsNullOrEmpty(a?.ToString().Trim())) Then
            Return True
        End If

        If array.All(Function(a) a Is Nothing) Then
            Return True
        End If

        Return False
    End Function

    Sub ThrowIfAnyIsNullOrEmpty(Of T)(ByVal array As T(), ByVal Optional messageOverride As String = "")
        If array Is Nothing Then
            Throw New ArgumentNullException(GetType(T).Name, messageOverride)
        End If

        If GetType(T) = GetType(String) AndAlso array.Any(Function(a) String.IsNullOrEmpty(a?.ToString().Trim())) Then
            Throw New ArgumentNullException(GetType(T).Name, If(String.IsNullOrEmpty(messageOverride), "All arguments required.", messageOverride))
        End If

        If array.Any(Function(a) a Is Nothing) Then
            Throw New ArgumentNullException(GetType(T).Name, "All arguments required.")
        End If
    End Sub

    Function IfAnyIsNullOrEmpty(Of T)(ByVal array As T()) As Boolean
        If array Is Nothing Then
            Return True
        End If

        If GetType(T) = GetType(String) AndAlso array.Any(Function(a) String.IsNullOrEmpty(a?.ToString().Trim())) Then
            Return True
        End If

        If array.Any(Function(a) a Is Nothing) Then
            Return True
        End If

        Return False
    End Function

    Function IsInvalidId(ByVal id As Guid?) As Boolean
        Return Not id.HasValue OrElse id = Guid.Empty
    End Function

    Function IsInvalidId(ByVal id As Long?) As Boolean
        Return id <= 0
    End Function

    Sub ThrowIfInvalidId(ByVal id As Long?, ByVal Optional messageOverride As String = "")
        If IsInvalidId(id) Then
            Throw New ArgumentException(NameOf(id), messageOverride)
        End If
    End Sub

    Sub ThrowIfInvalidId(ByVal id As Guid, ByVal Optional messageOverride As String = "")
        If id = Guid.Empty Then
            Throw New ArgumentNullException(NameOf(id), messageOverride)
        End If
    End Sub

    Sub ThrowIfOutOfRange(ByVal value As Short, ByVal min As Short, ByVal max As Short, ByVal Optional messageOverride As String = "")
        ThrowIfOutOfRange(CLng(value), CLng(min), CLng(max), messageOverride)
    End Sub

    Sub ThrowIfOutOfRange(ByVal value As Integer, ByVal min As Integer, ByVal max As Integer, ByVal Optional messageOverride As String = "")
        ThrowIfOutOfRange(CLng(value), CLng(min), CLng(max), messageOverride)
    End Sub

    Sub ThrowIfOutOfRange(ByVal value As Long, ByVal min As Long, ByVal max As Long, ByVal Optional messageOverride As String = "")
        If value < min OrElse value > max Then
            Throw New ArgumentOutOfRangeException(NameOf(value), messageOverride)
        End If
    End Sub

    Function IsInRange(ByVal value As Short, ByVal min As Short, ByVal max As Short, ByVal Optional allowOnLimit As Boolean = True, ByVal Optional messageOverride As String = "") As Boolean
        Return IsInRange(CLng(value), CLng(min), CLng(max), allowOnLimit, messageOverride)
    End Function

    Function IsInRange(ByVal value As Integer, ByVal min As Integer, ByVal max As Integer, ByVal Optional allowOnLimit As Boolean = True, ByVal Optional messageOverride As String = "") As Boolean
        Return IsInRange(CLng(value), CLng(min), CLng(max), allowOnLimit, messageOverride)
    End Function

    Function IsInRange(ByVal value As Long, ByVal min As Long, ByVal max As Long, ByVal Optional allowOnLimit As Boolean = True, ByVal Optional messageOverride As String = "") As Boolean
        If allowOnLimit Then
            Return value >= min AndAlso value <= max
        End If

        Return value > min AndAlso value < max
    End Function

    Function IsPhoneNumber(ByVal value As String) As Boolean
        Dim match = Regex.Match(value, "^\+?(\d[\d-. ]+)?(\([\d-. ]+\))?[\d-. ]+\d$", RegexOptions.IgnoreCase)
        Return match.Success
    End Function
End Module
