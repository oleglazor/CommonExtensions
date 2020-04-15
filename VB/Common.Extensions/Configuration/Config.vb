Public Module Config
    Public Enum RecipientGroup
        ConfirmBack
        AnotherGroup
    End Enum

    Public Enum Senders
        DoNotReply
    End Enum

    Public Class AppSettings

        Public Const SourceDirectorySetting As String = "SourceDirectory"
        Public Const DestinationDirectorySetting As String = "DestinationDirectory"
        Public Const LogDirectorySetting As String = "LogDirectory"

    End Class

End Module
