Imports System.Configuration
Imports System.ServiceModel.Configuration

Public Enum SettingType
    AppSettings
    ConnectionString
    ConfigSection
    WcfServiceSection
End Enum

Public Interface IConfigurationProvider
    Function [Get](Of T)(key As String, ByVal Optional settingType As SettingType = SettingType.AppSettings, ByVal Optional defaultValue As T = Nothing) As T
End Interface

Public Class ConfigurationProvider
    Implements IConfigurationProvider

    Private ReadOnly _cache As IMemoryCache = New MemoryCacheSingleton()

    Public Function [Get](Of T)(ByVal key As String, ByVal Optional settingType As SettingType = SettingType.AppSettings, ByVal Optional defaultValue As T = Nothing) As T Implements IConfigurationProvider.Get
        Try
            Dim settingValue = Nothing
            Dim thatWasStruct = False

            Try
                settingValue = _cache.[Get](Of T)(key)
            Catch
                thatWasStruct = True
            End Try

            If settingValue IsNot Nothing AndAlso Not thatWasStruct Then
                Return settingValue
            End If

            Select Case settingType
                Case SettingType.AppSettings
                    Dim stringSetting = ConfigurationManager.AppSettings.[Get](key)

                    If GetType(T) = GetType(Long) Then
                        settingValue = CType(CObj(Long.Parse(stringSetting)), T)
                    ElseIf GetType(T) = GetType(Integer) Then
                        settingValue = CType(CObj(Integer.Parse(stringSetting)), T)
                    ElseIf GetType(T) = GetType(Short) Then
                        settingValue = CType(CObj(Short.Parse(stringSetting)), T)
                    ElseIf GetType(T) = GetType(Boolean) Then
                        settingValue = CType(CObj(Convert.ToBoolean(stringSetting)), T)
                    Else
                        settingValue = CType(CObj(stringSetting), T)
                    End If

                Case SettingType.ConnectionString
                    settingValue = CType(CObj(ConfigurationManager.ConnectionStrings(key).ConnectionString), T)
                Case SettingType.ConfigSection
                    settingValue = CType(ConfigurationManager.GetSection(key), T)
                Case SettingType.WcfServiceSection
                    Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
                    Dim section = TryCast(config.GetSectionGroup("system.serviceModel"), ServiceModelSectionGroup)
                    settingValue = CType(CObj(section?.Services?.Services(key)), T)
                Case Else
                    Throw New Exception($"Unsupported setting type: {settingType}.")
            End Select

            Check.ThrowIfNullOrEmpty(settingValue)
            _cache.[Set](key, settingValue)
            Return settingValue
        Catch ex As Exception

            If defaultValue IsNot Nothing Then
                Return defaultValue
            End If

            Throw New ConfigurationErrorsException($"Can't find setting value for key {key}", ex)
        End Try
    End Function
   
End Class
