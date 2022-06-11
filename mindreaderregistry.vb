Imports Microsoft.Win32
Module mindreaderregistry

    Sub setmrkey(ByVal folder As String, ByVal key As String, ByVal keyvalue As String)
        Dim regKey As RegistryKey
        regKey = Registry.CurrentUser.OpenSubKey("Software", True)
        regKey = regKey.CreateSubKey("ActivityOwner.Com", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
        regKey = regKey.CreateSubKey("mindreader2", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
        regKey = regKey.CreateSubKey(folder, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
        regKey.SetValue(key, keyvalue)
        regKey = Nothing
    End Sub
    Function getmrkey(ByVal folder As String, ByVal key As String) As String
        Dim regKey As RegistryKey
        Try
            regKey = Registry.CurrentUser.OpenSubKey("Software", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            regKey = regKey.OpenSubKey("ActivityOwner.Com", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            regKey = regKey.OpenSubKey("mindreader2", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            regKey = regKey.OpenSubKey(folder, Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree)
            getmrkey = regKey.GetValue(key)
        Catch
            getmrkey = ""
        End Try
    End Function

End Module
