Imports Microsoft.Win32
Public Class Divera247Registry
    Public Function _RegistryCreateKey(ByVal _KeyStore As String, ByVal _ParentKey As String, ByVal _NewKey As String)

        Dim _RegistryKey As RegistryKey
        Try
            Select Case _KeyStore.ToUpper
                Case Is = "CURRENTUSER32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()

                Case Is = "CURRENTUSER64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey = _RegistryKey.CreateSubKey(_NewKey, True)
                    _RegistryKey.Close()
            End Select
            Return True
        Catch ex As Exception

            Return False
        End Try

    End Function

    Public Function _RegistryDeleteKey(ByVal _KeyStore As String, ByVal _ParentKey As String, ByVal _KeyToDelete As String)

        Dim _RegistryKey As RegistryKey
        Try
            Select Case _KeyStore.ToUpper
                Case Is = "CURRENTUSER32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()

                Case Is = "CURRENTUSER64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteSubKeyTree(_KeyToDelete, True)
                    _RegistryKey.Close()
            End Select

            Return True
        Catch ex As Exception

            Return False
        End Try


    End Function



    Public Function _InternalRegistryGetValue(ByVal _KeyStore As String, ByVal _ParentKey As String, ByVal _KeyName As String)
        Dim _RegistryKey As RegistryKey
        Try
            Select Case _KeyStore.ToUpper
                Case Is = "CURRENTUSER32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
                Case Is = "CLASSESROOT32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
                Case Is = "LOCALMACHINE32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
                Case Is = "CURRENTCONFIG32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True

                Case Is = "CURRENTUSER64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
                Case Is = "CLASSESROOT64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
                Case Is = "LOCALMACHINE64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
                Case Is = "CURRENTCONFIG64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    Return _RegistryKey.GetValue(_KeyName, "")
                    _RegistryKey.Close()
                    Return True
            End Select
            Return False
        Catch ex As Exception
            Return False
        End Try


    End Function

    Public Function _RegistrySetValue(ByVal _KeyStore As String, ByVal _ParentKey As String, ByVal _KeyName As String, ByVal _KeyValue As String, ByVal _KeyType As String)
        Dim _RegistryKey As RegistryKey
        Try
            Select Case _KeyStore.ToUpper
                Case Is = "CURRENTUSER32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()

                Case Is = "CURRENTUSER64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.SetValue(_KeyName, _KeyValue, CInt(_KeyType))
                    _RegistryKey.Close()
            End Select
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function _RegistryDeleteValue(ByVal _KeyStore As String, ByVal _ParentKey As String, ByVal _KeyName As String)
        Dim _RegistryKey As RegistryKey
        Try
            Select Case _KeyStore.ToUpper
                Case Is = "CURRENTUSER32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()

                Case Is = "CURRENTUSER64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
                Case Is = "CLASSESROOT64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
                Case Is = "LOCALMACHINE64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
                Case Is = "CURRENTCONFIG64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    _RegistryKey.DeleteValue(_KeyName, False)
                    _RegistryKey.Close()
            End Select
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function



    Public Function _RegistryValueExists(ByVal _KeyStore As String, ByVal _ParentKey As String, ByVal _ValueName As String)
        Dim _RegistryKey As RegistryKey
        Try
            Select Case _KeyStore.ToUpper
                Case Is = "CURRENTUSER32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
                Case Is = "CLASSESROOT32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
                Case Is = "LOCALMACHINE32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
                Case Is = "CURRENTCONFIG32"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry32).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If

                Case Is = "CURRENTUSER64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
                Case Is = "CLASSESROOT64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
                Case Is = "LOCALMACHINE64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
                Case Is = "CURRENTCONFIG64"
                    _RegistryKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentConfig, Microsoft.Win32.RegistryView.Registry64).OpenSubKey(_ParentKey, True)
                    If _RegistryKey.GetValue(_ValueName, "~NULL~") <> "~NULL~" Then
                        Return True
                    End If
            End Select
            Return False
        Catch ex As Exception
            Return False
        End Try

    End Function
End Class
