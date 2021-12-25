Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Mail

Public Class Divera247Control

    Public _DIVERA247API_KEY = ""
    Public _SMTP_HOST As String = ""
    Public _SMTP_PORT As Integer = 25
    Public _SMTP_SSL As Boolean = False
    Public _SMTP_UseDefaultCredentials As Boolean = False
    Public _SMTP_CRED As New Net.NetworkCredential("", "")
    Public _SMTP_MailFrom As String = ""

    Public Function _Init()
        Dim RG As New Divera247Registry
        _SMTP_HOST = RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_HOST")
        _SMTP_PORT = RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_PORT")
        _SMTP_SSL = CBool(RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_SSL"))
        _SMTP_UseDefaultCredentials = CBool(RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_UseDefaultCredentials"))
        Dim CRED As New NetworkCredential(RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_USER").ToString, RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_PASS").ToString)
        _SMTP_CRED = CRED
        _SMTP_MailFrom = RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "SMTP_FROM")
        _DIVERA247API_KEY = RG._InternalRegistryGetValue("CURRENTUSER32", "DIVERA247", "APIKEY")
        Return True
    End Function

    Public Function _Init(ByVal _host As String, ByVal _port As Integer, ByVal _ssl As Boolean, ByVal _cred As Net.NetworkCredential, ByVal _mailfrom As String, ByVal _apikey As String)
        Dim RG As New Divera247Registry
        _SMTP_HOST = _host
        _SMTP_PORT = _port
        _SMTP_SSL = _ssl
        _SMTP_CRED = _cred
        _SMTP_MailFrom = _mailfrom
        _DIVERA247API_KEY = _apikey
        Return True
    End Function


    Public Function _WebApiDivera247Alarm(ByVal _AlarmParameter As String, Optional ByVal _UseJSON As Boolean = True)
        Try
            Dim PostData = _AlarmParameter
            Dim Request As WebRequest = WebRequest.Create("https://www.divera247.com/api/alarm?accesskey=" & _DIVERA247API_KEY)

            Request.Method = "POST"
            Dim ByteArray As Byte() = Encoding.UTF8.GetBytes(PostData)

            If _UseJSON = True Then
                Request.ContentType = "application/json"
            Else
                Request.ContentType = "application/x-www-form-urlencoded"
            End If
            Request.ContentLength = ByteArray.Length
            Dim DataStream As Stream = Request.GetRequestStream()
            DataStream.Write(ByteArray, 0, ByteArray.Length)
            DataStream.Close()
            Dim Response As WebResponse = Request.GetResponse()
            DataStream = Response.GetResponseStream()
            Dim Reader As New StreamReader(DataStream)
            Dim ResponseFromServer As String = Reader.ReadToEnd()
            Reader.Close()
            DataStream.Close()
            Response.Close()
            Return ResponseFromServer
        Catch ex As Exception
            Return False
        End Try


    End Function

    Public Function _WebApiDivera247News(ByVal _NewsParameter As String, Optional ByVal _UseJSON As Boolean = True)
        Try
            Dim PostData = _NewsParameter
            Dim Request As WebRequest = WebRequest.Create("https://www.divera247.com/api/news?accesskey=" & _DIVERA247API_KEY)

            Request.Method = "POST"
            Dim ByteArray As Byte() = Encoding.UTF8.GetBytes(PostData)
            If _UseJSON = True Then
                Request.ContentType = "application/json"
            Else
                Request.ContentType = "application/x-www-form-urlencoded"
            End If
            Request.ContentLength = ByteArray.Length
            Dim DataStream As Stream = Request.GetRequestStream()
            DataStream.Write(ByteArray, 0, ByteArray.Length)
            DataStream.Close()
            Dim Response As WebResponse = Request.GetResponse()
            DataStream = Response.GetResponseStream()
            Dim Reader As New StreamReader(DataStream)
            Dim ResponseFromServer As String = Reader.ReadToEnd()
            Reader.Close()
            DataStream.Close()
            Response.Close()
            Return ResponseFromServer
        Catch ex As Exception
            Return False
        End Try


    End Function

    Public Function SendAlarm(ByVal _AlarmClass As Divera247Alarm)
        _WebApiDivera247Alarm(_AlarmClass._ToJSON)

        If _AlarmClass.smtp_to <> "" Then
            _SendMail(_AlarmClass.smtp_to, _AlarmClass.title, _AlarmClass._ToMAIL)
        End If

        Return True
    End Function


    Public Function SendNews(ByVal _NewsClass As Divera247News)
        _WebApiDivera247News(_NewsClass._ToJSON)

        If _NewsClass.smtp_to <> "" Then
            _SendMail(_NewsClass.smtp_to, _NewsClass.title, _NewsClass._ToMAIL)
        End If

        Return True
    End Function


    Public Function _SendMail(ByVal _xTo As String, ByVal _xSubject As String, ByVal _xBody As String, Optional ByVal xAttachments() As String = Nothing)
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = _SMTP_UseDefaultCredentials
            Smtp_Server.Credentials = _SMTP_CRED
            Smtp_Server.Port = _SMTP_PORT
            Smtp_Server.EnableSsl = _SMTP_SSL
            Smtp_Server.Host = _SMTP_HOST

            e_mail = New MailMessage()
            e_mail.From = New MailAddress(_SMTP_MailFrom)
            e_mail.To.Add(_xTo)
            e_mail.Subject = _xSubject
            e_mail.IsBodyHtml = False
            e_mail.Body = _xBody

            If xAttachments IsNot Nothing Then
                For Each item In xAttachments
                    e_mail.Attachments.Add(New Attachment(item))
                Next
            End If

            Smtp_Server.Send(e_mail)
            Return True
        Catch error_t As Exception
            Return False
        End Try
    End Function

End Class
