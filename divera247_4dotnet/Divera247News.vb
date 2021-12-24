Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Mail
Public Class Divera247News

    Public number As String = ""
    Public title As String = ""
    Public text As String = ""
    Public address As String = ""
    Public ric As String = ""
    Public person As String = ""
    Public smtp_to As String = ""

    Public Function _ToMAIL()
        Dim xMAIL As String = ""
        xMAIL = xMAIL & "##### Neue Mitteilung #####" & vbCrLf
        xMAIL = xMAIL & "" & vbCrLf
        xMAIL = xMAIL & "Einsatz-Nummer" & ": " & number & vbCrLf
        xMAIL = xMAIL & "Stichwort" & ": " & title & vbCrLf
        xMAIL = xMAIL & "Meldung" & ": " & text & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        xMAIL = xMAIL & "Adresse" & ": " & address & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        xMAIL = xMAIL & "RIC" & ": " & ric & vbCrLf
        xMAIL = xMAIL & "Personen" & ": " & person & vbCrLf
        xMAIL = xMAIL & "SMTP" & ": " & smtp_to & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        Return xMAIL
    End Function

    Public Function _ToJSON()
        Dim xJSON As String = "{" & vbCrLf
        xJSON = xJSON & """" & "number" & """" & ": " & """" & number & """" & "," & vbCrLf
        xJSON = xJSON & """" & "title" & """" & ": " & """" & title & """" & "," & vbCrLf
        xJSON = xJSON & """" & "text" & """" & ": " & """" & text & """" & "," & vbCrLf
        xJSON = xJSON & """" & "address" & """" & ": " & """" & address & """" & "," & vbCrLf
        xJSON = xJSON & """" & "ric" & """" & ": " & """" & ric & """" & "," & vbCrLf
        xJSON = xJSON & """" & "person" & """" & ": " & """" & person & """" & "" & vbCrLf
        xJSON = xJSON & "}"
        Return xJSON
    End Function

End Class
