Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Mail
Public Class Divera247Alarm

    Public number As String = ""
    Public title As String = ""
    Public priority As Boolean = False
    Public text As String = ""
    Public address As String = ""
    Public lat As String = ""
    Public lng As String = ""
    Public scene_object As String = ""
    Public caller As String = ""
    Public patient As String = ""
    Public remark As String = ""
    Public units As String = ""
    Public destination As Boolean = False
    Public destination_address As String = ""
    Public destination_lat As String = ""
    Public destination_lng As String = ""
    Public additional_text_1 As String = ""
    Public additional_text_2 As String = ""
    Public additional_text_3 As String = ""
    Public ric As String = ""
    Public vehicle As String = ""
    Public person As String = ""
    Public delay As String = ""
    Public coordinates As String = ""
    Public coordinates_x As String = "" 'coordinates-x
    Public coordinates_y As String = "" 'coordinates-y
    Public news As Boolean = False
    Public id As String = ""
    Public cluster_ids As String = ""
    Public group_ids As String = ""
    Public user_cluster_relation_ids As String = ""
    Public vehicle_ids As String = ""
    Public alarmcode_id As String = ""

    Public smtp_to As String = ""

    Public Function _ToMAIL()
        Dim xMAIL As String = ""
        xMAIL = xMAIL & "##### Neuer Alarm #####" & vbCrLf
        xMAIL = xMAIL & "" & vbCrLf
        xMAIL = xMAIL & "Einsatz-Nummer" & ": " & number & vbCrLf
        xMAIL = xMAIL & "Stichwort" & ": " & title & vbCrLf
        xMAIL = xMAIL & "Meldung" & ": " & text & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        xMAIL = xMAIL & "Melder" & ": " & caller & vbCrLf
        xMAIL = xMAIL & "Adresse" & ": " & address & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        xMAIL = xMAIL & "Zusatz-Text" & ": " & additional_text_1 & vbCrLf
        xMAIL = xMAIL & "Zusatz-Text" & ": " & additional_text_2 & vbCrLf
        xMAIL = xMAIL & "Zusatz-Text" & ": " & additional_text_3 & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        xMAIL = xMAIL & "Alarmvorlage" & ": " & alarmcode_id & vbCrLf
        xMAIL = xMAIL & "RIC" & ": " & ric & vbCrLf
        xMAIL = xMAIL & "Fahrzeuge" & ": " & vehicle & vbCrLf
        xMAIL = xMAIL & "Personen" & ": " & person & vbCrLf
        xMAIL = xMAIL & "SMTP" & ": " & smtp_to & vbCrLf
        xMAIL = xMAIL & "------------------------------" & vbCrLf
        Return xMAIL
    End Function

    Public Function _ToJSON()
        Dim xJSON As String = "{" & vbCrLf
        xJSON = xJSON & """" & "number" & """" & ": " & """" & number & """" & "," & vbCrLf
        xJSON = xJSON & """" & "title" & """" & ": " & """" & title & """" & "," & vbCrLf
        xJSON = xJSON & """" & "priority" & """" & ": " & """" & priority.ToString & """" & "," & vbCrLf
        xJSON = xJSON & """" & "text" & """" & ": " & """" & text & """" & "," & vbCrLf
        xJSON = xJSON & """" & "address" & """" & ": " & """" & address & """" & "," & vbCrLf
        xJSON = xJSON & """" & "lat" & """" & ": " & """" & lat & """" & "," & vbCrLf
        xJSON = xJSON & """" & "lng" & """" & ": " & """" & lng & """" & "," & vbCrLf
        xJSON = xJSON & """" & "scene_object" & """" & ": " & """" & scene_object & """" & "," & vbCrLf
        xJSON = xJSON & """" & "caller" & """" & ": " & """" & caller & """" & "," & vbCrLf
        xJSON = xJSON & """" & "patient" & """" & ": " & """" & patient & """" & "," & vbCrLf
        xJSON = xJSON & """" & "remark" & """" & ": " & """" & remark & """" & "," & vbCrLf
        xJSON = xJSON & """" & "units" & """" & ": " & """" & units & """" & "," & vbCrLf
        xJSON = xJSON & """" & "destination" & """" & ": " & """" & destination.ToString & """" & "," & vbCrLf
        xJSON = xJSON & """" & "destination_address" & """" & ": " & """" & destination_address & """" & "," & vbCrLf
        xJSON = xJSON & """" & "destination_lat" & """" & ": " & """" & destination_lat & """" & "," & vbCrLf
        xJSON = xJSON & """" & "destination_lng" & """" & ": " & """" & destination_lng & """" & "," & vbCrLf
        xJSON = xJSON & """" & "additional_text_1" & """" & ": " & """" & additional_text_1 & """" & "," & vbCrLf
        xJSON = xJSON & """" & "additional_text_2" & """" & ": " & """" & additional_text_2 & """" & "," & vbCrLf
        xJSON = xJSON & """" & "additional_text_3" & """" & ": " & """" & additional_text_3 & """" & "," & vbCrLf
        xJSON = xJSON & """" & "ric" & """" & ": " & """" & ric & """" & "," & vbCrLf
        xJSON = xJSON & """" & "vehicle" & """" & ": " & """" & vehicle & """" & "," & vbCrLf
        xJSON = xJSON & """" & "person" & """" & ": " & """" & person & """" & "," & vbCrLf
        xJSON = xJSON & """" & "delay" & """" & ": " & """" & delay & """" & "," & vbCrLf
        xJSON = xJSON & """" & "coordinates" & """" & ": " & """" & coordinates & """" & "," & vbCrLf
        xJSON = xJSON & """" & "coordinates-x" & """" & ": " & """" & coordinates_x & """" & "," & vbCrLf
        xJSON = xJSON & """" & "coordinates-y" & """" & ": " & """" & coordinates_y & """" & "," & vbCrLf
        xJSON = xJSON & """" & "news" & """" & ": " & """" & news & """" & "," & vbCrLf
        xJSON = xJSON & """" & "id" & """" & ": " & """" & id & """" & "," & vbCrLf
        xJSON = xJSON & """" & "cluster_ids" & """" & ": " & """" & cluster_ids & """" & "," & vbCrLf
        xJSON = xJSON & """" & "group_ids" & """" & ": " & """" & group_ids & """" & "," & vbCrLf
        xJSON = xJSON & """" & "user_cluster_relation_ids" & """" & ": " & """" & user_cluster_relation_ids & """" & "," & vbCrLf
        xJSON = xJSON & """" & "vehicle_ids" & """" & ": " & """" & vehicle_ids & """" & "," & vbCrLf
        xJSON = xJSON & """" & "alarmcode_id" & """" & ": " & """" & alarmcode_id & """" & "" & vbCrLf
        xJSON = xJSON & "}"
        Return xJSON
    End Function

End Class
