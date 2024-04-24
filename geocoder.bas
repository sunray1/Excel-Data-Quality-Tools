Attribute VB_Name = "geocoder"
Option Explicit

' Domain and URL for Google API
Public Const gstrGeocodingDomain = "https://maps.googleapis.com"
Public Const gstrGeocodingURL = "/maps/api/geocode/xml?"

' Google API Key see https://developers.google.com/maps/documentation/geocoding/get-api-key)
Public Const gstrKey = ""

' kludge to not overdo the API calls and add a delay
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Public Function URLEncode(StringToEncode As String, Optional _
   UsePlusRatherThanHexForSpace As Boolean = False) As String

Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
  Select Case asc(Mid(StringToEncode, CurChr, 1))
    Case 48 To 57, 65 To 90, 97 To 122
      TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
    Case 32
      If UsePlusRatherThanHexForSpace = True Then
        TempAns = TempAns & "+"
      Else
        TempAns = TempAns & "%" & Hex(32)
      End If
   Case Else
         TempAns = TempAns & "%" & _
              Format(Hex(asc(Mid(StringToEncode, _
              CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function

Public Function CallGeocodingAPI(lat As String, lng As String) As String
    ' Define XML and HTTP components
    Dim googleService As New MSXML2.XMLHTTP60
    Dim strLat As String
    Dim strLng As String
    Dim strQuery As String

    ' URL encode latitude and longitude
    strLat = URLEncode(lat)
    strLng = URLEncode(lng)

    ' Assemble the query string
    strQuery = gstrGeocodingURL
    strQuery = strQuery & "latlng=" & strLat & "," & strLng
    strQuery = strQuery & "&key=" & gstrKey

    ' Sleep to avoid rate limiting
    Sleep (5)

    ' Create HTTP request to query URL
    googleService.Open "GET", gstrGeocodingDomain & strQuery, False
    googleService.send

    ' Return response text
    CallGeocodingAPI = googleService.responseText
End Function

Public Function GetCountry(lat As String, lng As String) As String
    ' Call the API function to get response
    Dim responseText As String
    responseText = CallGeocodingAPI(lat, lng)

    ' Define XML component
    Dim googleResult As New MSXML2.DOMDocument60
    googleResult.LoadXML (responseText)

    ' Get the formatted address from the response
    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = googleResult.SelectSingleNode("//address_component[type='country']/long_name")

    If Not oNode Is Nothing Then
        GetCountry = oNode.Text
    Else
        GetCountry = ""
    End If
End Function

Public Function GetCountryCode(lat As String, lng As String) As String
    ' Call the API function to get response
    Dim responseText As String
    responseText = CallGeocodingAPI(lat, lng)

    ' Define XML component
    Dim googleResult As New MSXML2.DOMDocument60
    googleResult.LoadXML (responseText)

    ' Get the formatted address from the response
    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = googleResult.SelectSingleNode("//address_component[type='country']/short_name")

    If Not oNode Is Nothing Then
        GetCountryCode = oNode.Text
    Else
        GetCountryCode = ""
    End If
End Function

Public Function GetStateProvince(lat As String, lng As String) As String
    ' Call the API function to get response
    Dim responseText As String
    responseText = CallGeocodingAPI(lat, lng)

    ' Define XML component
    Dim googleResult As New MSXML2.DOMDocument60
    googleResult.LoadXML (responseText)

    ' Get the formatted address from the response
    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = googleResult.SelectSingleNode("//address_component[type='administrative_area_level_1']/long_name")

    If Not oNode Is Nothing Then
        GetStateProvince = oNode.Text
    Else
        GetStateProvince = ""
    End If
End Function

Public Function GetCounty(lat As String, lng As String) As String
    ' Call the API function to get response
    Dim responseText As String
    responseText = CallGeocodingAPI(lat, lng)

    ' Define XML component
    Dim googleResult As New MSXML2.DOMDocument60
    googleResult.LoadXML (responseText)

    ' Get the formatted address from the response
    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = googleResult.SelectSingleNode("//address_component[type='administrative_area_level_2']/long_name")

    If Not oNode Is Nothing Then
        GetCounty = oNode.Text
    Else
        GetCounty = ""
    End If
End Function

Public Function GetMunicipality(lat As String, lng As String) As String
    ' Call the API function to get response
    Dim responseText As String
    responseText = CallGeocodingAPI(lat, lng)

    ' Define XML component
    Dim googleResult As New MSXML2.DOMDocument60
    googleResult.LoadXML (responseText)

    ' Get the formatted address from the response
    Dim oNode As MSXML2.IXMLDOMNode
    Set oNode = googleResult.SelectSingleNode("//address_component[type='locality']/long_name")

    If Not oNode Is Nothing Then
        GetMunicipality = oNode.Text
    Else
        GetMunicipality = ""
    End If
End Function
