Attribute VB_Name = "getCoLTaxonomy"
Option Explicit

' Domain and URL for Google API
Public Const gstrTaxonDomain = "https://api.checklistbank.org"
Public Const gstrTaxonURL = "/dataset/3/nameusage/search?content=SCIENTIFIC_NAME&limit=1&"

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
  Select Case Asc(Mid(StringToEncode, CurChr, 1))
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
              Format(Hex(Asc(Mid(StringToEncode, _
              CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function

Public Function CallTaxonAPI(sciName As String)
    ' Define XML and HTTP components
    Dim googleService As New MSXML2.XMLHTTP60
    Dim strSciName As String
    Dim strQuery As String

    ' URL encode sciName
    strSciName = URLEncode(sciName)

    ' Assemble the query string
    strQuery = gstrTaxonURL
    strQuery = strQuery & "q=" & strSciName

    ' Sleep to avoid rate limiting
    Sleep (5)

    ' Create HTTP request to query URL
    googleService.Open "GET", gstrTaxonDomain & strQuery, False
    googleService.send

    ' Return response text
    CallTaxonAPI = googleService.responseText
    

End Function

Public Function GetFamily(sciName As String)
    ' Call the API function to get response
    Dim responseText As String
    Dim jsonDic As Dictionary
    Dim item As Variant
    Dim resultArray As Object
    
    responseText = CallTaxonAPI(sciName)
    
    
    ' Convert JSON string to JSON object
    Set jsonDic = JsonConverter.ParseJson(responseText)
    
    If jsonDic("empty") = True Then
        GetFamily = "Taxa Not Found"
        Exit Function
    End If
    
    
    Set resultArray = jsonDic("result")(1)("classification")
    For Each item In resultArray
        If item("rank") = "family" Then
            GetFamily = item("name")
            Exit Function
        End If
    Next item
    
    GetFamily = "No Family Rank"

End Function

Public Function GetKingdom(sciName As String)
    ' Call the API function to get response
    Dim responseText As String
    Dim jsonDic As Dictionary
    Dim item As Variant
    Dim resultArray As Object
    
    responseText = CallTaxonAPI(sciName)
    
    
    ' Convert JSON string to JSON object
    Set jsonDic = JsonConverter.ParseJson(responseText)
    
    If jsonDic("empty") = True Then
        GetKingdom = "Taxa Not Found"
        Exit Function
    End If
    
    
    Set resultArray = jsonDic("result")(1)("classification")
    For Each item In resultArray
        If item("rank") = "kingdom" Then
            GetKingdom = item("name")
            Exit Function
        End If
    Next item
    
    GetKingdom = "No Kingdom Rank"

End Function

Public Function GetPhylum(sciName As String)
    ' Call the API function to get response
    Dim responseText As String
    Dim jsonDic As Dictionary
    Dim item As Variant
    Dim resultArray As Object
    
    responseText = CallTaxonAPI(sciName)
    
    
    ' Convert JSON string to JSON object
    Set jsonDic = JsonConverter.ParseJson(responseText)
    
    If jsonDic("empty") = True Then
        GetPhylum = "Taxa Not Found"
        Exit Function
    End If
    
    
    Set resultArray = jsonDic("result")(1)("classification")
    For Each item In resultArray
        If item("rank") = "phylum" Then
            GetPhylum = item("name")
            Exit Function
        End If
    Next item
    
    GetPhylum = "No Phylum Rank"

End Function

Public Function GetClass(sciName As String)
    ' Call the API function to get response
    Dim responseText As String
    Dim jsonDic As Dictionary
    Dim item As Variant
    Dim resultArray As Object
    
    responseText = CallTaxonAPI(sciName)
    
    
    ' Convert JSON string to JSON object
    Set jsonDic = JsonConverter.ParseJson(responseText)
    
    If jsonDic("empty") = True Then
        GetClass = "Taxa Not Found"
        Exit Function
    End If
    
    
    Set resultArray = jsonDic("result")(1)("classification")
    For Each item In resultArray
        If item("rank") = "class" Then
            GetClass = item("name")
            Exit Function
        End If
    Next item
    
    GetClass = "No Class Rank"

End Function

Public Function GetOrder(sciName As String)
    ' Call the API function to get response
    Dim responseText As String
    Dim jsonDic As Dictionary
    Dim item As Variant
    Dim resultArray As Object
    
    responseText = CallTaxonAPI(sciName)
    
    
    ' Convert JSON string to JSON object
    Set jsonDic = JsonConverter.ParseJson(responseText)
    
    If jsonDic("empty") = True Then
        GetOrder = "Taxa Not Found"
        Exit Function
    End If
    
    
    Set resultArray = jsonDic("result")(1)("classification")
    For Each item In resultArray
        If item("rank") = "order" Then
            GetOrder = item("name")
            Exit Function
        End If
    Next item
    
    GetOrder = "No Order Rank"

End Function
