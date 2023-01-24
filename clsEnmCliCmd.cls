VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnmCliCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private authEndpoint As String
Private exitEndpoint As String
Private applEndpoint As String
Private blnAsync as Boolean

Private oServerXmlHttp as Object

Private pTimeWait as Integer


' Properties
Property Get TimeWait() As Integer
  TimeWait = pTimeWait
End Property

Property Let TimeWait(value As Integer)
  pTimeWait = value
End Property


''' PRIVATE FUNCTIONS '''

Private Sub Class_Initialize()
  Set oServerXmlHttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
  pTimeWait = 10 / 86400
  blnAsync = False
End Sub


''' PUBLIC FUNCTIONS '''

Public Sub login(ByVal baseUrl, ByVal login, ByVal password As String)
  authEndpoint = baseUrl & "/login?IDToken1=" & login & "&IDToken2=" & WorksheetFunction.EncodeURL(password)
  exitEndpoint = baseUrl & "/logout"
  applEndpoint = baseUrl & "/script-engine/services/command/"
  With oServerXmlHttp
    .SetOption(2) = (.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
    .Open "POST", authEndpoint, blnAsync
    .setRequestHeader "Content-Type", "application/json"
    .Send
    While oServerXmlHttp.readyState <> 4
        DoEvents
    Wend
  End With
End Sub


Public Sub logout()
  With oServerXmlHttp
    .Open "GET", exitEndpoint, blnAsync
    .Send
    While oServerXmlHttp.readyState <> 4
        DoEvents
    Wend
  End With
End Sub


Public Function execute(ByVal cmd as String)
  With oServerXmlHttp
    .Open "POST", applEndpoint, blnAsync

    sBoundary = "----WebKitFormBoundary" & RandomString(16)

    sPayLoad = sPayLoad & "--" & sBoundary & vbCrLf
    sPayLoad = sPayLoad & "Content-Disposition: form-data; name=""command""" & vbCrLf & vbCrLf
    sPayLoad = sPayLoad & cmd & vbCrLf
    sPayLoad = sPayLoad & "--" & sBoundary & "--" & vbCrLf

    .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & sBoundary
    .setRequestHeader "Content-Length", LenB(sPayLoad)

    .Send (sPayLoad)

    While oServerXmlHttp.readyState <> 4
        DoEvents
    Wend

    process_id = .getResponseHeader("process_id")
    procUrl = applEndpoint & "output/" & process_id & "?max_size=20000"
    
    Application.Wait (Now + timeWait)
    
    .Open "GET", procUrl, blnAsync
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .setRequestHeader "Accept", "text/plain; charset=UTF-8" '"application/json, text/javascript, */*; q=0.01"
    .Send ("Command=" & cmd)
    
    While oServerXmlHttp.readyState <> 4
        DoEvents
    Wend

    execute = .responsetext
  End With 
End Function


''' HELPER FUNCTIONS '''

Private Function RandomString(Length As Integer)
  'PURPOSE: Create a Randomized String of Characters
  'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

  Dim CharacterBank As Variant
  Dim x As Long
  Dim str As String

  'Test Length Input
    If Length < 1 Then
      MsgBox "Length variable must be greater than 0"
      Exit Function
    End If

  CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
    "k", "l", "m", "n", "o", "p", "q", "r","s", "t", "u", "v", "w", "x", _
    "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", _
    "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", _
    "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    

  'Randomly Select Characters One-by-One
    For x = 1 To Length
      Randomize
      str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next x

  'Output Randomly Generated String
  RandomString = str

End Function
