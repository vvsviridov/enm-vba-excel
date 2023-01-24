VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnmRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private authEndpoint As String
Private exitEndpoint As String
Private applEndpoint As String
Private blnAsync as Boolean

Private oServerXmlHttp as Object

''' PRIVATE FUNCTIONS '''

Private Sub Class_Initialize()
  Set oServerXmlHttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
  blnAsync = False
End Sub


''' PUBLIC FUNCTIONS '''

Public Sub login(ByVal baseUrl, ByVal login, ByVal password As String)
  authEndpoint = baseUrl & "/login?IDToken1=" & login & "&IDToken2=" & WorksheetFunction.EncodeURL(password)
  exitEndpoint = baseUrl & "/logout"
  applEndpoint = baseUrl & "/configuration-tasks/v1/tasks"
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


Public Function execute(ByVal payLoad as String)
  With oServerXmlHttp
    .Open "POST", applEndpoint, blnAsync
    .setRequestHeader "Content-Type", "application/json"
    .Send(payLoad)
    While oServerXmlHttp.readyState <> 4
        DoEvents
    Wend
    execute = .responsetext
  End With 
End Function
