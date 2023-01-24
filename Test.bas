Attribute VB_Name = "Test"


Public Sub CreateCliCmdObject()

    baseUrl = "https://enm.example.com"
    login = "login"
    password = "password"

    Dim oEnmCliCmd As clsEnmCliCmd
    Set oEnmCliCmd = New clsEnmCliCmd
	
    With oEnmCliCmd
        .login baseUrl, login, password
        cmd = "cmedit get NetworkElement=Node1"
        response = .execute(cmd)
        .logout
    End With
	
	ActiveSheet.Range("A1") = response
	
End Sub


Sub Rest_Execution()
    
    baseUrl = "https://enm.example.com"
    login = "login"
    password = "password"

    Dim oEnmRest As clsEnmRest
    Set oEnmRest = New clsEnmRest

    With oEnmRest
        .login baseUrl, login, password
        payLoad = "{""name"": ""readCells"",""fdn"": ""NetworkElement=BSC1""}"
        response = .execute(payLoad)
        .logout
    End With
	
	ActiveSheet.Range("A2") = response
    
End Sub
