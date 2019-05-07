Sub IsItOnline()
    Debug.Print IsSiteOnline("http://www.tomasvasquez.com.br")
End Sub

Function IsSiteOnline(pURL As String) As Boolean
On Error GoTo GetErr
     Dim resText As String
     Dim objHttp As Object
     Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
     objHttp.Open "GET", pURL, False
     objHttp.Send ""
     IsSiteOnline = objHttp.Status = 200

SetOutput:
    Exit Function

GetErr:
    IsSiteOnline = False
    GoTo SetOutput

End Function
