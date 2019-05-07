Function getHtmlFromUrl(pURL As String) As String
     Dim resText As String
     Dim objHttp As Object
     Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
     objHttp.Open "GET", pURL, False
     objHttp.Send ""
     getHtmlFromUrl = Mid(objHttp.ResponseText, 1, 255)

End Function

'Sources:
'[01] http://www.tomasvasquez.com.br/forum/viewtopic.php?t=3463﻿﻿
'[02] https://www.mrexcel.com/forum/excel-questions/707305-excel-vba-check-if-certain-website-online.html
