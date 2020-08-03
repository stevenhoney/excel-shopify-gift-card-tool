Attribute VB_Name = "Module1"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Sub vbaGenerateCodes()
    Dim httpObject As New MSXML2.ServerXMLHTTP60
'    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    Dim sURL As String
    Dim jsonBody As String
    Dim sgetresult As String
    Dim finResult As String
    
    Range("AA:AA").Clear
    Range("F2:F2000").Clear
    
    If Range("APIKey").Value = "" Or Range("APIPassword").Value = "" Or Range("endpoint").Value = "" Or Range("storeURL").Value = "" Then
        MsgBox "Please set up API credentials", vbInformation, "Shopify Gift Card Tool"
        Exit Sub
    End If
    
    If MsgBox("Are you sure? This cannot be reversed.  Gift Cards assigned to customers will immediately be sent via email and/or text message.", vbYesNo + vbInformation, "Shopify Gift Card Tool") = vbNo Then
        Exit Sub
    End If
    
    Dim r As Integer
    r = 2
    
    sURL = Range("APILink").Value
    
    Do While Trim(Cells(r, 1)) <> ""
    

        jsonBody = Replace(Replace(Replace(Range("jsonBody").Value, "{{note}}", IIf(Cells(r, 2) <> "", Cells(r, 2), "null")), "{{initial_value}}", IIf(Cells(r, 1) <> "", Cells(r, 1), "null")), "{{code}}", IIf(Cells(r, 3) <> "", Cells(r, 3), "null"))
        jsonBody = Replace(Replace(jsonBody, "{{template_suffix}}", IIf(Cells(r, 4) <> "", Cells(r, 4), "null")), "{{customer_id}}", IIf(Cells(r, 5) <> "", Cells(r, 5), "null"))
        jsonBody = Replace(Replace(jsonBody, """null""", "null"), Chr(10), " \n ")

          
        httpObject.Open "POST", sURL, False
                
        httpObject.setRequestHeader "Authorization", "Basic " & Replace(EncodeBase64(Range("APIKey").Value + ":" + Range("APIPassword").Value), Chr(10), "")
        httpObject.setRequestHeader "Content-Type", "application/json"
        'httpObject.setRequestHeader "auth_token", Range("token").Value
        httpObject.Send (Replace(jsonBody, """null""", "null"))
        
        sgetresult = httpObject.responsetext
        
        finResult = Replace(Replace(Replace(Replace(sgetresult, "{", ""), "}", ""), "[", ""), "]", "")
        
        Cells(r, 27) = sgetresult
        
        If InStr(1, finResult, "errors") > 0 Then
            finResult = Replace(finResult, """", "")
        Else
            finResult = Replace(Range("gcLink"), "{{gift_card}}", Mid(sgetresult, InStr(1, sgetresult, "id") + 4, 12))
        End If
        
        Cells(r, 6) = finResult
        
        If Left(finResult, 5) = "https" Then
            Range("F" & r).Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=Range("F" & r).Value, TextToDisplay:=Range("F" & r).Value
        End If
        
        Sleep (500)
        
        r = r + 1
    Loop
    
    
End Sub

Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function
