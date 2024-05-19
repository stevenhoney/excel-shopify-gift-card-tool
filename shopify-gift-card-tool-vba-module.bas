Attribute VB_Name = "Module1"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Sub vbaGenerateCodes()
    Dim httpObject As New MSXML2.ServerXMLHTTP60
'    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    Dim sURL As String
    Dim jsonBody As String
    Dim sgetresult As String
    Dim finResult As String
    
    Range("G3:G100000").Clear
    Range("H3:H100000").Clear
    Range("I3:I100000").Clear
    
    If Range("APIKey").Value = "" Or Range("endpoint").Value = "" Or Range("storeURL").Value = "" Then
        MsgBox "Please set up API credentials", vbInformation, "Shopify Gift Card Tool"
        Exit Sub
    End If
    
    If MsgBox("Are you sure? This cannot be reversed, gift cards will be immediately sent if assigned to a customer.", vbYesNo + vbInformation, "Shopify Gift Card Tool") = vbNo Then
        Exit Sub
    End If
    
    Dim r As Integer
    r = 3
    
    sURL = Range("APILink").Value
    
    Do While Trim(Cells(r, 1)) <> ""
    

        jsonBody = Replace(Replace(Replace(Range("jsonBody").Value, "{{note}}", IIf(Cells(r, 2) <> "", Cells(r, 2), "null")), "{{initial_value}}", IIf(Cells(r, 1) <> "", Cells(r, 1), "null")), "{{code}}", IIf(Cells(r, 3) <> "", Cells(r, 3), "null"))
        jsonBody = Replace(Replace(Replace(jsonBody, "{{template_suffix}}", IIf(Cells(r, 4) <> "", Cells(r, 4), "null")), "{{customer_id}}", IIf(Cells(r, 5) <> "", Cells(r, 5), "null")), "{{expires_on}}", IIf(Cells(r, 6) <> "", Cells(r, 6), "null"))
        jsonBody = Replace(Replace(jsonBody, """null""", "null"), Chr(10), " \n ")

          
        httpObject.Open "POST", sURL, False
                
        httpObject.setRequestHeader "Content-Type", "application/json"
        httpObject.setRequestHeader "X-Shopify-Access-Token", Range("API_Token").Value
        httpObject.Send (Replace(jsonBody, """null""", "null"))
        
        sgetresult = httpObject.responsetext
        
        finResult = Replace(Replace(Replace(Replace(sgetresult, "{", ""), "}", ""), "[", ""), "]", "")
        
        Cells(r, 9) = sgetresult
        
        If InStr(1, finResult, "errors") > 0 Then
            finResult = Replace(finResult, """", "")
        Else
            finResult = Replace(Range("gcLink"), "{{gift_card}}", Mid(sgetresult, InStr(1, sgetresult, "id") + 4, 12))
        End If
        
        Cells(r, 7) = finResult
        
        If Left(finResult, 5) = "https" Then
            Range("G" & r).Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=Range("G" & r).Value, TextToDisplay:=Range("G" & r).Value
        End If
        
        Sleep (65)
        
        r = r + 1
    Loop
    
    
End Sub
