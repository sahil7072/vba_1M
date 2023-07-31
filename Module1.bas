Attribute VB_Name = "Module1"
Sub gstgovin()
Dim k as long 
Dim driver As New ChromeDriver

With driver

        i = 2
        'New
        
            .Get "https://services.gst.gov.in/services/searchtp"
            Do While Sheet1.Cells(i, "A") <> ""
            
        .FindElementById("for_gstin").SendKeys (Sheet1.Cells(i, "A"))
        .FindElementById("fo-captcha").SendKeys ""
        .FindElementByCss("#lottable > div.tbl-format > div:nth-child(3) > div > div:nth-child(3) > p.wordCls", timeout:=25000).WaitDisplayed
        
         
         'WebDriverWait wait = new WebDriverWait(driver, 15)
         'driver Wait.Until(ExpectedConditions.elementToBeClickable(By.ID("lotsearch")))
        
        'Application.Wait Now() + TimeValue("00:00:10")
                       
        Legal_Name_of_Business = .FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div/form/div[5]/div/div[2]/div[1]/div/div[1]/p[2]").Text
        Trade_Name = .FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div/form/div[5]/div/div[2]/div[1]/div/div[2]/p[2]").Text
        Registered_on = .FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div/form/div[5]/div/div[2]/div[1]/div/div[3]").Text
        GSTIN_Status = .FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div/form/div[5]/div/div[2]/div[2]/div/div[2]/p[2]").Text
        'On Error Resume Next

          If .FindElementsByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div[1]/form/div[5]/div/div[2]/div[2]/div/div[2]/p[3]").Count > 0 Then
            Date_Of_Cancellation = .FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div[1]/form/div[5]/div/div[2]/div[2]/div/div[2]/p[3]").Text
            Sheet1.Cells(i, "F") = Date_Of_Cancellation
            
            'Debug.Print ("Yes")
            Else
            'Debug.Print ("No")
            End If

        Principal_Place_of_Business = .FindElementByXPath("/html/body/div[2]/div[2]/div/div[2]/div/div/form/div[5]/div/div[2]/div[3]/div/div[3]/p[2]").Text
        
        
        Sheet1.Cells(i, "B") = Legal_Name_of_Business
        Sheet1.Cells(i, "C") = Trade_Name
        Sheet1.Cells(i, "D") = Registered_on
        Sheet1.Cells(i, "E") = GSTIN_Status
        Sheet1.Cells(i, "G") = Principal_Place_of_Business
        


        
       i = i + 1
     Loop
     
    End With
    

End Sub
