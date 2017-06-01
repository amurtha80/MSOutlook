Sub PopulateTemplate()
    Set oInspector = Application.ActiveWindow
    Set myItem = oInspector.CurrentItem
    
    'Provide the email address(es) for the receipients and populate in the To field of email
    strEmailAddr = InputBox("Enter email address(es), separated by semicolon(;)")
    myItem.To = Replace(myItem.To, "[EmailAddr]", strEmailAddr)
    myItem.Display
    
    'Provide the project name for the email and populate in both the subject and body of email
    strProjName = InputBox("Enter the Project Name")
    myItem.Subject = Replace(myItem.Subject, "[ProjectName]", strProjName)
    myItem.HTMLBody = Replace(myItem.HTMLBody, "[ProjectName]", strProjName)
    myItem.Display
    
    'Provide the time of day for the greeting and populate in the body of email
    strTimeofDay = InputBox("Enter Morning, Afternoon, or Evening")
    myItem.HTMLBody = Replace(myItem.HTMLBody, "[TimeofDay]", strTimeofDay)
    myItem.Display
    
    'Provide the first name for the greeting and populate in body of email
    strContactName = InputBox("Enter the First Name for the Greeting")
    myItem.HTMLBody = Replace(myItem.HTMLBody, "[ContactName]", strContactName)
    myItem.Display
    
    ' Resolve all recipients (Same as pressing the "Check Names" button)
    Call myItem.Recipients.ResolveAll
    
    'Free memory
    Set strEmailAddr = Nothing
    Set strProjName = Nothing
    Set strTimeofDay = Nothing
    Set strContactName = Nothing
End Sub
