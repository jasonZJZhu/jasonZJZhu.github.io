<%
Set mail = createObject("CDO.Message")
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.gmail.com"
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "testing4759@gmail.com"
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "XBJCfrzFzRSB4GmzX$4E"
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
mail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
mail.Configuration.Fields.Update

mail.To = "testing4759@gmail.com"
mail.From = "testing4759@gmail.com"
mail.Subject = Request.Form("name") + Now()
mail.TextBody = Request.Form("comments")
mail.Send

Response.Write("Comments Recieved!")
Set mail = nothing
%>