<%
	Function GetRemoteFile(strURL)
		Dim objXML, strContents, arrLines
		Set objXML=Server.CreateObject("Microsoft.XMLHTTP")
		
		'read text file...
		objXML.Open "GET", strURL, False
		objXML.Send
		GetRemoteFile = objXML.ResponseText
		Set objXML=Nothing
	End Function
	
			sub sendEmail(mailfrom,mailto,mailsubject,mailbody)
			emailto=mailto
			set config = CreateObject("CDO.Configuration")
			sch = "http://schemas.microsoft.com/cdo/configuration/"
			with config.Fields
			.item(sch & "sendusing") = 2 ' cdoSendUsingPort
			.item(sch & "smtpserver") = "smtp.mandrillapp.com"
			.item(sch & "smtpserverport") = 587
			.item(sch & "smtpauthenticate") = 1 'basic auth
			.item(sch & "sendusername") = "cto@onedaycart.com"
			.item(sch & "sendpassword") = "h-7GBNaJdEW4ZMgJw-T02w"
			.update
			end with
			
					with CreateObject("CDO.Message")
					
					.configuration = config
					.to = mailto
					.from = mailfrom
					.subject = mailsubject
					.HTMLBody = "<img src=http://i.imgur.com/xDsq6Sa.png><hr>" + mailbody + "</hr> <b>www.onedaycart.com</b>"
					.send()
					end with
					
			end sub
			
			sub sendHTMLEmail(mailfrom,mailto,mailsubject,mailurl)
			emailto=request.querystring("emailto")
			set config = CreateObject("CDO.Configuration")
			sch = "http://schemas.microsoft.com/cdo/configuration/"
			with config.Fields
			.item(sch & "sendusing") = 2 ' cdoSendUsingPort
			.item(sch & "smtpserver") = "smtp.mandrillapp.com"
			.item(sch & "smtpserverport") = 587
			.item(sch & "smtpauthenticate") = 1 'basic auth
			.item(sch & "sendusername") = "cto@onedaycart.com"
			.item(sch & "sendpassword") = "h-7GBNaJdEW4ZMgJw-T02w"
			.update
			end with
						with CreateObject("CDO.Message")
						
						.configuration = config
						.to = mailto
						.from = mailfrom
						.subject = mailsubject
						.HTMLBody  = GetRemoteFile(mailurl)
						.send()
						end with
			end sub

%>