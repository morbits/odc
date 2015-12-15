<!--#include file="email.asp"-->
<%
					Sub OpenDatabase(ByRef dbConn)
					    Set dbConn = Server.CreateObject("ADODB.Connection")
					
					    'Edit this Fields only
					    
					    		db_server="182.50.133.109"
					    		db="onedaycart"
					    		db_user="odc_beta1"
					    		db_password="odc098123"
					    
						dbConn.ConnectionTimeout = 150
						dbConn.CommandTimeout = 150
						dbConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & db_user & ";Password=" & db_password & ";Initial Catalog= "  & db  & ";Data Source=" & db_server
					 
					End Sub
					
					Function Query(sql)
						set con=Server.CreateObject("ADODB.Connection")
						set rs=Server.CreateObject("ADODB.Recordset")
						opendatabase con
			
						set rs=con.execute(sql)
						set Query=rs
					end Function
					
					
%>