<%@LANGUAGE="VBSCRIPT" %>
<!--#Include File="connection.asp"-->
<!-- #include file="adovbs.inc" -->

<%
 
strURL= Request.ServerVariables("URL")



' Create a RecordSet object
Set Rs = Server.CreateObject("ADODB.RecordSet")
 
' Open the table
Rs.Open "ArbCases_Tab", conn, adOpenKeySet, adLockPessimistic, adCmdTable

' Add a new record
Rs.AddNew
    Rs("FileNo") = Request.Form("fileno")
    Rs("AgmtNo") = Request.Form("agmtno")
	Rs("WorkName") = Request.Form("workname")
	Rs("Contractor") = Request.Form("contractor")
	Rs("Arbitrator") = Request.Form("arbitrator")
	Rs("Arb_Appointment") = Request.Form("arbapptdate")
	Rs("Award_Date") = Request.Form("arbawarddate")
	Rs("Claim_Amount") = Request.Form("claimamt")
	Rs("Award_Amount") = Request.Form("awardamt")
	Rs("Counter_Claim") = Request.Form("counteramt")
	Rs("Counter_Award") = Request.Form("counterawardamt")
	Rs("Divn_Office") = Request.Form("division")
	Rs("Divn_Incharge") = Request.Form("divnincharge")
	Rs("ArbCase_Status") = "Pending"
	

  ' Update the record
  Rs.Update

  ' Retrive the ID
  lUserID = Rs("ArbCaseID")

Response.Redirect "index.asp"  
' Close the RecordSet
Rs.Close
Set Rs = Nothing
conn.close							'Close the connection to the database
set conn = nothing					'Release the connection object from memory

%>

