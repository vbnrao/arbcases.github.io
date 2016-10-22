<%@LANGUAGE="VBSCRIPT"%>
<%Response.Buffer=True%>
<%on error resume next%>
<!--#Include File="connection.asp"-->
<!--#include file="System/md5.asp"-->

<!DOCTYPE html>
<!--[if IE 8]> <html lang="en" class="ie8"> <![endif]-->
<!--[if IE 9]> <html lang="en" class="ie9"> <![endif]-->
<!--[if !IE]><!--> <html lang="en"> <!--<![endif]-->

<!-- BEGIN HEAD -->
<head>
     <meta charset="UTF-8" />
    <title>IKP Knowledge Park Admin Dashboard | Login Page</title>
    <meta content="width=device-width, initial-scale=1.0" name="viewport" />
	<meta content="" name="description" />
	<meta content="" name="author" />
     <!--[if IE]>
        <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
        <![endif]-->
    <!-- GLOBAL STYLES -->
     <!-- PAGE LEVEL STYLES -->
     <link rel="stylesheet" href="assets/plugins/bootstrap/css/bootstrap.css" />
    <link rel="stylesheet" href="assets/css/login.css" />
    <link rel="stylesheet" href="assets/plugins/magic/magic.css" />
     <!-- END PAGE LEVEL STYLES -->
   <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.3.0/respond.min.js"></script>
    <![endif]-->
</head>
    <!-- END HEAD -->

<!-- Login COde goes here  -->

<%

dim strLogOff
strLogOff = Request.QueryString("msg")

   
   Dim Account_ID, User_Name, Password, Login_Attempts
   Dim rs, Query, msg, msg1, msgclr, bgclr
   bgclr = "#336699"
   msgclr = "#336699"
   msg = "Enter your user name and "
   msg1 = "Password"
   if request.Form("Login")="Sign in" then
      
      if request.form("username") = "" or request.form("password") = "" then
         msgclr = "Red"
		 msg = "You must enter User Name or Password !"
      else 
        
         User_Name = request.form("username")
         'Password = request.form("password")
         
         Query ="select User_Name, LastName, Account_Status,Login_Attempts,Allowed_Login_Attempts, Account_Level, Last_Login_Date, Last_Login_Time from USERS_TAB where User_Name='" & User_Name & "'"	
         'conn.Open
         set rs = conn.execute(Query)
         User_Name = rs("User_Name")     
            
         if err.number <> 0 then
            msgclr = "Red"
		    msg = "User Name or "
		    msg1 = "Password entered is incorrect!"      
            Call Close_All()
         else
            
            if lcase(rs("Account_Status")) = "disabled" then
               Call Close_All()
               msgclr = "Red"
               msg = "Your account has been locked!" 
               msg1 = "Please contact your Administrator!"    
            else           
                  
               if rs("Login_Attempts") >= rs("Allowed_Login_Attempts") then
                  rs.close
                  set rs = server.CreateObject("Adodb.Recordset")
                  rs.open "Select Account_Status,Login_Attempts From USERS_TAB Where User_Name='" & User_Name & "'",conn,2,3
                  rs("Account_Status") = "disabled"
                  rs("Login_Attempts") = 0
                  rs.update
                  Call Close_All()
                  msgclr = "Red"
                  msg = "Your account has been locked! "
                  msg1 = "Please contact your Administrator!"
               else
Query ="select Account_ID,User_Name,Password,LastName,Account_Level,dept_id, Login_Attempts,Last_Login_Date,Last_Login_Time from USERS_TAB where User_Name='" & User_Name & "'"
'conn.open
set rs = conn.execute(Query)
Account_ID = rs("Account_ID")
                        
                  if err.number <> 0 then
                     Call Close_All()
                     msgclr = "#FF0000"
		             msg = "User Name or Password entered is incorrect!"
		          else   
		             Password=MD5(Request.Form("password") & Account_ID)
 
		            'response.write(User_Name) 
		            'response.write("********testing*****")
		            'response.write(Password)
		             
		             if Password = rs("Password") then     
	                    
Response.Cookies("User_Login") = rs.Fields("User_Name")
Response.Cookies("Password") = rs.Fields("Password")
Response.Cookies("User_Name") = rs.Fields("LastName")
Response.Cookies("Account_Level") = rs.Fields("Account_Level")
Response.Cookies("dept_id") = rs.Fields("dept_id")
Response.Cookies("Last_Login_Date") = rs.Fields("Last_Login_Date")
Response.Cookies("Last_Login_Time") = rs.Fields("Last_Login_Time")


	                    
	                    rs.close
	                    set rs = server.CreateObject("Adodb.Recordset")
		                rs.open "Select Login_Attempts,Last_Login_Date,Last_Login_Time From USERS_TAB Where User_Name='" & User_Name & "'",conn,2,3	
	                    rs("Login_Attempts") = 0
                        rs("Last_Login_Date") = Date
                        rs("Last_Login_Time") = Time
                        rs.update
                        Call Close_All()
		response.Redirect "index.asp"
              
              
                     else
                        Login_Attempts = rs("Login_Attempts") + 1        
                        rs.close
                        set rs = server.CreateObject("Adodb.Recordset")
                        rs.open "Select Login_Attempts From USERS_TAB Where User_Name='" & User_Name & "'",conn,2,3
                        rs("Login_Attempts") = Login_Attempts
                        rs.update
		                Call Close_All() 
		                msgclr = "Yellow"
		                msg = "User Name or Password "
		                msg1 = "entered is incorrect!"                           	                   
                     end if
                  end if
               end if
            end if
         end if  
      end if
	end if
	
if strLogOff ="LogOff" then
'bgclr = "White"
msgclr= "Red"
msg = "You have been Successfully "
msg1 = "Logged out!"

Response.clear
'Session.Abandon()
end if


    
   Sub Close_All()
      rs.close
      conn.close
      rs = nothing
      conn = nothing
   End Sub


%>




<!-- Login Code Ends here  -->


    <!-- BEGIN BODY -->
<body >

   <!-- PAGE CONTENT --> 
    <div class="container">
    <div class="text-center">
        <img src="assets/img/logo.png" id="logoimg" alt=" Logo" />
    </div>
    <div class="tab-content">
        <div id="login" class="tab-pane active">
            <form action="Login.asp" method=post name="login" id="login" class="form-signin">
                <p class="text-muted text-center btn-block btn btn-primary btn-rect">
                    <%= msg %><%= msg1 %>
                </p>
                <input type="text" placeholder="Username" name="username" id="username" class="form-control" required />
                <input type="password" placeholder="Password" name="password" id="password" class="form-control" required />
                <input type="submit" id="submit" name="Login" value="Sign in" class="btn text-muted text-center btn-danger">
                <!--<button class="btn text-muted text-center btn-danger" type="submit" name="Sign in" value="Sign in" >Sign in</button>-->
            </form>
        </div>
        <div id="forgot" class="tab-pane">
            <form action="#" class="form-signin">
                <p class="text-muted text-center btn-block btn btn-primary btn-rect">Enter your valid e-mail</p>
                <input type="email"  required="required" placeholder="Your E-mail"  class="form-control" />
                <br />
                <button class="btn text-muted text-center btn-success" type="submit">Recover Password</button>
            </form>
        </div>
        <div id="signup" class="tab-pane">
            <form action="#" class="form-signin">
                <p class="text-muted text-center btn-block btn btn-primary btn-rect">Please Fill Details To Register</p>
                 <input type="text" placeholder="First Name" class="form-control" />
                 <input type="text" placeholder="Last Name" class="form-control" />
                <input type="text" placeholder="Username" class="form-control" />
                <input type="email" placeholder="Your E-mail" class="form-control" />
                <input type="password" placeholder="password" class="form-control" />
                <input type="password" placeholder="Re type password" class="form-control" />
                <button class="btn text-muted text-center btn-success" type="submit">Register</button>
            </form>
        </div>
    </div>
    <div class="text-center">
        <ul class="list-inline">
            <li><a class="text-muted" href="#login" data-toggle="tab">Login</a></li>
            <li><a class="text-muted" href="#forgot" data-toggle="tab">Forgot Password</a></li>
            <li><a class="text-muted" href="#signup" data-toggle="tab">Signup</a></li>
        </ul>
    </div>


</div>

	  <!--END PAGE CONTENT -->     
	      
      <!-- PAGE LEVEL SCRIPTS -->
      <script src="assets/plugins/jquery-2.0.3.min.js"></script>
      <script src="assets/plugins/bootstrap/js/bootstrap.js"></script>
   <script src="assets/js/login.js"></script>
      <!--END PAGE LEVEL SCRIPTS -->

</body>
    <!-- END BODY -->
</html>
