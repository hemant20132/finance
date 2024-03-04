<?php
session_start();
?>
<html>
<head>
<title>Finanace Application </title>
		<script src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
		<script src="js/jquery-ui-1.10.3.min.js"></script>
		<script src="js/bootstrap.min.js"></script>
		<script src="js/responsive-tables.js"></script>
	
<script>

$(document).ready(function()
{	
	$("#loginname").val()=="";
	$("#loginpass").val()=="";
	
	$("#Submit").click(function()
	{
		if ( $("#loginname").val()=="")
		{
		alert ("Enter Login Name!!!");
		return false;
		}
		if ( $("#loginpass").val()=="")
		{
		alert ("Enter Login Password!!!");
		return false;
		}
	});
});
</script>	

<style>
		table tr td
		{	
		vertical-align: middle;
		}

</style>
	
</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;background-color:#ebebeb;overflow:hidden;">

<div style="">
			
			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
				<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;"></h1>
			</div>
			
			<div id="logindiv"  style="position:absolute;left:35%;top:30%;width:350px;height:120px;border:1px solid white;">
			<?php 
				
				if	( $_GET['logout']='YES' )
				{
					session_destroy();
					echo "logged out!!!";	
				}
			?>
		
			<form id="loginform" class="stdform" action="checklogin.php" method="post" autocomplete="off">
				<table style="margin-top:5px;margin-left:5px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif;font-size:10pt;" width="300px">
				<tr>
				<td width="60px" align="right">
				Login Name:</td>
				<td width="50px">&nbsp;</td>
				<td width="200px"><input class="form-control" width="100%" type="text" id="loginname"  value="" name="loginname" placeholder="wifiwifi" >
				</td>
				
				</tr>
				<td width="60px" align="right">
				Password:
				</td>
				<td width="50px">&nbsp;</td>
				<td width="200px">
				<input class="form-control" width="100%" id="loginpass" name="loginpass" type="password"  value="" placeholder="wifiwifi">
				</td>
				</tr>
				
				</tr>
				<td width="200px" align="right">
				<input  type="submit" value="Submit" width="100px" id="Submit" class="Default" align="middle" style="border-radius:5px;height:25px;align:middle">
				</td>
				<td width="50px">&nbsp;</td>
				<td width="200px">
				<input id="cancelbutton" type="button" value="Cancel" class="Default" align="middle" style="border-radius:5px;height:25px;align:middle" >
				</td>
				</tr>
				
				<tr>
				<td width="50px">&nbsp;</td>
				<td width="50px">&nbsp;</td>
				<td width="50px">&nbsp;</td>
				</tr>
				
				<tr>
				<td width="50px">&nbsp;</td>
				<td width="50px">&nbsp;</td>
				<td width="50px">&nbsp;</td>
				</tr>
				
				</table>
			</form>
			</div>
			
			
			
</div>


	
</body>
</html>
