<?php
	session_start();
	if ($_SESSION['loginname']=="")
	{
	header("location:index.php");
	}
if ($_REQUEST['uploaded'])
{
$uploaded=$_REQUEST['uploaded'];  
}
?>				
<html>
<head>
<title>Finanace Application </title>
	
		<script src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
		<script src="js/jquery-ui-1.10.3.min.js"></script>
		<script src="js/bootstrap.min.js"></script>
		<script src="js/responsive-tables.js"></script>
	
<style>
#Menupanel ul li  
{
width:220px;
height:40px;
border-bottom:1px solid white;
list-style-type:none;
padding-top:10px;
}

#Menupanel ul li a
{
text-decoration:none;
font-size:12pt;
}

</style>
<script>
</script>

</head>
<body style="background-color:lightgrey;font-family:Helvetica Neue,Helvetica,Arial,sans-serif;">
			
			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
				<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;"></h1>
			<?php
		
			
			echo "<div id='logindiv' style='position:absolute;color:#ffffff;font-size:10pt;width:180px;height:20px;right:120px;top:10px;'>
			<i class='iconsweets-user'></i>
			Welcome " . $_SESSION['loginname']
			. "&nbsp; <a style= 'color:#ffffff;text-decoration:none;' href='index.php?logout=YES'> Logout </a></div>";
	
			?>
			</div>
	
			<div class="leftpanel" style="position:absolute;top:8%;width:15%;height:100%;background-color:#000000;" >
	
			
			<div class="leftmenu" style="position:absolute;top:5%;" >      
            <ul class="nav nav-tabs nav-stacked" style="border:none;position:relative;top:17%;" >
				<li><a href="mainpage.php"><i class="iconsweets-home"></i> home</a></li>
                <li><a href="incomestatement.php"><i class="iconsweets-pricetag"></i> Income Statement</a></li>
                <li><a href="balancesheet.php"><i class="iconsweets-cog"></i> Balance Sheet</a></li>
                <li><a href="cashflow.php"><i class="iconsweets-recycle"></i> Cash_Flow</a></li>
                <li><a href="uploadexcelfile.php"><i class="iconsweets-excel"></i> Excelupload File</a></li>
                <li><a href="convertfile.php"><i class="iconsweets-shuffle"></i> Convert to Output File</a></li>
                <li style="border:none;"><a href="downloadoutputfile.php"><i class="iconsweets-download2"></i> Download Output File </a></li>
			</ul>	
			
			</div>
			</div>


			<div  id="mainwindow" style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;font-size:12pt;position:absolute;left:300px;top:200px;background-color:#6B8E23;">
				<form action= "upload_file.php"  method="POST" enctype="multipart/form-data">
				<label for="company" style="position:absolute;left:100px;top:50px;width:150px;">Company Name:</label>
				<input type="text" style="position:absolute;left:250px;top:50px;" name="companyname" id="company"><br>
				<label for="file" style="position:absolute;left:100px;top:100px;">Filename:</label>
				<input type="file" style="position:absolute;left:250px;top:100px;" name="file" id="file"><br>
				<input type="submit"  style="position:absolute;left:100px;top:150px;" name="submit" value="Submit"><br>
				<?php
				if ($uploaded<>"")
				{
				echo "<div style='width:250px;position:absolute;left:250px;top:10px;color:green;'>". $uploaded . "</div>" ; 
				}
				?>

			</form>

			</div>


</body>

</html>
