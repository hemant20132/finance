<?php
			session_start();
			$_SESSION['loginname']= $_GET['loginname'];
	
	if ( $_SESSION['loginname']!='' )
	{
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
function sureyesno()
{
var answer = confirm ("Are you sure about deleting entire records ?")
if (answer)
location.href="deletedata.php";
else
return false;
}

$(document).ready(function()
{
	$('div.leftmenu ul li.dropdown1 ul').hide();
	$('div.leftmenu ul li.dropdown1').click(function()
	{
		$('div.leftmenu ul li.dropdown1 ul').slideToggle();
	});

	$('div.leftmenu ul li.dropdown1 ul li').click(function()
	{
		$('div.leftmenu ul li.dropdown1 ul').show();

	});
	
});
</script>	

<style>
	
</style>
	
</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;background-color:#ebebeb;overflow:hidden;">

<div style="">
			
			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
			<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;"></h1>
				
			<?php
		
			
			echo "<div id='logindiv' style='position:absolute;color:#ffffff;font-size:10pt;width:180px;height:20px;right:120px;top:10px;'>
			<i class='iconsweets-user'></i>
			Welcome " . $_SESSION['loginname']
			. "&nbsp; <a style= 'color:#ffffff;text-decoration:none;' href='index.php?logout=YES'> Logout </a></div>";
	
			?>
			</div>
			
			
			
			<?php
			if ($_GET['msg']!='')
			{
			echo "<div id='msgdiv' style='font-size:14pt;position:absolute;width:150px;height:50px;left:300px;top:300px;border 1px solid lightgrey;'>
			".$msg."	
			</div>";
			}
			?>

			
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
			
				<li class="dropdown1"><a href="#"> <i class="iconsweets-tag"></i> Delete Data </a>
            			<ul style="position:relative;left:8px;">
                            <li><a href="incomedelete.php" style="text-decoration:none;">IncomeStatement </a></li>
                            <li><a href="balancedelete.php" style="text-decoration:none;">Balance Sheet </a></li>
							<li><a href="cashflowdelete.php" style="text-decoration:none;">Cash Flow </a></li>
							<li><a href="uploadfiledatadelete.php" style="text-decoration:none;">Upload File Data </a></li>
							
						</ul>
				</li>
				
			</ul>	
			
			</div>
			</div>
			
			
			
</div>


</body>
</html>
<?php }
else 
{
header("location:index.php");
}
 ?>