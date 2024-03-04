<html>
<head>
<title>Finanace Application </title>
		<script src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
	
<style>
#Menupanel ul li  
{
width:220px;
height:40px;
border-bottom:1px solid yellow;
list-style-type:none;
padding-top:10px;
}

#Menupanel ul li a
{
text-decoration:none;
font-size:12pt;
color:white;
}

</style>
<script>
$(document).ready(function()
{
$("#Menupanel").hide();
$("#menubutton").click(function()
{
$("#Menupanel").slideToggle();
});
});
</script>

</head>

<body style="background-color:#6B8E23;font-family:helvetica;">

<div style="background-color:#6B8E23;">
			<div  style="position:absolute;width:99%;height:10%;background-color:green;" id="topdiv" >
				<input type="button" id="menubutton" style="width:100px;" style="color:blue;" value="menu" >	
				<h1 style="color:yellow;position:absolute;left:40%;top:0%;">Finance Application</h1>
			</div>
	
			<div id="Menupanel" style="position:absolute;top:10%;left:9px;width:250px;height:90%;background-color:green;z-index:1;" >
			<ul  style="position:relative;left:-15px;" >
    			<li><a style="background-color:green;color:white;font-family:san-sarif,font-size:11pt;" href="incomestatement.php">Income Statement</a></li>
                <li><a href="balancesheet.php">Balance Sheet</a></li>
                <li><a href="cashflow.php">Cash_Flow</a></li>
                <li><a href="uploadexcelfile.php">Excelupload File</a></li>
                <li><a href="convertfile.php">Convert to Output File</a></li>
                <li><a href="financialratios.php">Financial Ratios </a></li>
				<li><a href="dupontanalysis.php">Du Pont Analysis</a></li>
				<li><a href="financialdashboard.php">financial dashboard</a></li>
			</ul>	
			</div>

			<div  id="mainwindow" style="background-color:#6B8E23;">

			</div>
</div>	
</body>
</html>
