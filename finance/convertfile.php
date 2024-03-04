<?php
ob_start();
	session_start();
	if ($_SESSION['loginname']=="")
	{
	header("location:index.php");
	}
$status="false";
$status=$_GET['msg'];
?>				
<html>
<head>

<title>Finanace Application </title>
		<script src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
		<link rel="stylesheet" href="http://cdn.datatables.net/1.10.2/css/jquery.dataTables.min.css">
		<script src="http://cdn.datatables.net/1.10.2/js/jquery.dataTables.min.js" ></script>
		<script src="TableTools.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
		<script src="js/jquery-ui-1.10.3.min.js"></script>
		<script src="js/bootstrap.min.js"></script>
<style>
table.dataTable thead .sorting
{
width:150px;
}
table.dataTable thead .sorting_asc
{
width:150px;
}

table.dataTable thead .sorting_desc
{
width:150px;
}

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
$(document).ready(function()
{
 	var rowclicked=false;
	$('#example').dataTable(
	{
        "processing" :true,
        "serverside" :true,
        "ajax" : "server_processing.php"
	});

  
	$('#example tbody').on( 'click', 'tr', function () {
        if ($(this).hasClass('selected')) {
            $(this).removeClass('selected');
		}
        else {
            $(this).addClass('selected');
		}
    });
	
	$('#example thead tr th.sorting_asc').css("width","150px");
	
	$('#example tr').click(function()	
	{
		rowclicked="true";
	});
	
   
    $('#convert').click(function () 
	{
	
	if ($('#example').find('.selected td:eq(3)').text()=='YES' && $('#example').find('.selected td:eq(4)').text() != '' )
	{
	alert ("Data for this excel sheet already posted to database");
	}
	
	var company1  = $("#example").find(".selected td:eq(1)").text();
	var inputfile1 = $("#example").find(".selected td:eq(2)").text();

	if (company1 == "" || inputfile1 == "")
	{
		alert("select row!!!");
		return false;
	}
	
	if (inputfile1 == "MANUAL")
	{
		location.href = "output.php?companyname="+company1+ "&inputfile="+inputfile1;
		return false;
	}
	
	if (inputfile1 != "MANUAL")
	{
		location.href = "convertupload1.php?companyname="+company1+"&inputfile="+inputfile1;
		return false;
	}
	
	});
 
 
});
</script>


</head>
<body style="background-color:#ebebeb;overflow:hidden;">

<div style="">
			
			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
			<?php
			echo "<div id='logindiv' style='position:absolute;color:#ffffff;font-size:10pt;width:180px;height:20px;right:120px;top:10px;'>
			<i class='iconsweets-user'></i>
			Welcome " . $_SESSION['loginname']
			. "&nbsp; <a style= 'color:#ffffff;text-decoration:none;' href='index.php?logout=YES'> Logout </a></div>";
			?>
			</div>
	
			<div class="leftpanel" style="position:absolute;top:8%;width:15%;height:100%;font-family:Helvetica Neue,Helvetica,Arial,sans-serif;background-color:#000000;" >
				
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

</div>
		
			<div  id="mainwindow" style="position:absolute;left:300px;top:100px;">
				<input type="button" style="position:absolute;left:0px;top:-10px;" class="default" id="convert" value="generate output file for selected row ">
				
				<form action="#"  method="POST">
				<center><h4> uploaded file data </h4></center>
				<table id="example" class="display" style=""  width="100%">
					<thead>
						<tr>
							<th>Serial No.</th>
							<th>Company Name</th>
							<th>Uploaded File</th>
							<th>Data Posted to DB</th>
							<th>Output File</th>
						</tr>
					</thead>
					</table>
			    </form>

			</div>
</div>	


</body>

</html>
