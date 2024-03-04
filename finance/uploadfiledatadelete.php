<?php
	session_start();
	if ($_SESSION['loginname']=="")
	{
	header("location:index.php");
	}
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
	
	$('div.leftmenu ul li.dropdown1 ul').hide();
	$('div.leftmenu ul li.dropdown1').click(function()
	{
		$('div.leftmenu ul li.dropdown1 ul').slideToggle();
	});

	$('div.leftmenu ul li.dropdown1 ul li').click(function()
	{
		$('div.leftmenu ul li.dropdown1 ul').show();

	});
	
	var rowclicked=false;
	$('#uploadfiledatatable').dataTable( {
        "processing": true,
        "serverSide": true,
        "ajax": "uploadfiledata_processing.php"
	});

  
  var table = $('#uploadfiledatatable').DataTable();
 
	$('#uploadfiledatatable tbody').on( 'click', 'tr', function () {
        if ($(this).hasClass('selected')) {
            $(this).removeClass('selected');
		}
        else {
			table.$('tr.selected').removeClass('selected');
            $(this).addClass('selected');
		}
    });
	
	
    $('#uploadfiledatadelete').click( function () {
	
	var company4  = $("#uploadfiledatatable").find(".selected td:eq(1)").text();
	var inputfilename4 = $("#uploadfiledatatable").find(".selected td:eq(2)").text();
	
	
	if (company4 == "" || inputfilename4 == "")
	{
		alert("select row!!!");
		return false;
	}

	if (company4 != "" && inputfilename4 != "")
	{
	location.href = "updeleterecord.php?companyupload="+company4+" &inputfilename="+inputfilename4 ;
	}	
	
	});
   
 
});
</script>


</head>
<body style="background-color:#ebebeb;overflow:hidden;">

<div style="">
			
			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
				<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;">Uniq Star Net Solution</h1>
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
				<li class="dropdown1"><a href="#"> <i class="iconsweets-tag"></i> Delete Data </a>
            			<ul style="position:relative;left:8px;">
                            <li><a href="incomedelete.php" style="text-decoration:none;">IncomeStatement </a></li>
                            <li><a href="balancedelete.php" style="text-decoration:none;">Balance Sheet </a></li>
							<li><a href="cashflowdelete.php" style="text-decoration:none;">Cash Flow </a></li>
							<li><a href="uploadfiledatadelete.php" style="text-decoration:none;">Upload File Data</a></li>
							
                        </ul>
				</li>
			</ul>	
			
			</div>
			</div>

</div>
		
			<div  id="mainwindow" style="position:absolute;left:300px;top:100px;">
				<input type="button" style="position:absolute;left:0px;top:-20px;" class="default" id="uploadfiledatadelete" value="delete selected row record">
				
				<form action="#"  method="POST">
				<center><h4> Uplaod File Data </h4></center>
				<table id="uploadfiledatatable" class="display" style=""  width="100%">
					<thead>
						<tr>
							<th>Serial No.</th>
							<th>Company Name</th>
							<th>Upload File Name</th>
							<th>Output File Name</th>
							</tr>
					</thead>
					</table>
			    </form>

			</div>
</div>	
</body>
</html>
