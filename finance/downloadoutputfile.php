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
	


</style>
<script>

$(document).ready(function()
{

$('#download').dataTable( {
        "processing": true,
        "serverSide": true,
        "ajax": "server_processing.php"
	
});

 var table = $('#download').DataTable();
 
    $('#download tbody').on( 'click', 'tr', function () {
        if ($(this).hasClass('selected')) {
            $(this).removeClass('selected');
		
		}
        else {
            $(this).addClass('selected');
		}
    });
   
   $('#downloadfile tr').click(function()	
	{
		rowclicked="true";
	});
	
   
    
    $('#downloadfile').click( function () {
	
	var company1 = $("#download").find(".selected td:eq(1)").text();
	var inputfile1=$("#download").find(".selected td:eq(2)").text();
	if (company1=="" || inputfile1=="")
	{
		alert("select row!!!");
		return false;
	}
	
	var filename1 = $("#download").find(".selected td:eq(4)").text();
	alert(filename1);
	location.href = "http://newsoftwares.rf.gd/finance/outputfilefolder/"+filename1;
	
	
	}); 
 

});
</script>

</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;background-color:#ebebeb;overflow:hidden;">

			
			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
				<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;"></h1>
			<?php
			echo "<div id='logindiv' style='position:absolute;color:#ffffff;font-size:10pt;width:180px;height:20px;right:120px;top:10px;'>
			<i class='iconsweets-user'></i>	Welcome " . $_SESSION['loginname']
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


			<div  id="mainwindow" style="position:absolute;left:300px;top:100px;">
		
			<input type="button"  style="position:absolute;left:0px;top:-10px;" id="downloadfile"  value="download selected row file">
				
			<form action="#"  method="POST">
				<center><h4 style=""> Download File Data </h4></center>
				<table id="download" class="display" style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;" cellspacing="0" width="100%">
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
</body>
</html>
