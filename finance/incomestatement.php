<?php	session_start();	if ($_SESSION['loginname']=="")	{	header("location:index.php");	}?><html>
<head>
<title>Finanace Application- Income Statement</title>		<script src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.1.js"></script>		<script src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.1.min.js"></script>		<link rel="stylesheet" href="css/style.default.css" />		
	<style>		table tr td		{			vertical-align: middle;		}	</style>			<script>		$(document).ready(function()	{			$("#incomesave").click(function()			{				if ( $("#company").val()=="" )				{				alert ("Enter Company Name!!!");				return false;				}								if ( $("#finyear").val()=="" )				{				alert ("Enter Financial Year!!!");				return false;				}								if ( $("#revsales").val()=="" )				{				alert ("Enter Revenue/sales Figure !!!");				return false;				}												if ( $("#costofrevenue").val()=="" )				{				alert ("Enter Cost of Revenue!!!");				return false;				}												if ( $("#grossprofit").val()=="" )				{				alert ("Enter Gross Profit!!!");				return false;				}								if ( $("#generaladmexpenses").val()=="" )				{				alert ("Enter General Admn. Expenses!!!");				return false;				}								if ( $("#researchdev").val()=="" )				{				alert ("Enter Research and Devlopment  Expenses!!!");				return false;				}								if ( $("#salesmarketingexp").val()=="" )				{				alert ("Enter Sales and Marketing  Expenses!!!");				return false;				}								if ( $("#depreciationamort").val()=="" )				{				alert ("Enter Depreciation and Amort amt.!!!");				return false;				}								if ( $("#interestexp1").val()=="" )				{				alert ("Enter Interest Expenses .!!!");				return false;				}				if ( $("#businespecific1").val()=="" )				{				alert ("Enter business specific Expenses 1 !!!");				return false;				}								if ( $("#businespecific2").val()=="" )				{				alert ("Enter business specific Expenses 2!!!");				return false;				}								if ( $("#businespecific3").val()=="" )				{				alert ("Enter business specific Expenses 3!!!");				return false;				}												if ( $("#miscexp").val()=="" )				{				alert ("Enter Miscellaneous Expenses !!!");				return false;				}								if ( $("#operatingexp").val()=="" )				{				alert ("Enter Operating Expenses !!!");				return false;				}												if ( $("#operatinginc").val()=="" )				{				alert ("Enter Operating Income !!!");				return false;				}								if ( $("#interestexp2").val()=="" )				{				alert ("Enter interest Expenses 2 !!!");				return false;				}								if ( $("#interestexp2").val()=="" )				{				alert ("Enter interest Expenses 2 !!!");				return false;				}								if ( $("#incomebeforetax").val()=="" )				{				alert ("Enter Income Before Tax !!!");				return false;				}												if ( $("#tax").val()=="" )				{				alert ("Enter Tax !!!");				return false;				}								if ( $("#netprofit").val()=="" )				{				alert ("Enter Net Profit !!!");				return false;				}								return true;							});					});				</script>	
</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;background-color:#ebebeb;overflow:hidden;">			

			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
				<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;"></h1>			<?php			echo "<div id='logindiv' style='position:absolute;color:#ffffff;font-size:10pt;width:180px;height:20px;right:120px;top:10px;'>			<i class='iconsweets-user'></i>	Welcome " . $_SESSION['loginname']			. "&nbsp; <a style= 'color:#ffffff;text-decoration:none;' href='index.php?logout=YES'> Logout </a></div>";			?>						</div>				
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
                <li style="border:none;"><a href="downloadoutputfile.php"><i class="iconsweets-download2"></i> Download Output File </a></li>				<li style="border:none;"><a href="deletedata.php"><i class="iconsweets-trashcan"></i> Delete Entire Data </a></li>							</ul>	
			
			</div>
			</div>
			
	<div  id="mainwindow" style="background-color:#ffffff;width:85%;position:absolute;left:15%;top:8%;height:90%;">
	<?php 				$poststatus=$_GET['poststatus'];				if ($poststatus!="")				{ ?>				<script>				alert ( "<?php echo $poststatus; ?>" );				</script>				<?php }				?>				
	<form class="stdform" action="postincomestatdata.php" method="post">
	<table style="position:absolute;left:51px;font-family:Helvetica Neue,Helvetica,Arial,sans-serif;font-size:10pt;top:0px;width:984px" width="100%">
		<br>		<tr><td colspan="5"><b>SALES</b></td>					<tr>		
			<td width="30">1</td>
			<td valign="top">Company Name:</td>
			<td><input type="text" class="form-control" name="company" id="company" value=""></td>
			<td width="100">&nbsp;</td>
			<td width="30">2</td>
			<td valign="top">Financial Year:</td>
			<td width="150"><input type="text" class="form-control" name="finyear" id="finyear" value=""  ></td>
		</tr>
		<tr>
			<td>3</td>
			<td>Total Revenue / Sales:</td>
			<td><input type="text" class="form-control"  name="revsales" id="revsales" value=""></td>
			<td>&nbsp;</td>
			<td>4</td>
			<td>Cost of Revenue:</td>
			<td><input type="text"  class="form-control" name="costofrevenue" id="costofrevenue" value=""></td>
		</tr>
		<tr>
			<td>5</td>
			<td>Gross Profit:</td>
			<td><input type="text" class="form-control" name="grossprofit" id="grossprofit" value=""></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="7"><b><br>
			<br>
			OPERATING EXPENSES</b></td>
		</tr>
		<tr>
			<td>6</td>
			<td>General/Administrative Expenses:</td>
			<td>				<input type="text" class="form-control" name="generaladmexpenses" id="generaladmexpenses" value="">			</td>					<td>&nbsp;</td>			
			<td>7</td>
			<td>Research & Development:</td>
			<td>
			<input type="text" class="form-control" name="researchdev" id="researchdev" value="" ></td>
		</tr>
		<tr>
			<td>8</td>
			<td>Sales & Marketing Expenses:</td>
			<td>
			<input type="text" class="form-control" name="salesmarketingexp" id="salesmarketingexp" value="" ></td>
			<td>&nbsp;</td>
			<td>9</td>
			<td>Depreciation & Amortization:</td>
			<td>
			<input type="text" class="form-control" name="depreciationamort" id="depreciationamort" value="" ></td>
		</tr>
		<tr>
			<td>10</td>
			<td>Interest Expenses</td>
			<td><input type="text" class="form-control" name="interestexp1" id="interestexp1" value="" ></td>
			<td>&nbsp;</td>
			<td>11</td>
			<td>Type Your Business Specific Expeses:</td>
			<td><input type="text" class="form-control"  name="businespecific1" id="businespecific1" value=""  ></td>
		</tr>
		<tr>
			<td>12</td>
			<td>Type Your Business Specific Expeses:</td>
			<td><input type="text" class="form-control" name="businespecific2" id="businespecific2" value=""  ></td>
			<td>&nbsp;</td>
			<td>13</td>
			<td>Type Your Business Specific Expeses:</td>
			<td><input type="text" class="form-control"  name="businespecific3" id="businespecific3" value=""  ></td>
		</tr>
		<tr>
			<td>14</td>
			<td>Miscellaneous:</td>
			<td><input type="text"  class="form-control" name="miscexp" id="miscexp" value=""  ></td>
			<td>&nbsp;</td>
			<td>15</td>
			<td>Operating Expenses:</td>
			<td><input type="text" class="form-control"  name="operatingexp" id="operatingexp" value=""  ></td>
		</tr>
		<tr>
			<td colspan="7">&nbsp;<p><b>PROFIT</b></td>
		</tr>
		<tr>
			<td>16</td>
			<td>Operating Income(EBIT):</td>
			<td><input type="text" class="form-control"  name="operatinginc" id="operatinginc" value=""  ></td>
			<td>&nbsp;</td>
			<td>17</td>
			<td>Interest Expenses:</td>
			<td><input type="text" class="form-control" name="interestexp2" id="interestexp2" value=""  ></td>
		</tr>
		<tr>
			<td>18</td>
			<td>Income Before Tax:</td>
			<td><input type="text" class="form-control"  name="incomebeforetax" id="incomebeforetax" value=""  ></td>
			<td>&nbsp;</td>
			<td>19</td>
			<td>Tax:</td>
			<td> <input type="text" class="form-control" name="tax" id="tax" value="" > </td>
		</tr>
		<tr>
			<td>20</td>
			<td>Net Profit:</td>
			<td><input type="text" class="form-control" name="netprofit" id="netprofit" value=""></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>			<td>&nbsp;</td>			<td>&nbsp;</td>			<td><input id="incomesave" style="border-radius:5px;height:25px;align:middle"  type="submit"  value="Save" ></td>			<td><input id="cancelbutton" type="button" value="Cancel" class="Default" align="middle" style="border-radius:5px;height:25px;align:middle" ></td>			<td>&nbsp;</td>			<td>&nbsp;</td>			<td>&nbsp;</td>		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
	
	</table>
&nbsp;
	
	

	
</form>		
			
			
			
</div>
	
	
			
			
			
			

</body>
</html>