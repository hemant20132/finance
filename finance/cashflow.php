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
		<script src="http://code.jquery.com/jquery-1.11.0.min.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
		
	
	<script>


	$(document).ready(function()
	{
			$("#incomesave").click(function()
			{
				if ( $("#company").val()=="" )
				{
				alert ("Enter Company Name!!!");
				return false;
				}
				
				if ( $("#finyear").val()=="" )
				{
				alert ("Enter Financial Year!!!");
				return false;
				}
				
				if ( $("#netincomestartingline").val()=="" )
				{
				alert ("Enter Net Income Starting Line!!!");
				return false;
				}
				
				if ( $("#depreciation").val()=="" )
				{
				alert ("Enter Depreciation!!!");
				return false;
				}
				
				if ( $("#amortization").val()=="" )
				{
				alert ("Enter Amortization!!!");
				return false;
				}
				
				if ( $("#deferredtaxes").val()=="" )
				{
				alert ("Enter Deferred Taxes!!!");
				return false;
				}
				
				if ( $("#noncashitems").val()=="" )
				{
				alert ("Enter Non Cash Items!!!");
				return false;
				}
				
				if ( $("#changeworkingcapital").val()=="" )
				{
				alert ("Enter Change in working capital !!!");
				return false;
				}
								
				if ( $("#accountreceivable").val()=="" )
				{
				alert ("Enter Change in working capital !!!");
				return false;
				}
				
				if ( $("#inventories").val()=="" )
				{
				alert ("Enter Inventories !!!");
				return false;
				}
				
				if ( $("#inventories").val()=="" )
				{
				alert ("Enter Inventories !!!");
				return false;
				}
				
				
				if ( $("#accountspayable").val()=="" )
				{
				alert ("Enter Accuntspayable !!!");
				return false;
				}
				
				if ( $("#otherliabilities").val()=="" )
				{
				alert ("Enter Other liabilities !!!");
				return false;
				}
				
				if ( $("#otheropcashflow").val()=="" )
				{
				alert ("Enter Other Cash Flow!!!");
				return false;
				}
				
				if ( $("#cashfrmoperations").val()=="" )
				{
				alert ("Enter Cash from operations!!!");
				return false;
				}
				
				if ( $("#purchaseoffixedassets").val()=="" )
				{
				alert ("Enter Purchase of Fixed Assets!!!");
				return false;
				}
				
				
				if ( $("#acquisitbusiness").val()=="" )
				{
				alert ("Enter Acquisition of Business!!!");
				return false;
				}
				
				if ( $("#saleofbusiness").val()=="" )
				{
				alert ("Enter Sale of Business!!!");
				return false;
				}
				
				if ( $("#saleoffixedassets").val()=="" )
				{
				alert ("Enter Sale of Fixed Assets!!!");
				return false;
				}
				
				if ( $("#otherinvestingcashflow").val()=="" )
				{ 
				alert ("Enter Other Investing Cash Flow !!!");
				return false;
				}
				
				if ( $("#cashfrominvesting").val()=="" )
				{ 
				alert ("Enter Cash from Investing !!!");
				return false;
				}
				
				if ( $("#financingcashflow").val()=="" )
				{ 
				alert ("Enter financing cash flow!!!");
				return false;
				}
				
				if ( $("#totalcashdividendspaid").val()=="" )
				{ 
				alert ("Enter total cash divident paid!!!");
				return false;
				}
				
				if ( $("#issuanceretirementofstock").val()=="" )
				{ 
				alert ("Enter isssuance of retirement stock!!!");
				return false;
				}
				
				if ( $("#issuanceofdebt").val()=="" )
				{ 
				alert ("Enter isssuance of debt!!!");
				return false;
				}
				
				if ( $("#cashfromfinancing").val()=="" )
				{ 
				alert ("Enter Cash from financing!!!");
				return false;
				}
				
				if ( $("#netchangeincash").val()=="" )
				{ 
				alert ("Enter Net change in cash!!!");
				return false;
				}
				
				if ( $("#netcashbeginingbalance").val()=="" )
				{ 
				alert ("Enter net cahs beginning balance !!!");
				return false;
				}
				
				if ( $("#netcashendingbalance").val()=="" )
				{ 
				alert ("Enter net cahs ending balance !!!");
				return false;
				}
				return true;
				
	});			
	});			



	
			
	
		
	</script>


	<style>
		table tr td
		{	
		vertical-align: middle;
		
		}
		table tr td input
		{	
		color:black;
		}
	</style>
	
	
</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif; background-color:#ebebeb;overflow:hidden;">

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
			   <li style="border:none;"><a href="deletedata.php"><i class="iconsweets-trashcan"></i> Delete Entire Data </a></li>
			</ul>	
			
			</div>
			</div>

	<div  id="mainwindow" style="background-color:#ffffff;width:85%;position:absolute;left:15%;top:8%;height:90%;overflow-y:scroll;">
	
			<?php 
				$poststatus=$_GET['poststatus'];
				if ($poststatus!="")
				{ ?>
				<script>
				alert ( "<?php echo $poststatus; ?>" );
				</script>
			<?php }
				?>
	<form class="stdform" action="postcashflowdata.php" method="post">
		
			<table style="position:absolute;left:3%;font-family:Helvetica Neue,Helvetica,Arial,sans-serif; font-size:10pt;">
			<tr>
			<td colspan="7"><b>OPERATING ACTIVITIES</b></td>
			</tr>
			
			<tr>
			<td width="30">1</td>
			<td width="200" >Company Name</td>
			<td><input type="text" class="form-control" name="company" id="company" value=""  />
			</td>
			<td width="150">&nbsp;</td>
			<td width="30">2</td>
			<td width="200">Financial Year</td>
			<td><input type="text" class="form-control" name="finyear" id="finyear" value=""  />
			</td>
			</tr>
		
			
			<tr>
			<td>3</td>
			<td>Net Income/Starting Line</td>
			<td><input type="text" class="form-control" name="netincomestartingline" id="netincomestartingline" value=""  />
			</td>
			<td width="150">&nbsp;</td>
			<td width="30">4</td>
			<td>Depreciation/Depletion</td>
			<td><input type="text" class="form-control" name="depreciation" id="depreciation" value=""  />
			</td>
			</tr>

			<tr>
			<td>5</td>
			<td>Amortization</td>
			<td><input type="text"  class="form-control" name="amortization" id="amortization" value=""  />
			</td>
			<td width="150">&nbsp;</td>
			<td>6</td>
			<td>Deferred Taxes</td>
			<td><input type="text"  class="form-control" name="deferredtaxes" id="deferredtaxes" value=""  />
			</td>
			</tr>

			
			<tr>
			<td>7</td>
			<td>Non-Cash Items</td>
			<td><input class="form-control" name="noncashitems" id="noncashitems" value=""  />
			</td>
			<td width="150">&nbsp;</td>
			<td>8</td>
			<td>Changes in Working Capital</td>
			<td>
			<input type="text" class="form-control" name="changeworkingcapital" id="changeworkingcapital" value=""  />
			</td>
			</tr>
		
			<tr>
			<td>9</td>
			<td>Account Receivable</td>
			<td>
			<input type="text" class="form-control"   name="accountreceivable" id="accountreceivable" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>10</td>
			<td>Inventories </td>
			<td>
			<input type="text"  class="form-control" name="inventories" id="inventories" value=""  />
			</td>
			</tr>
		
			<tr>
			<td>11</td>
			<td>Accounts Payable</td>
			<td>
			<input type="text" class="form-control"  name="accountspayable" id="accountspayable" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>12</td>
			<td>Other Liabilities</td>
			<td>
			<input type="text" class="form-control" name="otherliabilities" id="otherliabilities" value=""  />
			</td>
			</tr>
		

			<tr>
			<td>13</td>
			<td>
			Other Operating Cash Flow
			</td>
			<td>
			<input type="text" class="form-control" name="otheropcashflow" id="otheropcashflow" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>14</td>
			<td>
			Cash from Operations
			</td>
			<td>
			<input type="text" class="form-control"  name="cashfrmoperations" id="cashfrmoperations" value=""  />
			</td>
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
			
			
			
			
			<tr>
			<td colspan="7"><b>INVESTING ACTIVITIES</b></td>
			</tr>
	
			<tr>	
			<td width="30">15</td>
			<td>
			Purchase of Fixed Assets
			</td>
			<td>
			<input type="text" class="form-control" name="purchaseoffixedassets" id="purchaseoffixedassets" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>16</td>
			<td>
			 Acquisition of Business
			</td>
			<td>
			<input type="text" class="form-control" name="acquisitbusiness" id="acquisitbusiness" value=""  />
			</td>
			</tr>
	
			<tr>
			<td>17</td>
			<td>
			Sale of Business 
			</td>
			<td>
			<input type="text" class="form-control" name="saleofbusiness" id="saleofbusiness" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>18</td>
			<td>
			Sales of Fixed Asssets
			</td>
			<td>
			<input type="text" class="form-control" name="saleoffixedassets" id="saleoffixedassets" value=""  />
			</td>
			</tr>
			
			<tr>
			<td>19</td>
			<td>
			Other Investing Cash Flow
			</td>
			<td>
			<input type="text"  class="form-control"   name="otherinvestingcashflow" id="otherinvestingcashflow" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>20</td>
			<td>
			Cash from Investing
			</td>
			<td>
			<input type="text"  class="form-control"  name="cashfrominvesting" id="cashfrominvesting" value=""  />
			</td>
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
			
			<tr>
			<td colspan="7"><b>FINANCING ACTIVITIES</b></td>
			</tr>

			<tr>
			<td width="30">21</td>
			<td>
			Financing Cash Flow
			</td>
			<td>
			<input type="text" class="form-control" name="financingcashflow" id="financingcashflow" value=""  />
			</td>
			<td width="100">&nbsp;</td>
	
			<td>22</td>
			<td>
			Total Cash Dividends Paid
			</td>
			<td>
			<input type="text" class="form-control"  name="totalcashdividendspaid" id="totalcashdividendspaid" value=""  />
			</td>
			</tr>
	
			<tr>
			<td>23</td>
			<td>
			Issuance (Retirement) of Stock
			</td>
			<td>
			<input type="text" class="form-control"  name="issuanceretirementofstock" id="issuanceretirementofstock" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>24</td>
			<td>
			Issuance (Retirement) of Debt
			</td>
			<td>
			<input type="text" class="form-control" name="issuanceofdebt" id="issuanceofdebt" value=""  />
			</td>
			</tr>
	
			<tr>
			<td>25</td>
			<td>
			Cash from Financing
			</td>
			<td>
			<input type="text" class="form-control" name="cashfromfinancing" id="cashfromfinancing" value=""  />
			</td>
			<td width="100"> &nbsp; </td>
			<td>26</td>
			<td>
			Net Change in Cash
			</td>
			<td>
			<input type="text" class="form-control"  name="netchangeincash" id="netchangeincash" value=""  />
			</td>
			</tr>

			<tr>
			<td>27</td>
			<td>
			Net Cash - Beginning Balance
			</td>
			<td>
			<input type="text" class="form-control"  name="netcashbeginingbalance" id="netcashbeginingbalance" value=""  />
			</td>
			<td width="100">&nbsp;</td>
			<td>28</td>
			<td>
			Net Cash - Ending Balance
			</td>
			<td>
			<input type="text" class="form-control" name="netcashendingbalance" id="netcashendingbalance" value=""  />
			</td>
			</tr>
			
			<tr>
			<td>29</td>
			<td>Are all three sheets for same comapany entered correctly ?</td>
			<td>
	
				<select id="dataentered" name="dataentered" />
					<option value="YES">YES</option>
					<option value="NO" selected>NO</option>
				</select>
				
				</td>
			<td width="100">&nbsp;</td>
			<td>&nbsp;</td>
			
			<td  align="center"><input  id="savecashflow"  onclick="savecashflow_click()" style="border-radius:5px;height:25px;"  type="submit" value="Save" ></td>
		
			<td><input type="button" value="Cancel" class="Default" style="border-radius:5px;height:25px;" ></td>
			</tr>
		
		
	</table>
	</form>		
	</div>
	
	
</body>
</html>



