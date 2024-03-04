<html>
<head>
<title>Finanace Application </title>
		<script src="http://code.jquery.com/jquery-1.11.0.min.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
		
	
	<script>

		$(document).ready(function()
		{
		$("table tr td #savecashflow").click(function()
		{
				if (("table tr td input #company").value()=="")
				{
					alert ("Enter company name");
					return false;
				}
				
				if (("table tr td input #finyear").value()=="")
				{
					alert ("Enter financial year");
					return false;					
				}
		});
		});
		
	
	</script>
</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif; background-color:#ebebeb;overflow:hidden;">

			<div  style="position:absolute;width:100%;top:0px;height:8%;background-color:#000000;" id="topdiv" >
				<h1 style="color:#ffffff;position:relative;left:1%;top:0%;font-size:14pt;">Uniq Star Net Solution</h1>
			</div>
			
			<div class="leftpanel" style="position:absolute;top:8%;width:15%;height:100%;background-color:#000000;" >
			
			<div class="leftmenu" style="position:absolute;top:5%;" >      
            <ul class="nav nav-tabs nav-stacked" style="border:none;position:relative;top:17%;" >
				<li><a href="index.php"><i class="iconsweets-home"></i> home</a></li>
                <li><a href="incomestatement.php"><i class="iconsweets-pricetag"></i> Income Statement</a></li>
                <li><a href="balancesheet.php"><i class="iconsweets-cog"></i> Balance Sheet</a></li>
                <li><a href="cashflow.php"><i class="iconsweets-recycle"></i> Cash_Flow</a></li>
                <li><a href="uploadexcelfile.php"><i class="iconsweets-excel"></i> Excelupload File</a></li>
                <li><a href="convertfile.php"><i class="iconsweets-shuffle"></i> Convert to Output File</a></li>
                <li style="border:none;"><a href="downloadoutputfile.php"><i class="iconsweets-download2"></i> Download Output File </a></li>
			</ul>	
			
			</div>
			</div>

	<div  id="mainwindow" style="background-color:#ffffff;width:85%;position:absolute;left:15%;top:8%;height:90%;overflow-y:scroll;">
	
			<?php 
				$poststatus=$_GET['poststatus'];
				if ($poststatus!="")
				{
				echo "<div style='position:absolute;top:4%;left:40%;color:red;'><blink>".$poststatus."</blink></div>";
				}
			?>
	<form class="stdform" action="postcashflowdata.php" method="post">
		
			<table style="position:absolute;left:20%;font-family:Helvetica Neue,Helvetica,Arial,sans-serif; font-size:10pt;">
			<col width="50">
			<col width="250">
			<col width="200">
	
			<tr style="background-color:#ebebeb;">
			<th width="50">Sr. No.</th>
			<th>Name of the Field </th>
			<th>Value of the Field </th>
			<th width="50"></th>
			<th width="50">Sr. No.</th>
			<th>Name of the Field </th>
			<th>Value of the Field </th>
			
			</tr>
				
			<tbody>
			<tr>
			<td colspan="4"><b>OPERATING ACTIVITIES</b></td>
			</tr>
			
			<tr>
			<td>1</td>
			<td>Company Name</td>
			<td><input type="text" class="form-control" name="company" id="company" value=""  />
			</td>
			</tr>
		
			<tr>
			<td>2</td>
			<td>Financial Year</td>
			<td><input type="text" class="form-control" name="finyear" id="finyear" value=""  />
			</td>
			</tr>
		
			
			<tr>
			<td>3</td>
			<td>Net Income/Starting Line</td>
			<td><input type="text" class="form-control" name="netincomestartingline" id="netincomestartingline" value=""  />
			</td>
			</tr>
		
			<tr>
			<td>4</td>
			<td>Depreciation/Depletion</td>
			<td><input type="text" class="form-control" name="depreciation" id="depreciation" value=""  />
			</td>
			</tr>

			<tr>
			<td>5</td>
			<td>Amortization</td>
			<td><input type="text"  class="form-control" name="amortization" id="amortization" value=""  />
			</td>
			</tr>

			<tr>
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
			</tr>
			
			<tr>
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
			</tr>
			
		
			<tr>
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
			</tr>
		
			<tr>
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
			</tr>
		
		
			<tr>
			<td>14</td>
			<td>
			Cash from Operations
			</td>
			<td>
			<input type="text" class="form-control"  name="cashfrmoperations" id="cashfrmoperations" value=""  />
			</td>
			</tr>
		
			<tr>
			<td colspan="4"><b>INVESTING ACTIVITIES</b></td>
			</tr>
	
	
			<tr>
			<td>15</td>
			<td>
			Purchase of Fixed Assets
			</td>
			<td>
			<input type="text" class="form-control" name="purchaseoffixedassets" id="purchaseoffixedassets" value=""  />
			</td>
			</tr>
	
			<tr>
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
			</tr>
	
			<tr>
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
			</tr>
	
			<tr>
			<td>20</td>
			<td>
			Cash from Investing
			</td>
			<td>
			<input type="text"  class="form-control"  name="cashfrominvesting" id="cashfrominvesting" value=""  />
			</td>
			</tr>
				
			<tr>
			<td colspan="4"><b>FINANCING ACTIVITIES</b></td>
			</tr>
	

			<tr>
			<td>21</td>
			<td>
			Financing Cash Flow
			</td>
			<td>
			<input type="text" class="form-control" name="financingcashflow" id="financingcashflow" value=""  />
			</td>
			</tr>
				
	
			<tr>
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
			</tr>
	
			<tr>
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
			</tr>
			
			<tr>
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
			</tr>

			<tr>
			<td>28</td>
			<td>
			Net Cash - Ending Balance
			</td>
			<td>
			<input type="text" class="form-control" name="netcashendingbalance" id="netcashendingbalance" value=""  />
			</td>
			</tr>
			
			<tr>
			<td colspan="2" align="center"><input  id="savecashflow"  style="border-radius:5px;height:25px;"  type="submit" value="Save" ></td>
			<td><input type="button" value="Cancel" class="Default" style="border-radius:5px;height:25px;" ></td>
			</tr>
		</tbody>
	
		
	</table>
	</form>		
	</div>
	
	
</body>
</html>



