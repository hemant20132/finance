<html>
<head>
<title>Finanace Application </title>
		<script src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
	
	
</head>

<body style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif;background-color:#ebebeb;overflow:hidden;">

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
	<form class="stdform" action="postincomestatdata.php" method="post">
	<table style="position:absolute;left:20%;font-family:Helvetica Neue,Helvetica,Arial,sans-serif;font-size:10pt;">
			<col width="50">
			<col width="250">
			<col width="200">
	
			<tr style="background-color:#ebebeb;">
			<th width="50">Sr. No.</th>
			<th>Name of the Field </th>
			<th>Value of the Field </th>
			</tr>
				
			<tbody>
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
			<td colspan="4"><b>SALES</b></td>
			</tr>
			
			<tr>
			<td width="50">3</td>
			<td>Total Revenue / Sales</td>
			<td><input type="text" class="form-control"  name="revsales" id="revsales" value=""></td>
			</tr>
			
			<tr>
			<td>4</td>
			<td>Cost of Revenue	</td>
			<td><input type="text"  class="form-control" name="costofrevenue" id="costofrevenue" value=""></td>
			</tr>
			
			<tr>
			<td>5</td>
			<td>Gross Profit</td>
			<td><input type="text" class="form-control" name="grossprofit" id="grossprofit" value=""></td>
			</tr>
	
			<tr>
			<td colspan="4"><b>OPERATING EXPENSES</b></td>
			</tr>
			
			<tr>
			<td>6</td>
			<td>General/Administrative Expenses </td>
			<td>
			<input type="text" class="form-control" name="generaladmexpenses" id="generaladmexpenses" value=""></td>
			</tr>

			
			<tr>
			<td>7</td>
			<td>Research & Development</td>
			<td>
			<input type="text" class="form-control" name="researchdev" id="researchdev" value="" ></td>
			</tr>	
			
			<tr>
			<td>8</td>
			<td>Sales & Marketing Expenses </td>
			<td>
			<input type="text" class="form-control" name="salesmarketingexp" id="salesmarketingexp" value="" ></td>
			</tr>

			<tr>
			<td>9</td>
			<td>Depreciation & Amortization </td>
			<td>
			<input type="text" class="form-control" name="depreciationamort" id="depreciationamort" value="" ></td>
			</tr>

			<tr>
			<td>10</td>
			<td>Interest Expenses </td>
			<td>
			<input type="text" class="form-control" name="interestexp1" id="interestexp1" value="" ></td>
			</tr>
		
			<tr>
			<td>11</td>
			<td>Type Your Business Specific Expeses</td>
			<td><input type="text" class="form-control"  name="businespecific1" id="businespecific1" value=""  ></td>
			</tr>
			
			<tr>
			<td>12</td>
			<td>Type Your Business Specific Expeses </td>
			<td><input type="text" class="form-control" name="businespecific2" id="businespecific2" value=""  ></td>
			</tr>
			
			<tr>
			<td>13</td>
			<td>Type Your Business Specific Expeses </td>
			<td><input type="text" class="form-control"  name="businespecific3" id="businespecific3" value=""  ></td>
			</tr>
			
			<tr>
			<td>14</td>
			<td>Miscellaneous </td>
			<td><input type="text"  class="form-control" name="miscexp" id="miscexp" value=""  ></td>
			</tr>
			
			<tr>
			<td>15</td>
			<td>Operating Expenses </td>
			<td><input type="text" class="form-control"  name="operatingexp" id="operatingexp" value=""  ></td>
			</tr>
	
	
			<tr>
			<td style="font-weight:900;" colspan="4">PROFIT</td>
			</tr>
			
			<tr>
			<td>16</td>
			<td>Operating Income(EBIT)</td>
			<td><input type="text" class="form-control"  name="operatinginc" id="operatinginc" value=""  ></td>
			</tr>
			
			<tr>
			<td>17</td>
			<td>Interest Expenses</td>
			<td><input type="text" class="form-control" name="interestexp2" id="interestexp2" value=""  ></td>
			</tr>
			
			<tr>
			<td>18</td>
			<td>Income Before Tax</td>
			<td><input type="text" class="form-control"  name="incomebeforetax" id="incomebeforetax" value=""  ></td>
			</tr>
			
			<tr>
			<td>19</td>
			<td>Tax</td>
			<td><input type="text" class="form-control" name="tax" id="tax" value="" ></td>
			</tr>
			
			<tr>
			<td>20</td>
			<td>Net Profit</td>
			<td><input type="text" class="form-control" name="netprofit" id="netprofit" value=""></td>
			</tr>
	
			<tr>
			<td colspan="2" align="center"><input  style="border-radius:5px;height:25px;" type="submit" value="Save" ></td>
			<td><input type="button" value="Cancel" class="Default" style="border-radius:5px;height:25px;" ></td>
			</tr>
		</tbody>
	</table>
	
	

	
	</form>		
			
			
			
	</div>
	
	
			
			
			
			

</body>
</html>



