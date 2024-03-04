<html>
<head>
<title>Finanace Application </title>
		<script src="http://code.jquery.com/jquery-1.10.2.min.js"></script>
		<link rel="stylesheet" href="css/style.default.css" />
		<script src="js/jquery-ui-1.10.3.min.js"></script>
		<script src="js/bootstrap.min.js"></script>
		<script src="js/responsive-tables.js"></script>
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
		<form class="stdform" action="postbalancesheet.php" method="post">
		<table style="position:absolute;left:3%;font-family:Helvetica Neue,Helvetica,Arial,sans-serif;font-size:10pt;">
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
				
	
			<tr>
			<td>1</td>
			<td>Company Name</td>
			<td><input type="text" class="form-control" name="company" id="company" value=""  />
			</td>
			<td width="100px;"></td>
			<td>2</td>
			<td>Financial Year</td>
			<td><input type="text" class="form-control" name="finyear" id="finyear" value=""  />
			</td>
			</tr>
		
			<tr>
			<td colspan="4"><b>ASSETS</b></td>
			</tr>
			
			<tr>
			<td>3</td>
			<td>Cash & Equivalents</td>
			<td><input type="text" class="form-control" name="cashandequivalents" id="cashandequivalents" value=""></td>
			<td width="100px;"></td>
			<td>4</td>
			<td>Accounts Receivable</td>
			<td><input type="text" class="form-control" name="accountreceivables" id="accountreceivables" value=""></td>
			</tr>
		
			<tr>
			<td>5</td>
			<td>Inventory</td>
			<td><input type="text" class="form-control" name="inventory" id="inventory" value="" ></td>
			<td width="100px;"></td>
			<td>6</td>
			<td>Prepaid Expenses</td>
			<td><input type="text" class="form-control" name="prepaidexpenses" id="prepaidexpenses" value=""></td>
			</tr>
			
			<tr>
			<td>7</td>
			<td>Total Current Assets</td>
			<td><input type="text" class="form-control"  name="totalcurrentassets" id="totalcurrentassets" value=""></td>
			<td width="100px;"></td>
			<td>8</td>
			<td>Property/Equipment</td>
			<td><input type="text" class="form-control" name="propertyorequipment" id="propertyorequipments" value=""></td>
			</tr>
			
			<tr>
			<td>9</td>
			<td>Goodwill </td>
			<td><input type="text" class="form-control" name="goodwill" id="goodwill" value=""></td>
			<td width="100px;"></td>
			<td>10</td>
			<td>Intangibles</td>
			<td><input type="text" class="form-control"  name="intangibles" id="intagibles" value="" ></td>
			</tr>
	
			<tr>
			<td>11</td>
			<td>Long Term Investments</td>
			<td>
			<input type="text" class="form-control"  name="longterminvestments" id="longterminvestments" value="" ></td>
			<td width="100px;"></td>
			<td>12</td>
			<td>Note Receivable- Long Term</td>
			<td>
			<input type="text" class="form-control" name="notereceivablelongterm" id="notereceivablelongterm" value="" ></td>	
			</tr>

			<tr>
			<td>13</td>
			<td>Other Long Term Assets</td>
			<td>
			<input type="text" class="form-control" name="otherlongtermassets" id="otherlongtermassets" value=""  />
			</td>
			<td width="100px;"></td>
			<td>14</td>
			<td>Total Long Term Assets</td>
			<td><input type="text" class="form-control" name="totallongtermassets" id="totallongtermassets" value=""  />
			</td>
			</tr>
			
			<tr>
			<td>15</td>
			<td>Total Assets</td>
			<td><input type="text"  class="form-control" name="totalassets" id="totalassets" value=""  >
			</td>
			
			<td colspan="4"><b>LIABILITIES</b></td></tr>
			<td width="100px;"></td>
	
			<tr>
			<td>16</td>
			<td>Account Payable </td>
			<td><input type="text" class="form-control" name="accountpayable" id="accountpayable" value=""  >
			</td>
			</tr>
			
			<tr>
			<td>17</td>
			<td>Payable/Accrued</td>
			<td><input type="text" class="form-control"  name="payableoraccrued" id="payableoraccrued" value="">
			</td>
			<td width="100px;"></td>
			<td>18</td>
			<td>Accrued Expenses</td>
			<td><input type="text" class="form-control" name="accruedexpenses" id="accruedexpenses" value=""  >
			</td>
			</tr>
			
			<tr>
			<td>19</td>
			<td>Notes Payable/Short1 </td>
			<td><input type="text"  class="form-control" name="notespayableshort1" id="notespayableshort1" value="" >
			</td>
			<td width="100px;"></td>
			<td>20</td>
			<td> Current Port. of LT Debt/Capital Leases</td>
			<td><input type="text"  class="form-control" name="currentportofltdebt" id="currentportofltdept" value="" >
			</td>
			</tr>
			
			<tr>
			<td>21</td>
			<td>Other Current Liabilities </td>
			<td><input type="text" class="form-control" name="othercurrentliabilities" id="othercurrentliabilities" value="">
			</td>
			<td width="100px;"></td>
			<td>22</td>
			<td>Total Current Liabilities </td>
			<td><input type="text" class="form-control" name="totalcurrentliabilities" id="totalcurrentliabilities" value=""  >
			</td>
			</tr>
			
			<tr>
			<td>23</td>
			<td>Long Term Debt </td>
			<td><input type="text" class="form-control"  name="longtermdebt" id="longtermdebt" value=""  >
			</td>
			<td width="100px;"></td>
	
			<td>24</td>
			<td>Deferred Income Tax </td>
			<td><input type="text" class="form-control" name="deferredincometax" id="deferredincometax" value="" >
			</td>
			</tr>
			
			<tr>
			<td>25</td>
			<td>Minority Interest  </td>
			<td><input type="text" class="form-control"  name="minorityinterest" id="minorityinterest" value=""  >
			</td>
			<td width="100px;"></td>
			<td>26</td>
			<td>Other Liabilities </td>
			<td><input type="text" class="form-control"  name="otherliabilities" id="otherliabilities" value=""  >
			</td>
			</tr>

			<tr>
			<td>27</td>
			<td>Long-Term Liabilities  </td>
			<td><input type="text" class="form-control"  name="longtermliabilities" id="longtermliabilities" value=""  >
			</td>
			<td width="100px;"></td>
			<td>28</td>
			<td>Total Liabilities  </td>
			<td><input type="text" class="form-control" name="totalliabilities" id="totalliabilities" value=""  >
			</td>
			</tr>

			
			<tr>
			<td colspan="4"><b>EQUITY</b></td>
			</tr>
			
			<tr>
			<td>29</td>
			<td> Redeemable Preferred Stock</td>
			<td><input type="text" class="form-control"  name="redemprefstock" id="redemprefstock" value=""  >
			</td>
			<td width="100px;"></td>
			<td>30</td>
			<td>Preferred Stock - Non Redeemable</td>
			<td><input type="text" class="form-control" name="prefstknonredem" id="prefstknonredem" value=""  >
			</td>
			</tr>
		
			<tr>
			<td>31</td>
			<td> Common Stock</td>
			<td><input type="text" class="form-control" name="cmnstock" id="cmnstock" value=""  >
			</td>
			<td width="100px;"></td>
			<td>32</td>
			<td>Retained Earnings (Accumulated Deficit)</td>
			<td><input type="text"  class="form-control" name="retearn" id="retearn" value=""  >
			</td>
			
			<tr>
			<td>33</td>
			<td> Treasury Stock - Common</td>
			<td><input type="text" class="form-control" name="treasstk" id="treasstk" value=""  >
			</td>
			<td width="100px;"></td>
			<td>34</td>
			<td>Unrealized Gain (Loss)</td>
			<td><input type="text" class="form-control" name="unrealgain" id="unrealgain" value=""  >
			</td>
			</tr>
			
			<tr>
			<td>35</td>
			<td>Other Equity</td>
			<td><input type="text" class="form-control" name="othequity" id="othequity" value=""  >
			</td>
			<td width="100px;"></td>
			<td>36</td>
			<td>Total Equity</td>
			<td><input type="text" class="form-control" name="totequity" id="totequity" value=""  >
			</td>
			</tr>
			
			<tr>
			<td>37</td>
			<td> Liabilities & Equity</td>
			<td><input type="text" class="form-control" name="liabequity" id="liabequity" value=""  >
			</td>
			<td width="100px;"></td>
			<td>38</td>
			<td> Total Common Shares Outstanding</td>
			<td><input type="text"  class="form-control" name="totcomsharesout" id="totcomsharesout" value=""  >
			</td>
			</tr>

			<tr>
			<td>39</td>
			<td>Total Preferred Shares Outstanding</td>
			<td><input type="text" class="form-control" name="totprefsharesout" id="totprefharesout" value=""  >
			</td>
			</tr>
	
			<tr>
			<td></td>
			<td  align="center"><input  style="border-radius:5px;height:25px;" type="submit" value="Save" ></td>
			<td><input type="button" value="Cancel" class="Default" style="border-radius:5px;height:25px;" ></td>
			<td></td>
			</tr>
	
			</table>
			
			</div>

	</form>	
	</div>

			

			
			
			
			
			
			
			
			
			
</div>

</body>
</html>



