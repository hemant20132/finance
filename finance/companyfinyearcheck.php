<?php
$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
mysql_select_db("nacheeso_finance");
$companyname = $_GET['companyname'];
$finyear 	= $_GET['finyear'];
$query1 = "select finyear, company from incomestatement where finyear=". intval($finyear) . "and company='" . $companyname . "'";

$resultfound1= mysql_query($query1,$conn);

$infofound1 = mysql_fetch_array($resultfound1,MYSQL_ASSOC);

if ( $infofound1['finyear']== intval($finyear) and $infofound1['company'] == $companyname )
{
echo "The incomestatement for company". $companyname . "for financial year" .$finyear."already exists";
}
?>