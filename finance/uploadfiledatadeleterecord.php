<?php
$compnayupload =$_GET['companyupload'];
$inputfilename =$_GET['inputfilename'];


$conn = mysql_connect("localhost","nacheeso_hemant","7ByV+sA[CsL6") or die(mysql_error());
mysql_select_db("nacheeso_finance");

$queryupload= "DELETE FROM uploadfiledata WHERE companyname=". $companyupload . " and uploadfilename ='" . $inputfilename ."'" ;
mysql_query($queryupload,$conn);

header ("location:uploadfiledatadelete.php");


?>
