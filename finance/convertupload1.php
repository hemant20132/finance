<?php
session_start(); 
ob_start();
$companyname = $_GET['companyname'];
$inputfilename =  $_GET['inputfile'];
$_SESSION['companyname']=$companyname;
$_SESSION['inputfilename']=$inputfilename;
$conn=mysqli_connect('sql211.epizy.com','epiz_33239466','Horizon2022','epiz_33239466_new','3306','') or  die ("Error in query:".mysqli_error());
$query1 = "update uploadfiledata set dataposted='YES' where companyname='". $companyname . "'";
mysqli_query($conn,$query1);
header ("location:convertuploadedfile.php?companyname=". $_SESSION['companyname'] ."&inputfile=".$_SESSION['inputfilename']);
?>
