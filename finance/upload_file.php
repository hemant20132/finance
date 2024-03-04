<?php 
error_reporting(E_ALL);
 $companyname=$_REQUEST['companyname'];
 $target = "Excel/uploads/"; 
 $target = $target.basename( $_FILES['file']['name']); 
 $ok=1; 
 if(move_uploaded_file($_FILES['file']['tmp_name'], $target)) 
 {
echo "file moved";	 
  $conn=mysqli_connect('sql211.epizy.com','epiz_33239466','Horizon2022','epiz_33239466_new','3306','') or  die ("Error in query:".mysqli_error());
 $query="insert into uploadfiledata (companyname,uploadfilename) values ('" .$companyname. "','" .$_FILES['file']['name']. "')";
 mysqli_query($conn,$query);
 $uploaded="Successfully Uploaded excel file";
 header ('location:uploadexcelfile.php?uploaded='.$uploaded);
 } 
 else 
 {
 $uploaded="Sorry there was a problem uploading your file";
 header ('location:uploadexcelfile.php?uploaded='.$uploaded);
 }
 
 ?> 




