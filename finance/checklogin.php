<?php
		ob_start();
		$loginname=$_POST['loginname']; 
		$loginpass=$_POST['loginpass'];
	
		$conn=mysqli_connect('sql211.epizy.com','epiz_33239466','Horizon2022','epiz_33239466_new','3306','') or  die ("Error in query:".mysqli_error());
		$sqlquery = "select loginname, loginpass from userlogin where loginname='" .$loginname. "' and loginpass ='".$loginpass."'"; 
		$result=mysqli_query($conn,$sqlquery);
		$info=mysqli_fetch_array($result,MYSQLI_ASSOC);

		if ($info['loginname'] == $loginname and $info['loginpass'] == $loginpass)
		{					
			header("location:mainpage.php?loginsuccess=yes&loginname=".$loginname);
		}
		elseif ($info['loginname'] != $loginname or $info['loginpass'] != $loginpass )
		{
			header("location:index.php?loginsuccess=no");
			
		}		

?>
			