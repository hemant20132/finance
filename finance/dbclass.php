<?php
class db
{

protected $dbuser="root";
protected $dbpass="";
protected $dbserver="localhost";
protected $dbdatabase="test";

public function db_connect()
{

$connect=mysqli_connect($this->dbserver,$this->dbuser,$this->dbpass,$this->dbdatabase) ;

	if ( mysqli_connect_errno() ) 
	{
	echo "error in connection";
	exit();
	}
	return true;
}


public function db_insert($name,$address)
{
$name=$this->name;
$address=$this->address;


$query1= "insert into table1 (name,address) values ('" . $this->name. "','" . $this->address . "')";

	if (mysqli_query($query1,$connect))
	{
	echo "one record inserted successfully";
	exit();
	}

}





}
?>