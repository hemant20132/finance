<?php
include 'dbclass.php';

$dbconnect1= new db;

if ($dbconnect1->db_connect())
{
echo "successful connection now with database";
}

$dbconnect1->db_insert('hemant','desai');



?>