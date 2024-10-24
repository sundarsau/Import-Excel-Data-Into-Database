<?php

$conn = new mysqli("localhost","root","","test");

if ($conn->connect_error){
    die("DB connection failed ". $conn->connect_error);
}