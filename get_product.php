<?php
require_once('./db_connect.php');
$search_input = 'Super';
$query = "SELECT * FROM `product_name`";
$search_input = 'Huda';
$query = "SELECT * FROM `product_name` WHERE `Product_Name` LIKE '%$search_input%'";
//will spit out data from the Product Landing Page wire frame (not the ingredients)
//product_adding is only for making things in associative table
//SELECT * FROM `ingredients_rating` where `Ingredient`like '%username%'
//produce an array
//feed 1 by 1 through a loop to query -> a result with everything
//array your going to make is going to start with a number 
//[1,bunch of shit]
//when inserted array[0], loop through bunch of shit
//create insert function dynamically 
$result=mysqli_query($db,$query);
$output=[];
$output['success']=false;
if(mysqli_num_rows($result)){
    $row=mysqli_fetch_assoc($result);
    $output['success']=true;
    $output['products']=$row; 
} else {
    $output['error']='Can\'t find product';
}
echo 'json encoded: '.json_encode($output);
?>