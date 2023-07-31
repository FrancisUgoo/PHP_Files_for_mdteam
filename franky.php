
<?php


class franky
{
    
function __construct()
{     
    
 $urla0= url0; $urla1= url1; $urla2= url2; $urla3= url3; $urla4= url4;   
 $fromm = from_where_now;
    
 if (empty($urla0))
{
    require 'sattcontrollers/home.php';  //main satt home
    $vie=new home(); $vie->index($urla0); return false;
    //exit;
}  
    

else if (($urla0=='satt')|| ($fromm=='satt'))   //same for company
{
    
  // define('from_where_now', "satt");  
    
  if ($urla1 != '_____')
{ 
$ofile =from_where_now.'controllers/'.$urla1.'.php'; 
if(file_exists($ofile))
{
    require $ofile ;   //same as require_once or include
    $newpag = new $urla1; //class of controller
} 
else 
{
    include from_where_now.'controllers/error.php';
    $this->viewerr=new errot();
    $this->viewerr->index($urla1);
   // return false;   //if this return false wasnt specified then other other errors will appear below in browser
    exit;
}
}
else {
    require from_where_now.'controllers/home.php';
    $vie=new home();
     $vie->index('no_ps');
     return false;
} 
 

 if ($urla4 != '_____') 
{
    if($urla1==='error') {
        require 'view/'.from_where_now.'error/index.php';   exit;  
    }else{
          require from_where_now.'controllers/error.php';   //this error page will contains details of the issue
          $this->viewerr=new errot();
          $this->viewerr->forbidden($urla1)  ;    //it must be index may another function to take to different error page
          exit();
}
}

if ($urla3 != '_____') 
{
    if($urla1==='error') {
       require 'view/'.from_where_now.'error/index.php';   exit;  
    }else{
           $method=$urla2;
    if (method_exists($newpag,$method ))//check if method exist
           {
          $newpag->{$urla2}($urla3);
               return false;
  }  else {
    require from_where_now.'controllers/error.php';
            $this->viewerr=new errot();
            $this->viewerr->index($urla1); return false;   //if this return false wasnt specified then other other errors will appear below
           }
   }
}

if ($urla2 != '_____')
{
   if($urla1==='error') {
       require 'view/'.from_where_now.'error/index.php';  exit;
    }else{
           $method=$urla2;
 
    if ($urla3 == '_____')    
{
     require from_where_now.'controllers/error.php';   //this error page will contains details of the issue
              $this->viewerr=new errot();
               $this->viewerr->no_arguments($urla1)  ;    //it must be index may another function to take to different error page
               return false;
           }              
   else 
            {
      //if this return false wasnt specified then other other errors will appear below
            }
   }
}
else
{}  //check effect of this  $$$$  
  
if ($urla1 != '_____')
{
$newpag->index($urla1);
return false;
}
 else {
     require from_where_now.'controllers/error.php';
              $this->viewerr=new errot();
               $this->viewerr->index($urla1)  ;  
               return false;
}

require from_where_now.'controllers/home.php';
    $vie=new home();
     $vie->index('no_ps');
    return false;
    
}   
    
//-------------------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------------------

else {
 
define('DB_TYPE_checking', 'mysql');      //DEFINING THEM AS CONSTANTS.
define('DB_HOST_checking', 'localhost');

define('DB_USER_checking', 'ii5665u344hfh757hhfgzfbZZZerreuuuYYY74647_');
define('DB_PASS_checking', 'gfgtf77363***ZZZTT54dfdcv#uuuASECVG769876544GTGB');   
//-----------------------------------------------------------------------------------------------------------------------
   
  if ($urla1 != '_____')
{
    $ofile =$fromm.'controllers/'.$urla1.'.php';
  if(file_exists($ofile))
{
    require $ofile ;  $newpag = new $urla1; //class of controller
} 
else 
{
    include $fromm.'controllers/error.php';  $this->viewerr=new errot();
    $this->viewerr->index($urla1);   exit;
   // return false;   //if this return false wasnt specified then other other errors will appear below in browser
}
}
else {
    require $fromm.'controllers/home.php';
    $vie=new home();  $vie->index('no_ps');  return false;
} 


 if ($urla4 != '_____') 
{
    if($urla1==='error') {
        require 'view/'.$fromm.'error/index.php';   exit;  
    }else{
          require $fromm.'controllers/error.php';   //this error page will contains details of the issue
          $this->viewerr=new errot();
          $this->viewerr->forbidden($urla1)  ;    //it must be index may another function to take to different error page
          exit();
}
}

if ($urla3 != '_____') 
{
    if($urla1==='error') {
       require 'view/'.$fromm.'error/index.php';   exit;  
    }else{
           $method = $urla2;
    if (method_exists($newpag,$method ))//check if method exist
           {
          $newpag->{$urla2}($urla3);
               return false;
  }  else {
    require $fromm.'controllers/error.php';
            $this->viewerr = new errot();
            $this->viewerr->index($urla1); return false;   //if this return false wasnt specified then other other errors will appear below
           }
   }
}


if ($urla2 != '_____')
{
   if($urla1==='error') {
       require 'view/'.$fromm.'error/index.php';  exit;
    }else{
           $method = $urla2;
 
    if ($urla3 == '_____')    
{
     require $fromm.'controllers/error.php';   //this error page will contains details of the issue
              $this->viewerr=new errot();
               $this->viewerr->no_arguments($urla1)  ;    //it must be index may another function to take to different error page
               return false;
           }              
   else 
            {
      //if this return false wasnt specified then other other errors will appear below
            }
   }
}
else
{}

if ($urla1 != '_____')
{
$newpag->index($urla1);
return false;
}
 else {
     require $fromm.'controllers/error.php';
              $this->viewerr=new errot();
               $this->viewerr->index($urla1)  ;  
               return false;
}
//////////////////////

require $fromm.'controllers/home.php'; //main home page
    $vie=new home();
     $vie->index('no_ps');
    return false;


    
}


    
    
    
}    
}