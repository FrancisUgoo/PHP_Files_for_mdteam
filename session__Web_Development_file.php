<?php

class session
{
    
    private static $_sessionstarted= false;
    
   public static function init()
   {
      if (self:: $_sessionstarted==false)  //self::$xx replace £this->xxx  because it is static
      {
          //session_start(['cookie_lifetime' => 86400, 'read_and_close' => true,]); 
          //session_start(['cookie_lifetime' => 86400000,]);    //1 day = 86400
          
          session_start(['cookie_lifetime' => 86400000000000, 'cache_expire' => 86400000000000, ]);    //1 day = 86400
          
        self:: $_sessionstarted=true;     //giving the variable =true whenever session starts.
      }
 else {
         self:: $_sessionstarted=true;
      }
   }
   
   public static function secot($key,$value)
   {
       $urlnw = urlo;  $keym = "$urlnw"."$key"; //echo $keym; echo '  ';
       $_SESSION[$key]=$value;       //setcookie($key, $value, time()+25000000000);  //setcookie($key, $value, time()+60*60*24*300000);
       $_SESSION[$keym]=$value;      setcookie($keym, $value, time()+25000000000,"/","websatt.com");  //setcookie($keym, $value, time()+25000000000,"/","websatt.com"); 
 //echo $_COOKIE[$keym]; 
       }
       
   public static function secotta($key,$value)
   {
       $urlnw = urlo;  $keym = "$urlnw"."$key"; //echo $keym; echo '  ';
       $_SESSION[$key]=$value;       //setcookie($key, $value, time()+25000000000);  //setcookie($key, $value, time()+60*60*24*300000);
       $_SESSION[$keym]=$value;      setcookie($keym, $value, time()+3600,"/","websatt.com");  //setcookie($keym, $value, time()+25000000000,"/","websatt.com"); 
 //echo $_COOKIE[$keym]; 
       }     
   
   public static function cosetaap($key,$value)
   {
       setcookie($key, $value, time()+60*60*24*40,"/","websatt.com");  //setcookie($keym, $value, time()+25000000000,"/","websatt.com"); 
 //echo $_COOKIE[$keym]; 
       }
   
  
   
   public static function set($key,$value)
   {
       $urlnw = urlo;  $keym = "$urlnw"."$key";
       $_SESSION[$key]=$value;
       $_SESSION[$keym]=$value;
   }
   
    public static function setaap($key,$value)
   {
       $_SESSION[$key]=$value;
   }
   
   
    public static function getaap($key) //second key used when we want to make $_session an array 
   {   
    if (isset ($_COOKIE[$key]))  { $_SESSION[$key]=$_COOKIE[$key]; return $_SESSION[$key];}
     else { if (isset ($_SESSION[$key]))    //to avoid error of undefined variable when $_session[$key] is empty
            {  return $_SESSION[$key];}
       else {  return 'nopee'; }
   } }
   
   public static function getaa($key) //second key used when we want to make $_session an array 
   {   
       $urlnw = urlo;  $keym = "$urlnw"."$key";
       
   if (isset ($_COOKIE[$keym]))  { $_SESSION[$keym]=$_COOKIE[$keym]; return $_SESSION[$keym];}     
  else {     if (isset ($_SESSION[$keym]))    //to avoid error of undefined variable when $_session[$key] is empty
           {  return $_SESSION[$keym];}
           else { 
            return 11;
             }
       return false;
} }
   
   
   
   public static function get($key, $secondkey=false) //second key used when we want to make $_session an array 
   {       
       $urlnw = urlo;  $keym = "$urlnw"."$key"; 
        $tttmariconn= new maridatabase();
        $tttsattconn= new sattdatabase();
        
         //mimi add new db for new services
         $tttorgconn= new orgdatabase();
         $ttteventconn= new eventdatabase();
        
        $tttash = new hashing(); 
        $tttview = new view();
        $tttval = new validation();
       //second key is one of the children of $key  for example as below
       //sess::set('name', array( 'firstname'=>'ugochukwu', 'middlename' =>'Francis', 'fathersname'=>'obi'));
       //I can now get the value of middle name alone. from   // echo sess::get('name','middlename');
       if ($secondkey==true)
       {
           if (isset($_SESSION[$keym][$secondkey]))
           { return $_SESSION[$keym][$secondkey]; }
          else{  $sector=from_where_now;  return false;}
       }
       
       
       else      
       {
           if (isset ($_COOKIE[$keym]))  { $_SESSION[$keym]=$_COOKIE[$keym]; return $_SESSION[$keym];} 
           else if (isset ($_SESSION[$keym])) {  return $_SESSION[$keym];}
           else{  $sector=from_where_now;  return false;}
       }
       
       return false;
   }
   
   
   
   
   public static function display()
   {
       if (!empty($_SESSION))   //to avoid error of undefined variable when $_session is empty
      {
       echo '<pre>';
       print_r($_SESSION);
       echo '<pre>';
      }
 else 
           
 {}
   }

      public static function destroy()
   {
       if (self:: $_sessionstarted==true)
      {
           session_destroy(); 
          
        self:: $_sessionstarted=false;    //remember check '=' assigment one = but comparing with if 2 or 3 ==
      }
      else {
      return false;
      }
   }
   
   
   
}

