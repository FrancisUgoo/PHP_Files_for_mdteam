<?php


class great {
    
 //use $great = new great();    never add include "controllers/great.php"; 
    function index ($arg_for_nothing ) {
      
       $this->view = new view();
       
       $from_where= from_where_now;
        $this->view->requiring($from_where.'error/index'); exit;  //
        
    } 

    
        
     
   public function check_long_id_for_organization($args) {
       session::init();
       $from_where= from_where_now;
       $heredb =  $from_where.'database';
       
        $this->allconn= new $heredb();
        $this->ash = new hashing(); 
        $this->view = new view();
        $this->val = new validation();
        
       $args = rtrim($args,']');     $args= explode(']', $args); 
         if (isset  ($args[1]) ) {
            
              ////  $arg0=organization long id $arg1=organization url $arg2=$adminloginid----   optional
           $args0=$args[0];  $args1=$args[1]; 
           //organization,com to mari, satt and others
           
           $organizationsesdef_id =  $from_where.'sesdef_id';
           $sessionorganizationlongid =  'session'.$from_where.'longid';
           $organizationsesdef_webname =  $from_where.'sesdef_webname';
           $sessionorganizationkeyid =  'session'.$from_where.'keyid';
           $omacccomps =  'omacc'.$from_where.'ps';
           $company_loginid =  $from_where.'_loginid';
           $company_url =  $from_where.'_url';
           
           
           $sessionorganizationurll =  'session'.$from_where.'url';
          if  ((isset ($_COOKIE[urlo.$sessionorganizationurll]))and (!isset ($_SESSION[urlo.$organizationsesdef_webname])))   {
          $mpssurl = session::get($sessionorganizationurll); session::set($organizationsesdef_webname, $mpssurl); } 
          
          if  ((isset ($_COOKIE[urlo.$sessionorganizationlongid]))and (!isset ($_SESSION[urlo.$organizationsesdef_id])) )   {
          $mpsskey = session::get($sessionorganizationlongid); session::set($organizationsesdef_id, $mpsskey); } 
          
          if ( (isset ($_SESSION[urlo.$organizationsesdef_id]))  and (isset ($_COOKIE[urlo.$sessionorganizationlongid]))  and (isset ($_SESSION[urlo.$organizationsesdef_webname])) and (isset ($_COOKIE[urlo.$sessionorganizationkeyid]))  and (isset ($_COOKIE[urlo.$omacccomps]))   ) {
          
            
            $organization_long_id1=session::get($organizationsesdef_id);    $organization_long_id2=session::get($sessionorganizationlongid); // 2 different organization long id
            $organization_web_name1=session::get($organizationsesdef_webname);     //organization url
            
            $his_long_id1=session::get($sessionorganizationkeyid);  //my id
            
            $now_incom_id =session::get($omacccomps);  // present id
            
            $unique_trans =  'unique'.$from_where.'trans';
            
            $now_com_id = $this->ash->$unique_trans("$now_incom_id");  
         } 
          
         else {
             header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); 
             exit;
         }
            
            
            if (isset  ($args[2]) ) {
            $arg_for_id = $args[2]; 
            }
            else {
              $his_id = session::get($sessionorganizationkeyid);
              $arg_for_id = $his_id; 
            }
            
            if ($arg_for_id  <> $his_long_id1)
            {
           header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); exit;
            }
            
            else {
            
         //only member_loginid used bcos regardless of your organization , member_loginid this unique
          $prer = $this->allconn->prepare("SELECT * FROM `login_unique` WHERE `id1` = '$his_long_id1' and `id2` = '$now_com_id' ");
          $prer->execute();   $adminv= $prer->fetch();  $prer->setfetchmode(2); 
         $rows=$prer->rowcount();
        
          if($rows>0){
              
          $dbenc_nowloginid= $adminv['id2'];
          $dbfrom_loginid= $adminv[$company_loginid];
          $dbfrom_url= $adminv[$company_url];     
          $active_status= $adminv['status'];  
          
           
            if (   ($dbfrom_url  <> $args1) or ($dbfrom_url  <> $organization_web_name1) or   ($dbfrom_loginid  <> $args0) or ($organization_long_id1  <> $organization_long_id2) or ($dbfrom_loginid  <> $organization_long_id1)  or ($dbenc_nowloginid  <> $now_com_id)   )
            {
            header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); exit;
            }
            
            else If (($active_status  == 'logedout')) {
            header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); exit;  
            }
            
            else If (($active_status  == 'deactivated')) {
            header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); exit;  
            }
          else { return $his_long_id1;  }
      
        }  else {
            header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); exit;
            }     }
         
          } 
         
         else { 
             $this->view->requiring($from_where.'error/no_arguments/index');
             exit();
         }
    } 
   
 //mimi create the view/marierror folders and subfolders for new ones   
   
    
    
   public function check_roles_for_organization($args) {
              $from_where= from_where_now;
              $heredb =  $from_where.'database';
       
              $this->allconn= new $heredb();
              $this->ash = new hashing(); 
              $this->view = new view();
              $this->val = new validation();
              
      $args = rtrim($args,']');     $args= explode(']', $args); 
         if (isset  ($args[3]) ) {

    
              ////  $arg0=organization long id $arg1=organization url $arg2=$adminloginid from check id returned $arg3=roles seperated with (-)
           $args0=$args[0];  $args1=$args[1];  $arg_for_id = $args[2];   $args3=$args[3];
       
                   if ( $args3 == 'lovingJesusMyGOD__' )
       {
          $this->view->requiring($from_where.'error/index'); //page does not exist
             exit(); 
       }
       else {  
           
            $args3 = rtrim($args3,'-');     $args3= explode('-', $args3); 
          
             $company_logins =  $from_where.'_logins';
             $company_personal_loginid =  $from_where.'_personal_loginid';
             $company_loginid =  $from_where.'_loginid';
             $company_url =  $from_where.'_url';
             $company_personal_role =  $from_where.'_personal_role';
          
          $prer = $this->allconn->prepare("SELECT * FROM `$company_logins` WHERE `$company_personal_loginid` = '$arg_for_id' and  `$company_loginid` = '$args0' and `$company_url` = '$args1' ");
          $prer->execute();   $adminrole= $prer->fetch();  $prer->setfetchmode(2); 
          $rows=$prer->rowcount();
          
          if($rows>0){
              
              $valued = 'no';
      
      foreach ($args3 as $key =>$value)
      {  
        if ( $adminrole[$company_personal_role] == 'MAIN_ADMIN'  
                or $adminrole[$company_personal_role]  == $value    )
            { 
             $valued = 'yes';
            }   
      }
   
       if ( $valued <> 'yes' )
       {
          $this->view->requiring($from_where.'error/index'); //page does not exist
             exit(); 
       }
      
        else if ( $valued == 'yes' )
       {
         return 'yes';
       }
       }  else {
           header("location:".PAT.$from_where."_login/render_check_".$from_where."_admin_in/$args0]$args1]all"); exit; 
         }   } } 
         
         else { 
             $this->view->requiring($from_where.'error/no_arguments/index');
             exit();
         }
      
   }    
  
   
   
   
    
   //---------for satt---------------*******************-------------------------------------------------------------
   
  
     
   public function check_long_id_for_satt($args) {
       session::init();
        $this->sattconn= new sattdatabase();
        $this->ash = new hashing(); 
        $this->view = new view();
        $this->val = new validation();
        
       $args = rtrim($args,']');     $args= explode(']', $args); 
         if (isset  ($args[1]) ) {   
            //  $arg0=nothing  $arg1=$adminloginid----   optional
      
          
           if (  (isset ($_SESSION['sessionsattkeyid']))  and (isset ($_SESSION['omaccsattps']))   ) {
          
            
            $his_long_id1=session::get('sessionsattkeyid');  //my id
            
            $now_incom_id =session::get('omaccsattps');  // present id
            $now_com_id = $this->ash->uniquesatttrans("$now_incom_id");  
         } 
          
         else {
             header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;
         }
            
            
            if (isset  ($args[1]) ) {
            $arg_for_id = $args[1]; 
            }
            else {
              $his_id = session::get('sessionsattkeyid');
              $arg_for_id = $his_id;  //echo 2222; die; 
            }
            
            if ($arg_for_id  <> $his_long_id1) { //echo $arg_for_id; echo '-----------'; echo $his_long_id1; die; 
            header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;
            }   
            else {
            
         //only member_loginid used bcos regardless of your satt , member_loginid this unique
          $prer = $this->sattconn->prepare("SELECT * FROM `login_unique` WHERE `id1` = '$his_long_id1' and `id2` = '$now_com_id' ");
          $prer->execute();   $adminv= $prer->fetch();  $prer->setfetchmode(2); 
         $rows=$prer->rowcount();
        
          if($rows>0){
              
          $dbenc_nowloginid= $adminv['id2'];
          $active_status= $adminv['status'];  
          
           
            if (  ($dbenc_nowloginid  <> $now_com_id)) { 
                header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;
            }
            
            
            else If (($active_status  == 'logedout')) {
              header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;  
            }
            
            else If (($active_status  == 'deactivated')) {
         header("location:".PAT."satt_login/render_check_satt_admin_in/all");
          exit;  
            }
          else { return $his_long_id1;  }
      
        }  else {
             header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;
        }     
   }
   
   } else {header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;}
    } 
   
   
   
   
   public function check_roles_for_satt($args) {
              $this->sattconn= new sattdatabase();
              $this->ash = new hashing(); 
              $this->view = new view();
              $this->val = new validation();
              
      $args = rtrim($args,']');     $args= explode(']', $args); 
         if (isset  ($args[1]) ) {

              //// $arg0=$adminloginid from check id returned $arg1=roles seperated with (-)
           $arg_for_id = $args[0];   $args1=$args[1];
       
            $args1 = rtrim($args1,'-');     $args1= explode('-', $args1); 
          
          
          $prer = $this->sattconn->prepare("SELECT * FROM `satt_logins` WHERE `satt_personal_loginid` = '$arg_for_id'  ");
          $prer->execute();   $adminrole= $prer->fetch();  $prer->setfetchmode(2); 
          $rows=$prer->rowcount();
          
          if($rows>0){
              $valued = 'no';
      
      foreach ($args1 as $key =>$value)
      {  
        if ( $adminrole['satt_personal_role'] == 'MAIN_ADMIN'  
                or $adminrole['satt_personal_role']  == $value    )
            { 
             $valued = 'yes';
            }   
      }
   
       if ( $valued <> 'yes' )
       {
          $this->view->requiring('satterror/index'); //page does not exist
             exit(); 
       }
      
        else if ( $valued == 'yes' )
       {
         return 'yes';
       }
       }  else {
             header("location:".PAT."satt_login/render_check_satt_admin_in/all"); 
             exit;
        }    } 
         
         else { 
             $this->view->requiring('satterror/no_arguments/index');
             exit();
         }
      
   }    
       
    
   //------------End for satt-----------***************-------------------------*****************----------------------------
       
  /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  
       //------------End for Organization-----------***************-----------------------------------------------
   
    
   
    public function get_the_proper_pat_sat($args) {
     $args1="http://satt.websatt.com/";   return $args1;
 } 
   public function get_the_proper_pat($args) {
     $args1="http://$args.websatt.com/";   return $args1;
 }    
   public function get_the_proper_patn($args) { //localhost/satt to be replaced with the ccnga.com
     $args1="www.$args.websatt.com";   return $args1;
 }  
  public function get_the_proper_patnn($args) { //localhost/satt to be replaced with the ccnga.com
     $args1="$args.websatt.com";   return $args1;
 } 
   
   
    
    
    
    }

