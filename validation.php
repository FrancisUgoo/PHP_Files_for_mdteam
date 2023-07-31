<?php

class validation {
    
  public $empty = "";
  function __construct() {}
    
 function hashnlinkk($password) {     
    $urlao= urlo; $linkpassword1= strtoupper("$password"); $linkpassword = substr($linkpassword1, 0, 60);
    $keym3 = "$urlao"."HStze5467rtw223rTEvc"."$urlao"; //that is salt    
       $hashing_element= hash_init('sha256', HASH_HMAC, $keym3);
       hash_update($hashing_element, $linkpassword); return hash_final($hashing_element);
    }  
    
 function checkhashnlinkk($password) {
    $urlao= urlo; $linkpassword1= strtoupper("$password"); $linkpassword = substr($linkpassword1, 0, 60);
    $keym3 = "$urlao"."HStze5467rtw223rTEvc"."$urlao"; //that is salt    
       $hashing_element= hash_init('sha256', HASH_HMAC, $keym3);
       hash_update($hashing_element, $linkpassword); return hash_final($hashing_element);
    }     
  
  
function checklinka($inputa) {
 $capinputa = strtoupper("$inputa");
 $capinputa= str_replace(' ', '', $capinputa);  
 $httpp= substr($capinputa,0,4);
 $httppos = strpos($capinputa, 'HTTP://'); $httpspos1 = strpos($capinputa, 'HTTPS://'); $lenlink = strlen($capinputa);

     if ($lenlink <= 0) { return 'good';}
else if (($lenlink > 4) and ($httpp != "HTTP")) { return 'bad';}     
else if (($lenlink > 0) and ($httppos === false) and ($httpspos1 === false)) { return 'bad';} 
else if (($lenlink > 0) and ($lenlink < 11)) { return 'bad';}
else { return 'good';}
}
 
 // test this guy too 
    function purifyinput($input)  {  $input= filter_var($input, FILTER_SANITIZE_STRING);
   
    $input= str_replace("'", "", $input); $input= str_replace(";", "", $input);  $input= str_replace('<', '', $input); //remove <
    $input= str_replace('>', '', $input); $input= str_replace('"', '', $input);  $input= str_replace('`', '', $input); //remove "
    $input= str_replace('*', '', $input);   //remove "
    
    $input= str_replace('--', '', $input);//every comment remove it
    $input= str_replace('=', 'is equal to', $input); //Equal to
    $input= str_replace('+', ' ', $input);

    $input=trim($input); $input=  htmlspecialchars($input); return $input;
    
}
   
    

function testinput($input)  {  $input= filter_var($input, FILTER_SANITIZE_STRING);
  
    if ( (empty($input)) or ($input == '') or ($input == ' ') ) {
        $this->empty ="Hello, this field is empty"; }
    
    else {
    $input= str_replace("'", "", $input); $input= str_replace(";", "", $input); $input= str_replace("<", '', $input);   //remove <
    $input= str_replace(">", '', $input); $input= str_replace('"', "", $input); $input= str_replace('`', '', $input);   //remove "
    $input= str_replace('*', '', $input);   //remove "
    
    $input= str_replace('--', '', $input);//every comment remove it
    $input= str_replace('=', 'is equal to', $input); //Equal to
    $input= str_replace('+', ' ', $input);
    
    $input=trim($input); $input=  htmlspecialchars($input); return $input;
    }
}




function testinputwithspace($input)  {  $input= filter_var($input, FILTER_SANITIZE_STRING);
    
    if(empty($input)) { $this->empty ="Hello, this field is empty"; }
    
    else {
   $input= str_replace("'", "", $input); $input= str_replace(";", "", $input);  $input= str_replace('<', '', $input);   //remove <
   $input= str_replace('>', '', $input); $input= str_replace('"', '', $input);  $input= str_replace('`', '', $input);   //remove "
   $input= str_replace('*', '', $input);   //remove "
   
   $input= str_replace('--', '', $input);//every comment remove it
   $input= str_replace('=', 'is equal to', $input); //Equal to
   $input= str_replace('+', ' ', $input);
    
   $input=trim($input); $input=  htmlspecialchars($input); return $input;
    }
    }
    
    
      function purifyvvinput($input)  {  $input= filter_var($input, FILTER_SANITIZE_STRING);
   
    $input= str_replace("'", "", $input); $input= str_replace(";", "", $input);  $input= str_replace('<', '', $input); //remove <
    $input= str_replace('>', '', $input); $input= str_replace('"', '', $input);  $input= str_replace('`', '', $input); //remove "
    $input= str_replace('*', '', $input);   //remove "
    
    $input= str_replace('--', '', $input);//every comment remove it
    $input= str_replace('=', 'is equal to', $input); //Equal to
    $input= str_replace('+', ' ', $input);
    
    $input=trim($input); $input=  htmlspecialchars($input); return $input;
}  
    
    
    function getimagefolder ($imagevalueneeded,$from)  {
 //$from ie from images folder or images_org folder or images_omacc folder
  $args = $imagevalueneeded;   
  $args = rtrim($args,'__');     $args= explode('__', $args); 
 
    if (isset  ($args[1]) ) {
        $fileimagename = $args[0].'/'.$imagevalueneeded; 
        $fileimagepath = $from.'/'.$fileimagename;
   
        return $fileimagepath;
    }
    else {   
       $fileimagepath = 'defaultimages'.'/'.$imagevalueneeded; 
       
      return $fileimagepath; 
    }
   
}  

    
    

function deleteunusedimage ($imagevalueneeded)  {
  $sector=from_where_now;
  //if  ((isset ($_SESSION['sessionchurchfrom'])) and (isset ($_SESSION['testandmoveimagetype']))) //cant find how to replace sessionchurchfrom 
  if  (isset ($_SESSION['testandmoveimagetype']))
  {  $filll = session::get('testandmoveimagetype');}
  else {$filll = 'no';}
   
  $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args); 
  
  
  
  if ($sector == 'satt' or $filll == 'satt') {
      $sattconn= new sattdatabase();
      
         $file = 'images_satt/'.$args[0].'/'.$imagevalueneeded;
          if(file_exists($file))
          {
if (!unlink($file))  { return 'undone';  }
else {
    $pre = $sattconn->prepare("DELETE FROM `satt_images` WHERE `satt_image_new_name`='$imagevalueneeded'");
    $pre->execute(); return 'done';
  }  } else {return 'undone';}
  } 
   
 
  else if ($sector == 'mari' or $filll == 'mari') {
    $mariconn= new maridatabase(); 
    
      $file = 'images_mari/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { return 'undone';}
else {
    $pre = $mariconn->prepare("DELETE FROM `mari_images` WHERE `mari_image_new_name`='$imagevalueneeded'");
    $pre->execute();  return 'done';
     }
      } else {return 'undone';}
  }

 //mimi new else if ($sector == 'mari' or $filll == 'mari') for new services following the mari format and create the corresponding mari_images table 
   
 
  else if ($sector == 'org' or $filll == 'org') {
    $orgconn= new orgdatabase(); 
    
      $file = 'images_org/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { return 'undone';}
else {
    $pre = $orgconn->prepare("DELETE FROM `org_images` WHERE `org_image_new_name`='$imagevalueneeded'");
    $pre->execute();  return 'done';
     }
      } else {return 'undone';}
  }

 
  else if ($sector == 'event' or $filll == 'event') {
    $eventconn= new eventdatabase(); 
    
      $file = 'images_event/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { return 'undone';}
else {
    $pre = $eventconn->prepare("DELETE FROM `event_images` WHERE `event_image_new_name`='$imagevalueneeded'");
    $pre->execute();  return 'done';
     }
      } else {return 'undone';}
  }
  
  
  
  
}    
    
 
function deleteallimages_mari ($arggs)  {
 $arggs = rtrim($arggs,']');   $arggs= explode(']', $arggs); 
          $arggs0=$arggs[0]; //table name 
          $arggs1=$arggs[1]; //pics column name 
          $arggs2=$arggs[2]; //url longid
          $arggs3=$arggs[3]; //url    
    
   $mariconn= new maridatabase();   
  $pre = $mariconn->prepare("SELECT * FROM `$arggs0`  WHERE  `mari_loginid` ='$arggs2' AND `mari_url`='$arggs3' ");
  $pre->execute(); $pre->setfetchmode(2);  $details= $pre->fetchall();   $rows=$pre->rowcount();
      if ($rows >0) {   
    foreach ($details as $key =>$value) {    
     $imagevalueneeded = $value[$arggs1]; 
   $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args);  
    if (isset  ($args[1]) ) { 
      $file = 'images_mari/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { $hhh = 1;}
    $pred = $mariconn->prepare("DELETE FROM `mari_images` WHERE `mari_image_new_name`='$imagevalueneeded'");
    $pred->execute();   
     } }
   } return 'done';
   }  else { return 'done';}
}    

function deleteallimages_event ($arggs)  {
 $arggs = rtrim($arggs,']');   $arggs= explode(']', $arggs); 
          $arggs0=$arggs[0]; //table name 
          $arggs1=$arggs[1]; //pics column name 
          $arggs2=$arggs[2]; //url longid
          $arggs3=$arggs[3]; //url    
    
   $eventconn= new eventdatabase();   
  $pre = $eventconn->prepare("SELECT * FROM `$arggs0`  WHERE  `event_loginid` ='$arggs2' AND `event_url`='$arggs3' ");
  $pre->execute(); $pre->setfetchmode(2);  $details= $pre->fetchall();   $rows=$pre->rowcount();
      if ($rows >0) {   
    foreach ($details as $key =>$value) {    
     $imagevalueneeded = $value[$arggs1]; 
   $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args);  
    if (isset  ($args[1]) ) { 
      $file = 'images_event/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { $hhh = 1;}
    $pred = $eventconn->prepare("DELETE FROM `event_images` WHERE `event_image_new_name`='$imagevalueneeded'");
    $pred->execute();  
     } }
   } return 'done';
   }  else { return 'done';}
}    
 
function deleteallimages_org ($arggs)  {
 $arggs = rtrim($arggs,']');   $arggs= explode(']', $arggs); 
          $arggs0=$arggs[0]; //table name 
          $arggs1=$arggs[1]; //pics column name 
          $arggs2=$arggs[2]; //url longid
          $arggs3=$arggs[3]; //url    
    
   $orgconn= new orgdatabase();   
  $pre = $orgconn->prepare("SELECT * FROM `$arggs0`  WHERE  `org_loginid` ='$arggs2' AND `org_url`='$arggs3' ");
  $pre->execute(); $pre->setfetchmode(2);  $details= $pre->fetchall();   $rows=$pre->rowcount();
      if ($rows >0) {   
    foreach ($details as $key =>$value) {    
     $imagevalueneeded = $value[$arggs1]; 
   $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args);  
    if (isset  ($args[1]) ) { 
      $file = 'images_org/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { $hhh = 1;}
    $pred = $orgconn->prepare("DELETE FROM `org_images` WHERE `org_image_new_name`='$imagevalueneeded'");
    $pred->execute();  
     } }
   } return 'done';
   }  else { return 'done';}
}    
 



 

function addnewfolder ($recentfoldername,$mainfoldername, $from)  {
    //$from for church homepage banner images should omacc
    
   (int)$strlength= strlen($recentfoldername); 
   (int)$oldstrlength= strlen($mainfoldername); 
   
   $newfoldername = $recentfoldername;
   
   if (($strlength <= $oldstrlength) or ("$mainfoldername" == "$recentfoldername")) {
       $newfoldername = "$mainfoldername"."1";
   }
 else {
       $numadd = str_replace("$mainfoldername", "", $recentfoldername);
       $numadd = intval($numadd); $numadde = $numadd + 1;
       $newfoldername = "$mainfoldername"."$numadde";
   }
    
   if ($from == 'satt')  {
      $newfolderpath =  'images_satt/'.$newfoldername ;
      
      if (is_dir($newfolderpath) === false) {
          mkdir($newfolderpath, 0777);          
          return $newfoldername;
      }  
      else {return $newfoldername;}
      
   }
   
 else {
     $newfolderpath =  'images_'.$from.'/'.$newfoldername ;
      if (is_dir($newfolderpath) === false) {
          mkdir($newfolderpath, 0777);
          return $newfoldername;
      }  
      else {return $newfoldername;} 
   }
   //create the new folder 
    // return new folder nameeven if exists already
}
 


    function compress($source, $destination, $quality,$newwidthno,$newheightno) {
$from_where= from_where_now; $this->view = new view();
		$info = getimagesize($source);
                // It gets the size of the image
                list( $width,$height ) = getimagesize( $source );
                // It makes the new image width of 350
                $re_memory = round($width * $height); //putnewindex
                  if ($re_memory > 50000000) {$this->view->requiring($from_where.'error/no_arguments/indexpics');  exit();}
            else {
               
                $newheight = $height; $newwidth = $width;
                 
               if (($height != $newheightno) || ($width != $newwidthno)) {
                 $newheight = $newheightno; $newwidth = $newwidthno;
               }
               
               
          $bacolor = 'white'; $nowurl = urlo;
          if ($newheightno < 150) {
          //add if here mimi
         if ($from_where == 'satt') { $cconn= new sattdatabase(); $pre1 = $cconn->prepare("SELECT `background_color` FROM `satt_details`  WHERE `satt_detailsid` = '10' ");
       $pre1->execute();  $details1= $pre1->fetch(); $bacolor = "$details1[0]"; }
       
        else if ($from_where == 'org') { $cconn= new orgdatabase(); $pre1 = $cconn->prepare("SELECT `background_color` FROM `org_details`  WHERE `org_url` ='$nowurl' ");
       $pre1->execute();  $details1= $pre1->fetch(); $bacolor = "$details1[0]"; }
       
        else if ($from_where == 'mari') { $cconn= new maridatabase(); $pre1 = $cconn->prepare("SELECT `background_color` FROM `mari_details`  WHERE `mari_url` ='$nowurl' ");
       $pre1->execute();  $details1= $pre1->fetch(); $bacolor = "$details1[0]"; }
       
        else if ($from_where == 'event') { $cconn= new eventdatabase(); $pre1 = $cconn->prepare("SELECT `background_color` FROM `event_details`  WHERE `event_url` ='$nowurl' ");
       $pre1->execute();  $details1= $pre1->fetch(); $bacolor = "$details1[0]"; }
            }    
         $ggcc = $this->getrgbvall($bacolor); $ggcc = rtrim($ggcc,']');   $ggcc= explode(']', $ggcc);
          
              $thumb = imagecreatetruecolor( $newwidth, $newheight ); 
               $imgcolour = imagecolorallocate($thumb, $ggcc[0], $ggcc[1], $ggcc[2]);
               imagefill($thumb, 0, 0, $imgcolour);
               imagesavealpha($thumb, true);
   ///////////////////////////////////////////////////////////////////////////////////////////////////////////            
              
                if ($info['mime'] == 'image/jpeg')      {$image = imagecreatefromjpeg($source);}

		elseif ($info['mime'] == 'image/pjpeg') {$image = imagecreatefromgif($source);}
                
                elseif ($info['mime'] == 'image/jpg')   {$image = imagecreatefromgif($source);}
                
                elseif ($info['mime'] == 'image/gif')   {$image = imagecreatefromgif($source);}

		elseif ($info['mime'] == 'image/png')   {$image = imagecreatefrompng($source);}
  ///////////////////////////////////////////////////////////////////////////////////////////////////////////////         
            
           if (!$image) { $this->view->requiring($from_where.'error/no_arguments/indexpics');
             exit();}
           else {
                // Resize
            //imagecopyresized($thumb, $image, 0, 0, 0, 0, $newwidth, $newheight, $width, $height);
            imagecopyresampled($thumb, $image, 0, 0, 0, 0, $newwidth, $newheight, $width, $height);

            
            $exif = exif_read_data($source);
            if ($exif && isset($exif['Orientation'])) {
              $orientation = $exif['Orientation'];
              $deg = 0;
              if ($orientation == 3) {$deg = 180;}
              if ($orientation == 6) {$deg = 270;}
              if ($orientation == 8) {$deg = 90;}
              $thumb = imagerotate($thumb, $deg, 0);
            }
            
            
            // It then save the new image to the location specified by $destination variable
                 imagejpeg($thumb, $destination, $quality);

		return $destination;
           }  } }
          
    

      
    function alltestandmoveimage($image,$folder,$from)  { //added $from
           
 //-------------------------------------------------------------------------------------------------------------------------       
 
 if ($image['type']=='image/jpeg'
             || $image['type']=='image/pjpeg'
              || $image['type']=='image/jpg'
               || $image['type']=='image/JPG'
               || $image['type']=='image/gif'
                   || $image['type']=='image/png'
             
              and $image['size']<1000000000 ) 
           {  
          
         $quality= 90; 
         $largeheight= 700; $largewidth= 900; 
         $bigheight= 700; $bigwidth= 700;  
         $mediumheight= 300; $mediumwidth= 300;
         $smallheight= 127; $smallwidth= 127;
         $speakheight= 300; $speakwidth= 300;
         $galheight= 400; $galwidth= 400;
         $logoheight= 80; $logowidth= 127;
             
         //  move_uploaded_file($image['tmp_name'],$this->images.$image['name']);  //not needed again
         session::init();  
        
        
         $cconn= new sattdatabase();  session::set('testandmoveimagetype', 'satt'); $prebb = 'satt';
          //add if here mimi
         if ($from == 'mari') { $cconn= new maridatabase(); session::set('testandmoveimagetype', 'mari'); $prebb = 'mari';}
         if ($from == 'org') { $cconn= new orgdatabase(); session::set('testandmoveimagetype', 'org'); $prebb = 'org';}
         if ($from == 'event') { $cconn= new eventdatabase(); session::set('testandmoveimagetype', 'event'); $prebb = 'event';}
        
        $ash = new hashing(); 
        $view = new view();
        $filll = session::get('testandmoveimagetype');
         
         $imageid =  $prebb.'_imageid'; $images =  $prebb.'_images'; $image_m_folder =  $prebb.'_image_m_folder';
         $image_folder =  $prebb.'_image_folder'; $image_old_name =  $prebb.'_image_old_name'; $image_new_name =  $prebb.'_image_new_name';
         
         
        
       $prefolder2 = $cconn->prepare("SELECT  MAX(`$imageid`) FROM `$images` WHERE `$image_m_folder`='$folder'  ");
       $prefolder2->execute();    $lastfolder= $prefolder2->fetch();   
       $rowcs= $prefolder2->rowcount();  
       if ($rowcs > 0) {
       $lastfolderid= $lastfolder[0];
       
       $prefolder22 = $cconn->prepare("SELECT  `$image_folder` FROM `$images` WHERE `$imageid`='$lastfolderid'  ");
       $prefolder22->execute();    $lastfolder22= $prefolder22->fetch();   
       $rowcs22= $prefolder22->rowcount();  
    
        if ($rowcs22 > 0) {
      
       $lastfoldername= $lastfolder22[0];
        } else {
     
       $lastfoldername= $folder; 
        }
       
        } else {
       //$lastfolderid= 0;
      $lastfoldername= $folder;
        }
        
       
        
       $prefolder3 = $cconn->prepare("SELECT  `$imageid` FROM `$images` WHERE `$image_folder`='$lastfoldername'  ");
       $prefolder3->execute();     $rowcs2= $prefolder3->rowcount(); 
       
        if ($rowcs2 <= 1000) {
      $myfoldername= $lastfoldername;
        } else if ($rowcs2 > 1000) {
       $myfoldername = $this->addnewfolder($lastfoldername,$folder,"$prebb"); 
        }
       
        (int)$strlength= strlen($image['name']);
        if ($strlength >= 68) {
        $imageoldname1 = substr($image['name'], 0, 68); $imageoldname= str_replace(' ', '', $imageoldname1); }  //remove space
        else if ($strlength < 68) { $imageoldname1 = $image['name']; $imageoldname= str_replace(' ', '', $imageoldname1);}
        
       $prefolder4 = $cconn->prepare("INSERT INTO `$images` (`$imageid`, `$image_m_folder`, `$image_folder`, `$image_old_name`, `$image_new_name`) VALUES (NULL, '$folder', '$myfoldername', '$imageoldname', '')");
       $prefolder4->execute();
      
       
       $prefolder5 = $cconn->prepare("SELECT  MAX(`$imageid`) FROM `$images` WHERE `$image_m_folder`='$folder' and `$image_folder`='$myfoldername' and `$image_old_name`='$imageoldname'  ");
       $prefolder5->execute();    $lastimage5= $prefolder5->fetch(); $lastimage5id_notyet= $lastimage5[0];
       
       
       $transimgg =  $prebb.'transimg';
        $lastimage5id1= $ash->$transimgg("$lastimage5id_notyet");
        $lastimagename_id = $myfoldername . '__' . $lastimage5id1 . $lastimage5id_notyet.'.jpg'; 
        
        
        $this->images = 'images_'.$prebb.'/'.$myfoldername.'/';  //change to omacc_images/
        $destination =  $this->images.$lastimagename_id;
        
         
        $testmovimgg =  $prebb.'testandmoveimage';
         session::set($testmovimgg, $lastimagename_id);
         $fil = session::get($testmovimgg);
         
        //see what to do  //small,medium,big,logo must be prefixed to the subfolder name
          $pos_small = strpos($folder, 'small');   
          $pos_speak = strpos($folder, 'speak'); 
          $pos_gal = strpos($folder, 'gal');
          $pos_medium = strpos($folder, 'medium');  
          $pos_big = strpos($folder, 'big');
          $pos_large = strpos($folder, 'large');
          $pos_logo = strpos($folder, 'logo');
    
    if  ($pos_small !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$smallwidth,$smallheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    } 
    else if  ($pos_logo !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$logowidth,$logoheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    }
    
    else if  ($pos_speak !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$speakwidth,$speakheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    } 
    
     else if  ($pos_gal !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$galwidth,$galheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    }
    
    else if  ($pos_medium !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$mediumwidth,$mediumheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    } 
    
    else if  ($pos_big !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$bigwidth,$bigheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    } 
    
    
    else if  ($pos_large !== false)
    {
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$largewidth,$largeheight); 
        
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         return $image['name'];
    } 
    
    
    
    else { 
         $savedimage = $this->compress($image['tmp_name'],$destination,$quality,$bigwidth,$bigheight); 
         
         $prec = $cconn->prepare("UPDATE `$images` SET `$image_new_name`='$lastimagename_id'   WHERE `$imageid` ='$lastimage5id_notyet' and `$image_m_folder`='$folder' and `$image_folder`='$myfoldername'");
         $prec->execute();
         
         
         return $image['name'];
    }
              
           }
           
           else {
         return 'noo.xyz';}
    }
    
   
    
    function checkemaila($args) {
     $args = filter_var($args, FILTER_SANITIZE_EMAIL);  
     if(filter_var($args,FILTER_VALIDATE_EMAIL))  {return 'yes';} 
     else { return 'no'; }   
    }


    //add the table 
    //add the ash->transimg
    
   function getrgbvall($args) { 
    
       if ($args == "BLACK") {return '0]0]0'; }  else if ($args == "WHITE") {return '255]255]255'; } //000000 ffffff
  else if ($args == "TAN") {return '210]180]140'; }  else if ($args == "ALICEBLUE") {return '240]248]255'; } //#d2b48c  f0f8ff
  else if ($args == "BURLYWOOD") {return '222]184]135'; }  else if ($args == "CYAN") {return '0]255]255'; } //deb887  00ffff
  else if ($args == "CORNSILK") {return '255]248]220'; }  else if ($args == "DARKBLUE") {return '0]0]139'; } //fff8dc  00008b
  else if ($args == "DARKSEAGREEN") {return '143]188]143'; }  else if ($args == "BUTTONFACE") {return '240]240]240'; } //8fbc8f  f0f0f0
  else if ($args == "GREY") {return '128]128]128'; }  else if ($args == "GRAY") {return '128]128]128'; } // 808080
  else if ($args == "KHAKI") {return '240]230]140'; }  else if ($args == "LIGHTBLUE") {return '173]216]230'; } //f0e68c  add8e6
  else if ($args == "LIGHTSKYBLUE") {return '135]206]250'; }  else if ($args == "SKYBLUE") {return '135]206]235'; } //87cefa  87ceeb
  else if ($args == "LIGHTYELLOW") {return '255]255]224'; }  else if ($args == "LIGHTGREEN") {return '144]238]144'; } //ffffe0  90ee90
  else if ($args == "BLUE") {return '0]0]255'; }  else if ($args == "MAGENTA") {return '255]0]255'; } //0000ff  ff00ff
  else if ($args == "OLIVE") {return '128]128]0'; }  else if ($args == "PLUM") {return '221]160]221'; } //808000  dda0dd or 142,69,133  8e4585
  else if ($args == "PINK") {return '255]192]203'; }  else if ($args == "PALEGREEN") {return '152]251]152'; } //ffc0cb  98fb98
  else if ($args == "STEELBLUE") {return '70]130]180'; }  else if ($args == "SNOW") {return '255]250]250'; } //4682b4 fffafa
  else if ($args == "TURQUOISE") {return '64]224]208'; }  else if ($args == "GREENYELLOW") {return '173]255]47'; } // 40e0d0  adff2f
  else if ($args == "WHEAT") {return '245]222]179'; }  else  {return '255]255]255'; } // f5deb3  
       
   }
    
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
//////////////////////////////////startt add to online////////////////////////////////////////////////////////////////////////////////////////   
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
   
 function deletethisimage_mari ($arggs)  {
 $arggs = rtrim($arggs,']');   $arggs= explode(']', $arggs); 
          $arggs0=$arggs[0]; //table name 
          $arggs1=$arggs[1]; //pics column name 
          $arggs2=$arggs[2]; //url longid
          $arggs3=$arggs[3]; //url   
          $arggs4=$arggs[4]; //pics column id 
          
          $idcol = $arggs0.'id';
    
   $mariconn= new maridatabase();   
  $pre = $mariconn->prepare("SELECT * FROM `$arggs0`  WHERE `$idcol`='$arggs4' AND `mari_loginid` ='$arggs2' AND `mari_url`='$arggs3' ");
  $pre->execute();     $pre->setfetchmode(2); $details= $pre->fetch();   $rows=$pre->rowcount();
      if ($rows >0) {   
     
     $imagevalueneeded = $details[$arggs1]; 
   $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args);  
    if (isset  ($args[1]) ) { 
      $file = 'images_mari/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { $hhh = 1;}
    $pred = $mariconn->prepare("DELETE FROM `mari_images` WHERE `mari_image_new_name`='$imagevalueneeded'");
    $pred->execute();   
     } }
    return 'done';
   }  else { return 'done';}
}   
   

function deletethisimage_event ($arggs)  {
 $arggs = rtrim($arggs,']');   $arggs= explode(']', $arggs); 
          $arggs0=$arggs[0]; //table name 
          $arggs1=$arggs[1]; //pics column name 
          $arggs2=$arggs[2]; //url longid
          $arggs3=$arggs[3]; //url   
          $arggs4=$arggs[4]; //pics column id 
          
          $idcol = $arggs0.'id';
    
   $eventconn= new eventdatabase();   
  $pre = $eventconn->prepare("SELECT * FROM `$arggs0`  WHERE `$idcol`='$arggs4' AND `event_loginid` ='$arggs2' AND `event_url`='$arggs3' ");
  $pre->execute();     $pre->setfetchmode(2); $details= $pre->fetch();   $rows=$pre->rowcount();
      if ($rows >0) {   
     
     $imagevalueneeded = $details[$arggs1]; 
   $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args);  
    if (isset  ($args[1]) ) { 
      $file = 'images_event/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { $hhh = 1;}
    $pred = $eventconn->prepare("DELETE FROM `event_images` WHERE `event_image_new_name`='$imagevalueneeded'");
    $pred->execute();   
     } }
    return 'done';
   }  else { return 'done';}
} 


function deletethisimage_org ($arggs)  { 
 $arggs = rtrim($arggs,']');   $arggs= explode(']', $arggs); 
          $arggs0=$arggs[0]; //table name 
          $arggs1=$arggs[1]; //pics column name 
          $arggs2=$arggs[2]; //url longid
          $arggs3=$arggs[3]; //url   
          $arggs4=$arggs[4]; //pics column id 
          
          $idcol = $arggs0.'id';
    
   $orgconn= new orgdatabase();   
  $pre = $orgconn->prepare("SELECT * FROM `$arggs0`  WHERE `$idcol`='$arggs4' AND `org_loginid` ='$arggs2' AND `org_url`='$arggs3' ");
  $pre->execute();     $pre->setfetchmode(2); $details= $pre->fetch();   $rows=$pre->rowcount(); 
      if ($rows >0) {   
     
     $imagevalueneeded = $details[$arggs1];  
   $args = $imagevalueneeded; $args = rtrim($args,'__');     $args= explode('__', $args);  
    if (isset  ($args[1]) ) { 
      $file = 'images_org/'.$args[0].'/'.$imagevalueneeded;
       if(file_exists($file))
        {
if (!unlink($file)) { $hhh = 1;}
    $pred = $orgconn->prepare("DELETE FROM `org_images` WHERE `org_image_new_name`='$imagevalueneeded'");
    $pred->execute();   
     } }
    return 'done';
   }  else { return 'done';}
} 


   
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
//////////////////////////////////endd add to online////////////////////////////////////////////////////////////////////////////////////////////////   
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
 //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////   
        
    
}