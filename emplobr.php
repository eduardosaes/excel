<?php

header("Content-Type:text/html;charset=utf-8");
	$h="localhost";
	$u="root";
	$p="infor1234";
	$con=mysql_connect($h,$u,$p) or die (mysql_error());
	mysql_select_db("pruvas",$con) or die (mysql_error());
	mysql_query("SET NAMES 'utf8'");

 
 require_once("PHPExcel/Classes/PHPExcel/IOFactory.php");
 $nombreArchivo = 'excel/02-20160630-personal_contratado_junio_2016.xls';
// Cargo la hoja de cálculo
 $objPHPExcel = PHPExcel_IOFactory::load($nombreArchivo);

 //Asigno la hoja de calculo activa
$objPHPExcel->setActiveSheetIndex(0);
//Obtengo el numero de filas del archivo
$numRows = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();



 for ($i = 2; $i <= $numRows; $i++) {
    
    $informacion[] = array(
        'codigo' => $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getCalculatedValue(),
        'cedula' => $objPHPExcel->getActiveSheet()->getCell('B'.$i)->getCalculatedValue(),
        'nombre' => $objPHPExcel->getActiveSheet()->getCell('H'.$i)->getCalculatedValue(),
        'cargo' => $objPHPExcel->getActiveSheet()->getCell('J'.$i)->getCalculatedValue(),
        'f_ingreso' => $objPHPExcel->getActiveSheet()->getCell('K'.$i)->getCalculatedValue(),
        'actividad' => $objPHPExcel->getActiveSheet()->getCell('C'.$i)->getCalculatedValue(),
        'sueldo' => $objPHPExcel->getActiveSheet()->getCell('AM'.$i)->getCalculatedValue(),
        

        's1' => $objPHPExcel->getActiveSheet()->getCell('R'.$i)->getCalculatedValue(),
        's2' => $objPHPExcel->getActiveSheet()->getCell('S'.$i)->getCalculatedValue(),
        's3' => $objPHPExcel->getActiveSheet()->getCell('U'.$i)->getCalculatedValue(),
        's4' => $objPHPExcel->getActiveSheet()->getCell('V'.$i)->getCalculatedValue(),
        's5' => $objPHPExcel->getActiveSheet()->getCell('W'.$i)->getCalculatedValue(),
        's6' => $objPHPExcel->getActiveSheet()->getCell('AF'.$i)->getCalculatedValue(),
        's7' => $objPHPExcel->getActiveSheet()->getCell('AG'.$i)->getCalculatedValue(),
        's8' => $objPHPExcel->getActiveSheet()->getCell('AH'.$i)->getCalculatedValue(),
        's9' => $objPHPExcel->getActiveSheet()->getCell('AI'.$i)->getCalculatedValue(),


    );

 }

 foreach ($informacion as $registro) {
			    $codigo=$registro['codigo'];	
			    $cedula=$registro['cedula'];
			    $nombre=$registro['nombre'];
			    $cargo=$registro['cargo'];
			    $f_ingreso=$registro['f_ingreso'];
			    $actividad=$registro['actividad'];

			    $s1=$registro['s1'];
			    $s2=$registro['s2'];
			    $s3=$registro['s3'];
			    $s4=$registro['s4'];
			    $s5=$registro['s5'];
			    $s6=$registro['s6'];
			    $s7=$registro['s7'];
			    $s8=$registro['s8'];
			    $s9=$registro['s9'];	
			    $s10=$registro['sueldo'];

			 $sueldo = $s1+$s2+$s3+$s4+$s5+$s6+$s7+$s8+$s9+$s10;	


		    $sql = "INSERT INTO `obreros`(`cod`, `cedula`, `nomp`, `cargo`, `fechaing`, `actividad`, `sueldo`) VALUES ('$codigo','$cedula','$nombre','$cargo','$f_ingreso','$actividad','$sueldo')";
		    mysql_query($sql);
 }



 ?>