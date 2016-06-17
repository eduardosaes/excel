<?php

header("Content-Type:text/html;charset=utf-8");
	$h="localhost";
	$u="root";
	$p="infor1234";
	$con=mysql_connect($h,$u,$p) or die (mysql_error());
	mysql_select_db("pruvas",$con) or die (mysql_error());
	mysql_query("SET NAMES 'utf8'");

 
 require_once("PHPExcel/Classes/PHPExcel/IOFactory.php");
 $nombreArchivo = 'excel/01-20160630-empleados_fijos_junio_2016.xls';
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
        
    );

 }

 foreach ($informacion as $registro) {
			    $codigo=$registro['codigo'];	
			    $cedula=$registro['cedula'];
			    $nombre=$registro['nombre'];
			    $cargo=$registro['cargo'];
			    $f_ingreso=$registro['f_ingreso'];
			    $actividad=$registro['actividad'];	
			    $sueldo=$registro['sueldo'];
		    $sql = "INSERT INTO `obreros`(`cod`, `cedula`, `nomp`, `cargo`, `fechaing`, `actividad`, `sueldo`) VALUES ('$codigo','$cedula','$nombre','$cargo','$f_ingreso','$actividad','$sueldo')";
		    mysql_query($sql);
 }



 ?>