<?php


ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', 300);

require_once __DIR__ . '/vendor/autoload.php';



function get_pays($objWorkSheet){

    $list_pays = [];
    $rows = $objWorkSheet->getRowIterator();
    //$excel_array = [];
    foreach ($rows as $row) {
        $rowIndex = $row->getRowIndex();
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);

        $one_pays = $objWorkSheet->getCellByColumnAndRow(0, $rowIndex)->getValue();
        array_push($list_pays, $one_pays);
    }
    $list_pays = array_values(array_unique($list_pays));
    unset($list_pays[0]);
    $list_pays = array_values($list_pays);

    return $list_pays;
}


function echo_matrix($matrix) {
    echo "<table style='border : black 1px solid'>";
    foreach ($matrix as $key => $value) {
        echo "<tr>";
        foreach ($value as $key2 => $value2) {
            $result[$key2][$key] = $value2;
            echo "<td style='border : black 1px solid'>";
            echo $matrix[$key][$key2];
            echo "</td>";
        }
        echo "</tr>";
    }
    echo "</table>";
}


function transpose($matrix) {
    foreach ($matrix as $key => $value) {
        foreach ($value as $key2 => $value2) {
            $result[$key2][$key] = $value2;
        }
    }
    return $result;
}



function excel_to_array($objWorkSheet){
    /**
     * INDEX COLONNE start = 0
     * INDEX LIGNE start = 1
     */
    $rows = $objWorkSheet->getRowIterator();
    //$excel_array = [];
    foreach ($rows as $row) {

        $rowIndex = $row->getRowIndex();
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);

        $colIndex = 1;
        foreach ($cellIterator as $cell) {
            if ($colIndex != 1 && $colIndex != 2 && $colIndex != 4) {
                $excel_array[$rowIndex][$colIndex] = $cell->getValue();
            }
            $colIndex++;
        }
    }
/*
    echo $rowIndex." lignes\n";
    echo $colIndex." colonnes";
*/
    return $excel_array;
}



/**
 * @var $cell PHPExcel_Cell
 */

$file = "data_5.xlsx";
//$file = "test.xlsx";

// EXCEL READER
$inputFileType = PHPExcel_IOFactory::identify($file);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objReader->setReadDataOnly(true);
$objPHPExcel = $objReader->load($file);
$objWorkSheet = $objPHPExcel->getActiveSheet();

// on récupère la liste des pays trouvé dans le excel
$list_pays = get_pays($objWorkSheet);

// on récupère les données de l'excel sous forme de tableau PHP
$excel_array = excel_to_array($objWorkSheet);

//echo_matrix($excel_array);


// parcourt de 80 lignes en 80 lignes -> donc 1 boucle par pays
// de 1 à 81, de 82 à 161, etc...
$new_array = [];
// on initialise un compteur pour connaitre au cbième pays on en est,
// -> sert à associer à avec le pays correspond dans la liste de pays
$cpt = 0;
for ($i = 1; $i < count($excel_array); $i = $i + 80){
    // pour chaque pays, on parcourt de ligne en ligne
    // on crée un tableau pays[] pour chaque pays
    // que l'on transpose ensuite
    $pays = [];
    // on récupère les années en ligne dans le tableau original
    $year = $excel_array[1];
    // on supprime le libellé
    unset($year[3]);
    // on met un nouveau libellé
    array_unshift($year, 'Année');
    // on tri les index de ce tableau d'années
    $year = array_values($year);
    // on met en index 1 du tableau d'1 pays, les années
    $pays[1] = $year;
    for($row = $i+1; $row <= $i + 80; $row++){
        $pays[$row] = $excel_array[$row];
        $pays[$row] = array_values($pays[$row]);
    }

    // on transpose le tableau d'un pays
    $pays = transpose($pays);


    // ajouter les pays au débtut du tableau
    array_unshift($pays[0], 'Pays');
    for($row = 1; $row < 58; $row++){
        if($cpt < 61){
            array_unshift($pays[$row], $list_pays[$cpt]);
        }
    }

    // garder les titres que pour le 1er pays
    if($i != 1){
        unset($pays[0]);
    }

    // on fusionne le pays courant avec le nouveau tableau
    $new_array = array_merge($new_array, $pays);
    $cpt++;
}


//echo_matrix($new_array);






$output_file = "output_data_5.xlsx";
//$output_file = "output.xlsx";

// EXCEL WRITER
$outputFileType = PHPExcel_IOFactory::identify($file);
$objReader = PHPExcel_IOFactory::createReader($outputFileType);
$objPHPExcel = $objReader->load($output_file);
$objWorkSheet = $objPHPExcel->getActiveSheet();
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $outputFileType);


$nb_lines = count($new_array);


$error = false;

for($x = 0; $x < $nb_lines; $x++){
    $nb_col = count($new_array[$x]);
    for($y = 0; $y < $nb_col; $y++){

        $cell_x = $x + 1;
        $cell_y = $y;

        if(isset($new_array[$x][$y])) {
            $objWorkSheet->setCellValueByColumnAndRow($cell_y, $cell_x, $new_array[$x][$y]);
        }
    }
}

$objWriter->save($output_file);


