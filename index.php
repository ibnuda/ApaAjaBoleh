<?php
/**
 * Created by PhpStorm.
 * User: Ibnu
 * Date: 12/06/2016
 * Time: 00.33
 */

require_once dir('') . '/Classes/PHPExcel.php';

// Object initialization.
$phpExcel = new PHPExcel();

// Set title and stuff like that.
$phpExcel->getProperties()->setTitle("Ini apa ya");

$phpExcel->setActiveSheetIndex(0)
    ->setCellValue("A1", "Nama Orang")->setCellValue("A2", "Hobi")->setCellValue("A3", "Tingkatan")
    ->setCellValue("B1", "Tomi Ganteng")->setCellValue("B2", "Mancing Kerusuhan")->setCellValue("B3", "Pemula")
    ->setCellValue("C1", "Pepe")->setCellValue("C2", "Enternet-an")->setCellValue("C3", "Brewok");

// We have to create another sheet before using it.
$phpExcelNewSheet = new PHPExcel_Worksheet($phpExcel, "Lembar Kedua");
$phpExcel->addSheet($phpExcelNewSheet);

// Change sheet index.
$phpExcel->setActiveSheetIndex(1)
    ->setCellValue("A1", "Nama Orang")->setCellValue("A2", "Kerjaan")->setCellValue("A3", "Tingkatan")
    ->setCellValue("B1", "Ibnu Tampan")->setCellValue("B2", "Nunggu Rumah")->setCellValue("B3", "Grandmaster")
    ->setCellValue("C1", "Wojak")->setCellValue("C2", "Enternet-an")->setCellValue("C3", "Brewok");

// Create the excel file from the previous modifications.
// Excel2007 -> for .xlsx.
// Otherwise, use Excel5
$phpWriter = PHPExcel_IOFactory::createWriter($phpExcel, "Excel5");

$phpExcel->setActiveSheetIndex(0);
// Save it to disk.
$phpWriter->save(str_replace('.php', '.xls', __FILE__));

// Show the content of the file to the browser.

$fileReader = PHPExcel_IOFactory::createReader('Excel5');
echo str_replace('.php', '.xls', __FILE__);
$readFile = $fileReader->load(str_replace('.php', '.xls', __FILE__));

// Show the content of the file to the browser.
foreach ($readFile->getWorksheetIterator() as $worksheet) {
    echo 'Laman ', $worksheet->getTitle() . "<br>";
    echo '<table>';
    foreach ($worksheet->getRowIterator() as $row){
        echo '<tr>';
        $iter = $row->getCellIterator();
        $iter->setIterateOnlyExistingCells(true);
        foreach ($iter as $item) {
            echo '<td>' . $item->getCalculatedValue() . '</td>';
        }
        echo '</tr>';
    }
    echo '</table>';
}

// Maybe we can use it as input value for an upload form in an "edit excel file" scenario.
echo '<form action="form_action.php" method="POST">';
foreach ($readFile->getWorksheetIterator() as $worksheet) {
    echo 'Laman ', $worksheet->getTitle() . "<br>";
    foreach ($worksheet->getRowIterator() as $row){
        $iter = $row->getCellIterator();
        $iter->setIterateOnlyExistingCells(true);
        foreach ($iter as $item) {
            echo '<input type="text" value="' . $item->getCalculatedValue() . '" name="' . $item->getCoordinate() . '">';
        }
        echo '<br>';
    }
}
echo '<input type="submit">';
echo '</form>';
