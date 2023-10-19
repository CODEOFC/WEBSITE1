<?php
// Include the PHPExcel library
require_once 'PHPExcel-1.8/Classes/PHPExcel.php';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Retrieve form data from POST request
    $name = $_POST['name'];
    $email = $_POST['email'];
    $phone = $_POST['phone'];
    $address = $_POST['address'];

    // Create a PHPExcel object
    $objPHPExcel = new PHPExcel();

    // Create a worksheet
    $worksheet = $objPHPExcel->getActiveSheet();

    // Set the column headers
    $worksheet->setCellValue('A1', 'Name');
    $worksheet->setCellValue('B1', 'Email');
    $worksheet->setCellValue('C1', 'Phone');
    $worksheet->setCellValue('D1', 'Address');

    // Populate the Excel sheet with form data
    $row = 2;
    $worksheet->setCellValue('A' . $row, $name);
    $worksheet->setCellValue('B' . $row, $email);
    $worksheet->setCellValue('C' . $row, $phone);
    $worksheet->setCellValue('D' . $row, $address);

    // Save the Excel file
    $writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $writer->save('financial_applications.xlsx');

    echo "Form data has been successfully written to Excel.";
} else {
    echo "This script is intended to be executed in a web context (e.g., as a form submission).";
}
?>
