<?php
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "carrental";

$conn = mysqli_connect($servername, $username, $password, $dbname);

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'id');
$sheet->setCellValue('B1', 'Nama Merek');
$sheet->setCellValue('C1', 'Tanggal Dibuat');
$sheet->setCellValue('D1', 'Tanggal Update');

$query = mysqli_query($conn, "SELECT * FROM tblbrands");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['BrandName']);
    $sheet->setCellValue('C' . $i, $row['CreationDate']);
    $sheet->setCellValue('D' . $i, $row['UpdationDate']);
    $i++;
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];

$i = $i - 1;
$sheet->getStyle('A1:D' . $i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Merek.xlsx');
?>

Penginputan Data Berhasil