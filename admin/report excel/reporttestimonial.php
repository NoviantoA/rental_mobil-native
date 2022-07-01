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
$sheet->setCellValue('B1', 'Email');
$sheet->setCellValue('C1', 'Testimonial');
$sheet->setCellValue('D1', 'Tanggal Posting');
$sheet->setCellValue('E1', 'Status');

$query = mysqli_query($conn, "SELECT * FROM tbltestimonial");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['UserEmail']);
    $sheet->setCellValue('C' . $i, $row['Testimonial']);
    $sheet->setCellValue('D' . $i, $row['PostingDate']);
    $sheet->setCellValue('E' . $i, $row['status']);
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
$sheet->getStyle('A1:E' . $i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Testimonial.xlsx');
?>

Penginputan Data Berhasil