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
$sheet->setCellValue('B1', 'User Email');
$sheet->setCellValue('C1', 'Mobil');
$sheet->setCellValue('D1', 'Dari Tanggal');
$sheet->setCellValue('E1', 'Sampai Tanggal');
$sheet->setCellValue('F1', 'Pesan');
$sheet->setCellValue('G1', 'Status');
$sheet->setCellValue('H1', 'Tanggal Posting');

$query = mysqli_query($conn, "SELECT * FROM tblbooking");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['userEmail']);
    $sheet->setCellValue('C' . $i, $row['VehicleId']);
    $sheet->setCellValue('D' . $i, $row['FromDate']);
    $sheet->setCellValue('E' . $i, $row['ToDate']);
    $sheet->setCellValue('F' . $i, $row['message']);
    $sheet->setCellValue('G' . $i, $row['Status']);
    $sheet->setCellValue('H' . $i, $row['PostingDate']);
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
$sheet->getStyle('A1:H' . $i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Booking.xlsx');
?>

Penginputan Data Berhasil