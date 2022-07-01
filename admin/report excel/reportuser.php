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
$sheet->setCellValue('B1', 'Nama Depan');
$sheet->setCellValue('C1', 'Email');
$sheet->setCellValue('D1', 'Password');
$sheet->setCellValue('E1', 'Nomor Telephon');
$sheet->setCellValue('F1', 'Ulang Tahun');
$sheet->setCellValue('G1', 'Alamat');
$sheet->setCellValue('H1', 'Kota');
$sheet->setCellValue('I1', 'Negara');
$sheet->setCellValue('J1', 'Tanggal Registrasi');
$sheet->setCellValue('K1', 'Tanggal Update');

$query = mysqli_query($conn, "SELECT * FROM tblusers");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['FullName']);
    $sheet->setCellValue('C' . $i, $row['EmailId']);
    $sheet->setCellValue('D' . $i, $row['Password']);
    $sheet->setCellValue('E' . $i, $row['ContactNo']);
    $sheet->setCellValue('F' . $i, $row['dob']);
    $sheet->setCellValue('G' . $i, $row['Address']);
    $sheet->setCellValue('H' . $i, $row['City']);
    $sheet->setCellValue('I' . $i, $row['Country']);
    $sheet->setCellValue('J' . $i, $row['RegDate']);
    $sheet->setCellValue('K' . $i, $row['UpdationDate']);
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
$sheet->getStyle('A1:K' . $i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data User.xlsx');
?>

Penginputan Data Berhasil