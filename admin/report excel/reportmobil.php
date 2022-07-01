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
$sheet->setCellValue('B1', 'Nama Mobil');
$sheet->setCellValue('C1', 'Merek');
$sheet->setCellValue('D1', 'Deskripsi');
$sheet->setCellValue('E1', 'Harga Perhari');
$sheet->setCellValue('F1', 'Bahan Bakar');
$sheet->setCellValue('G1', 'Model Tahun');
$sheet->setCellValue('H1', 'Tempat Duduk');
$sheet->setCellValue('I1', 'Gambar 1');
$sheet->setCellValue('J1', 'Gambar 2');
$sheet->setCellValue('K1', 'Gambar 3');
$sheet->setCellValue('L1', 'Gambar 4');
$sheet->setCellValue('M1', 'Gambar 5');
$sheet->setCellValue('N1', 'AC');
$sheet->setCellValue('O1', 'Kunci Pintu Elektronik');
$sheet->setCellValue('P1', 'Kunci Anti Maling');
$sheet->setCellValue('Q1', 'Bantuan Rem');
$sheet->setCellValue('R1', 'Power Steering');
$sheet->setCellValue('S1', 'Air Bag Pengemudi');
$sheet->setCellValue('T1', 'Air Bag Penumpang');
$sheet->setCellValue('U1', 'Power Windows');
$sheet->setCellValue('V1', 'CD Player');
$sheet->setCellValue('W1', 'Pengaman Central');
$sheet->setCellValue('X1', 'Sensor Kerusakan');
$sheet->setCellValue('Y1', 'Jok Kulit');
$sheet->setCellValue('Z1', 'Tanggal Registrasi');
$sheet->setCellValue('A2', 'Tanggal Update');

$query = mysqli_query($conn, "SELECT * FROM tblvehicles");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['VehiclesTitle']);
    $sheet->setCellValue('C' . $i, $row['VehiclesBrand']);
    $sheet->setCellValue('D' . $i, $row['VehiclesOverview']);
    $sheet->setCellValue('E' . $i, $row['PricePerDay']);
    $sheet->setCellValue('F' . $i, $row['FuelType']);
    $sheet->setCellValue('G' . $i, $row['ModelYear']);
    $sheet->setCellValue('H' . $i, $row['SeatingCapacity']);
    $sheet->setCellValue('I' . $i, $row['Vimage1']);
    $sheet->setCellValue('J' . $i, $row['Vimage2']);
    $sheet->setCellValue('K' . $i, $row['Vimage3']);
    $sheet->setCellValue('L' . $i, $row['Vimage4']);
    $sheet->setCellValue('M' . $i, $row['Vimage5']);
    $sheet->setCellValue('N' . $i, $row['AirConditioner']);
    $sheet->setCellValue('O' . $i, $row['PowerDoorLocks']);
    $sheet->setCellValue('P' . $i, $row['AntiLockBrakingSystem']);
    $sheet->setCellValue('Q' . $i, $row['BrakeAssist']);
    $sheet->setCellValue('R' . $i, $row['PowerSteering']);
    $sheet->setCellValue('S' . $i, $row['DriverAirbag']);
    $sheet->setCellValue('T' . $i, $row['PassengerAirbag']);
    $sheet->setCellValue('U' . $i, $row['PowerWindows']);
    $sheet->setCellValue('V' . $i, $row['CDPlayer']);
    $sheet->setCellValue('W' . $i, $row['CentralLocking']);
    $sheet->setCellValue('X' . $i, $row['CrashSensor']);
    $sheet->setCellValue('Y' . $i, $row['LeatherSeats']);
    $sheet->setCellValue('Z' . $i, $row['RegDate']);
    $sheet->setCellValue('AA' . $i, $row['UpdationDate']);
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
$sheet->getStyle('A1:AA' . $i)->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Data Mobil.xlsx');
?>

Penginputan Data Berhasil