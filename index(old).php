<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file_excel'])) {
    try {
        $fileTmpPath = $_FILES['file_excel']['tmp_name'];
        $reader = IOFactory::createReaderForFile($fileTmpPath);
        $spreadsheet = $reader->load($fileTmpPath);

        // Ambil sheet aktif (Data Mentah)
        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray();

        // 1. IDENTIFIKASI HEADER (Baris ke-5 adalah Index 4)
        $headerRow = $data[4];
        $pivotData = [];

        // Mapping index kolom berdasarkan nama (case-insensitive & trim)
        $colIdx = [];
        foreach ($headerRow as $index => $name) {
            $cleanName = strtolower(trim($name ?? ''));
            if ($cleanName == 'nama pengirim') $colIdx['pengirim'] = $index;
            if ($cleanName == 'nama penerima') $colIdx['penerima'] = $index;
            if ($cleanName == 'ac') $colIdx['ac'] = $index;
            if ($cleanName == 'trx terkirim') $colIdx['trx'] = $index;
        }

        // 2. PROSES PIVOT (Mulai dari baris ke-6 / Index 5)
        for ($i = 5; $i < count($data); $i++) {
            $row = $data[$i];

            // Ambil data berdasarkan mapping index
            $pengirim = isset($colIdx['pengirim']) ? ($row[$colIdx['pengirim']] ?? 'N/A') : 'N/A';
            $penerima = isset($colIdx['penerima']) ? ($row[$colIdx['penerima']] ?? 'N/A') : 'N/A';
            $ac       = isset($colIdx['ac'])       ? ($row[$colIdx['ac']]       ?? '-')   : '-';
            $valTrx   = isset($colIdx['trx'])      ? ($row[$colIdx['trx']]      ?? 0)     : 0;
            $trx      = (float)str_replace(',', '', $valTrx);

            // Buat Unique Key untuk Grouping (Penerima + Pengirim + AC)
            $key = $penerima . "|" . $pengirim . "|" . $ac;

            if (!isset($pivotData[$key])) {
                $pivotData[$key] = [
                    'penerima' => $penerima,
                    'pengirim' => $pengirim,
                    'ac' => $ac,
                    'sum_trx' => 0,
                    'count_ac' => 0
                ];
            }

            $pivotData[$key]['sum_trx'] += $trx;
            $pivotData[$key]['count_ac'] += 1;
        }

        // 3. BUAT SPREADSHEET BARU
        $newSpreadsheet = new Spreadsheet();

        // --- SHEET 1: HASIL PIVOT ---
        $sheetPivot = $newSpreadsheet->getActiveSheet();
        $sheetPivot->setTitle('Hasil Pivot');
        $headers = ['NAMA PENERIMA', 'NAMA PENGIRIM', 'AC', 'Count of AC', 'Sum of Trx Terkirim'];
        $sheetPivot->fromArray($headers, NULL, 'A1');

        $currentRow = 2;
        foreach ($pivotData as $item) {
            $sheetPivot->setCellValue('A' . $currentRow, $item['penerima']);
            $sheetPivot->setCellValue('B' . $currentRow, $item['pengirim']);
            $sheetPivot->setCellValue('C' . $currentRow, $item['ac']);
            $sheetPivot->setCellValue('D' . $currentRow, $item['count_ac']);
            $sheetPivot->setCellValue('E' . $currentRow, $item['sum_trx']);
            $currentRow++;
        }

        // --- SHEET 2: DATA MENTAH (Copy dari file asli) ---
        $sheetMentah = $newSpreadsheet->createSheet();
        $sheetMentah->setTitle('Data Mentah');
        $sheetMentah->fromArray($data, NULL, 'A1');

        // 4. DOWNLOAD FILE
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="Hasil_Pivot_Bakul.xlsx"');
        $writer = new Xlsx($newSpreadsheet);
        $writer->save('php://output');
        exit;
    } catch (Exception $e) {
        echo "Gagal memproses file: " . $e->getMessage();
    }
}
?>

<!DOCTYPE html>
<html>

<head>
    <title>PHP Bakul Pivot</title>
    <style>
        body {
            font-family: sans-serif;
            background: #f4f7f6;
            display: flex;
            justify-content: center;
            padding-top: 100px;
        }

        .box {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
            text-align: center;
        }

        button {
            background: #28a745;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
        }
    </style>
</head>

<body>
    <div class="box">
        <h2>Bakul Excel Processor (PHP)</h2>
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="file_excel" required>
            <br><br>
            <button type="submit">Upload & Download Pivot</button>
        </form>
    </div>
</body>

</html>
