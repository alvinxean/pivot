<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file_excel'])) {
    try {
        $fileTmpPath = $_FILES['file_excel']['tmp_name'];
        $originalFileName = $_FILES['file_excel']['name'];

        $fileInfo = pathinfo($originalFileName);
        $exportFileName = $fileInfo['filename'] . ' (Pivot).xlsx';

        $reader = IOFactory::createReaderForFile($fileTmpPath);
        $spreadsheet = $reader->load($fileTmpPath);

        $sheet = $spreadsheet->getActiveSheet();
        $data = $sheet->toArray();

        // 1. IDENTIFIKASI HEADER (Baris ke-5 adalah Index 4)
        $headerRow = $data[4];
        $pivotData = [];

        $colIdx = [];
        foreach ($headerRow as $index => $name) {
            $cleanName = strtolower(trim($name ?? ''));
            if ($cleanName == 'no') $colIdx['no'] = $index;
            if ($cleanName == 'nama pengirim') $colIdx['pengirim'] = $index;
            if ($cleanName == 'nama penerima') $colIdx['penerima'] = $index;
            if ($cleanName == 'ac') $colIdx['ac'] = $index;
            if ($cleanName == 'trx terkirim') $colIdx['trx'] = $index;
        }

        $grandTotalCount = 0;
        $grandTotalSum = 0;

        // 2. PROSES PIVOT
        for ($i = 5; $i < count($data); $i++) {
            $row = $data[$i];

            $noVal = isset($colIdx['no']) ? trim($row[$colIdx['no']] ?? '') : '';
            if ($noVal === '' || !is_numeric($noVal)) {
                continue;
            }

            $pengirimRaw = isset($colIdx['pengirim']) ? ($row[$colIdx['pengirim']] ?? 'N/A') : 'N/A';
            $penerimaRaw = isset($colIdx['penerima']) ? ($row[$colIdx['penerima']] ?? 'N/A') : 'N/A';
            $ac          = isset($colIdx['ac'])       ? trim($row[$colIdx['ac']] ?? '-')   : '-';
            $valTrx      = isset($colIdx['trx'])      ? ($row[$colIdx['trx']]      ?? 0)     : 0;

            $penerima = trim($penerimaRaw);
            $pengirim = trim($pengirimRaw);
            $trx      = (float)str_replace(',', '', $valTrx);

            $key = strtolower($penerima) . "|" . strtolower($pengirim) . "|" . $ac;

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

            $grandTotalSum += $trx;
            $grandTotalCount += 1;
        }

        // 3. BUAT SPREADSHEET BARU
        $newSpreadsheet = new Spreadsheet();
        $sheetPivot = $newSpreadsheet->getActiveSheet();
        $sheetPivot->setTitle('Hasil Pivot');
        $headers = ['NAMA PENERIMA', 'NAMA PENGIRIM', 'AC', 'Count of AC', 'Sum of Trx Terkirim'];
        $sheetPivot->fromArray($headers, NULL, 'A1');

        $currentRow = 2;
        foreach ($pivotData as $item) {
            $sheetPivot->setCellValue('A' . $currentRow, $item['penerima']);
            $sheetPivot->setCellValue('B' . $currentRow, $item['pengirim']);
            $sheetPivot->setCellValueExplicit('C' . $currentRow, $item['ac'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheetPivot->setCellValue('D' . $currentRow, $item['count_ac']);
            $sheetPivot->setCellValue('E' . $currentRow, $item['sum_trx']);

            $sheetPivot->getStyle('E' . $currentRow)->getNumberFormat()->setFormatCode('#,##0');
            $currentRow++;
        }

        // --- TAMBAHKAN BARIS GRAND TOTAL ---
        $sheetPivot->setCellValue('A' . $currentRow, 'Total Terakhir / Grand Total');
        $sheetPivot->setCellValue('D' . $currentRow, $grandTotalCount);
        $sheetPivot->setCellValue('E' . $currentRow, $grandTotalSum);

        $sheetPivot->getStyle('E' . $currentRow)->getNumberFormat()->setFormatCode('#,##0');

        // --- 🛠️ STYLING: BOLD & ALIGNMENT EXCEL 🛠️ ---

        // A. Format Tebal (Bold) untuk semua sel di baris Grand Total & Header
        $sheetPivot->getStyle('A' . $currentRow . ':E' . $currentRow)->getFont()->setBold(true);
        $sheetPivot->getStyle('A1:E1')->getFont()->setBold(true);

        // B. Loop Global: Atur Lebar Otomatis dan Set Rata Kiri dari Baris 2 sampai baris Grand Total
        $columns = ['A', 'B', 'C', 'D', 'E'];
        foreach ($columns as $col) {
            $sheetPivot->getStyle($col . '2:' . $col . $currentRow)
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_LEFT);

            $sheetPivot->getColumnDimension($col)->setAutoSize(true);
        }

        // C. Override Khusus (Ditulis paling terakhir agar tidak tertimpa):
        // 1. Baris Header (Baris 1) dipaksa Rata Tengah (Center)
        $sheetPivot->getStyle('A1:E1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // 2. Isi kolom Count of AC (Kolom D) di baris Grand Total dipaksa Rata Tengah (Center)
        $sheetPivot->getStyle('D' . $currentRow)
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);


        // --- SHEET 2: DATA MENTAH (Copy dari file asli) ---
        $sheetMentah = $newSpreadsheet->createSheet();
        $sheetMentah->setTitle('Data Mentah');
        $sheetMentah->fromArray($data, NULL, 'A1');

        // 4. DOWNLOAD FILE
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $exportFileName . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($newSpreadsheet);
        $writer->save('php://output');
        exit;
    } catch (Exception $e) {
        echo "Gagal memproses file: " . $e->getMessage();
    }
}
?>

<!DOCTYPE html>
<html lang="id">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bakul Excel Processor</title>

    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

    <link rel="stylesheet" href="style.css">
</head>

<body>

    <div class="container">
        <div class="logo-area">
            <i class="fa-solid fa-file-excel"></i>
            <h2>Bakul Pivot Processor</h2>
            <p class="subtitle">Unggah file laporan mentah Excel Anda untuk otomatis dikonversi menjadi data format ringkasan pivot terstruktur.</p>
        </div>

        <form method="POST" enctype="multipart/form-data" id="uploadForm">
            <div class="drop-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
                <div class="drop-icon"><i class="fa-solid fa-cloud-arrow-up"></i></div>
                <div class="drop-text">Pilih file atau seret ke sini</div>
                <div class="drop-hint">Hanya mendukung format .xlsx / .xls</div>
                <input type="file" name="file_excel" id="fileInput" accept=".xlsx, .xls" required>
            </div>

            <div class="file-preview" id="filePreview">
                <div class="file-info">
                    <i class="fa-regular fa-file-excel"></i>
                    <span class="file-name" id="fileNameDisplay">Nama_file_kamu.xlsx</span>
                </div>
                <span style="font-size: 11px; color: var(--primary); font-weight:600;"><i class="fa-solid fa-circle-check"></i> Ready</span>
            </div>

            <button type="submit" id="btnSubmit">
                <div class="spinner" id="btnSpinner"></div>
                <span id="btnText">Proses & Unduh Pivot</span>
            </button>
        </form>

        <footer>Powered by PhpSpreadsheet Engine v2</footer>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const filePreview = document.getElementById('filePreview');
        const fileNameDisplay = document.getElementById('fileNameDisplay');
        const uploadForm = document.getElementById('uploadForm');
        const btnSubmit = document.getElementById('btnSubmit');
        const btnSpinner = document.getElementById('btnSpinner');
        const btnText = document.getElementById('btnText');

        function handleFileSelection(files) {
            if (files.length > 0) {
                const file = files[0];
                fileNameDisplay.textContent = file.name;
                filePreview.style.display = 'flex';
                dropZone.style.borderColor = 'var(--primary)';
            }
        }

        fileInput.addEventListener('change', (e) => handleFileSelection(e.target.files));

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drag-over');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('drag-over');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-over');
            if (e.dataTransfer.files.length > 0) {
                fileInput.files = e.dataTransfer.files;
                handleFileSelection(e.dataTransfer.files);
            }
        });

        uploadForm.addEventListener('submit', () => {
            btnSubmit.disabled = true;
            btnSpinner.style.display = 'block';
            btnText.textContent = 'Sedang Memproses...';

            setTimeout(() => {
                btnSubmit.disabled = false;
                btnSpinner.style.display = 'none';
                btnText.textContent = 'Proses & Unduh Pivot';
            }, 3000);
        });
    </script>
</body>

</html>