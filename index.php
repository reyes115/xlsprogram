<?php
require 'vendor/autoload.php'; // Asegúrate de instalar PhpSpreadsheet con Composer

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

session_start();

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFile'])) {
    $file = $_FILES['excelFile']['tmp_name'];
    
    if ($file) {
        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        
        // Leer los datos
        $data = $sheet->toArray();
        $_SESSION['excel_data'] = $data;
        
        echo "<h3>Vista previa del archivo</h3>";
        echo "<table class='table table-bordered' id='excelTable'>";
        foreach ($data as $row) {
            echo "<tr>";
            foreach ($row as $cell) {
                echo "<td>" . htmlspecialchars($cell ?? '') . "</td>"; // Manejar valores nulos
            }
            echo "</tr>";
        }
        echo "</table>";
        echo "<button class='btn btn-primary' id='transformButton'>Transformar Fechas</button>";
    } else {
        echo "<div class='alert alert-danger'>Error al subir el archivo.</div>";
    }
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['transformarFechas'])) {
    if (!isset($_SESSION['excel_data'])) {
        echo "No hay datos para transformar.";
        exit;
    }
    
    $data = $_SESSION['excel_data'];
    
    // Aplicar la lógica de transformación de fechas en la fila 4, columnas 5 a 18
    if (!empty($data[3])) {
        for ($key = 4; $key <= 17; $key++) {
            if (!empty($data[3][$key]) && preg_match('/\d{2}-\d{2}/', $data[3][$key])) {
                $date_parts = explode('-', $data[3][$key]);
                $month = $date_parts[0];
                $day = $date_parts[1];
                
                if (checkdate($month, $day, 2025)) {
                    $date = DateTime::createFromFormat('m-d-Y', "$month-$day-2025");
                    if ($date) {
                        setlocale(LC_TIME, 'es_ES.UTF-8');
                        $data[3][$key] = strftime('%d de %B de %Y', $date->getTimestamp());
                    }
                }
            }
        }
    }
    
    $_SESSION['excel_data'] = $data;
    echo json_encode($data);
    exit;
}
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Subir y procesar Excel</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <script>
        function subirArchivo() {
            let formData = new FormData(document.getElementById('uploadForm'));
            fetch('', {
                method: 'POST',
                body: formData
            })
            .then(response => response.text())
            .then(html => {
                document.getElementById('preview').innerHTML = html;
                document.getElementById('uploadForm').reset();
                let transformButton = document.getElementById('transformButton');
                if (transformButton) {
                    transformButton.addEventListener('click', transformarFechas);
                }
            })
            .catch(error => console.error('Error:', error));
        }
        
        function transformarFechas() {
            fetch('', {
                method: 'POST',
                headers: {'Content-Type': 'application/x-www-form-urlencoded'},
                body: 'transformarFechas=1'
            })
            .then(response => response.json())
            .then(data => {
                let table = document.getElementById('excelTable');
                let rows = table.getElementsByTagName('tr');
                let headerRow = rows[3];
                let cells = headerRow.getElementsByTagName('td');
                for (let i = 4; i <= 17; i++) {
                    cells[i].innerText = data[3][i];
                }
            })
            .catch(error => console.error('Error:', error));
        }
    </script>
</head>
<body class="container mt-5">
    <h2 class="text-center">Subir un archivo Excel</h2>
    <form id="uploadForm" enctype="multipart/form-data" onsubmit="event.preventDefault(); subirArchivo();" class="mb-3">
        <input type="file" name="excelFile" accept=".xls,.xlsx" required class="form-control mb-2">
        <button type="submit" class="btn btn-success">Subir</button>
    </form>
    <div id="preview"></div>
</body>
</html>
