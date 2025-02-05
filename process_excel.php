<?php
require 'vendor/autoload.php'; // Asegúrate de instalar PhpSpreadsheet con Composer

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

session_start();

// Array de referencia con los datos de los trabajadores
$trabajadores = [
    'AFQ00001' => ['clave' => '6203', 'nombre' => 'SANTIAGO NORBERTO DIANA LAURA'],
    'AFQ00002' => ['clave' => '6306', 'nombre' => 'ALAN GAMALIEL CONTRERAS GUADARRAMA'],
    'AFQ00003' => ['clave' => '2158', 'nombre' => 'BECERRIL VELAZQUEZ PAOLA'],
    'AFQ00004' => ['clave' => '6390', 'nombre' => 'MORENO BONILLA RAUL IVAN'],
    'AFQ00005' => ['clave' => '7', 'nombre' => 'LOPEZ REYES RODRIGO EDUARDO'],
    'AFQ00006' => ['clave' => '8', 'nombre' => 'GONZALEZ VALENTIN CESAR RAUL'],
    'AFQ00007' => ['clave' => '9', 'nombre' => 'MARTINEZ MIRANDA ANGELICA'],
    'AFQ00010' => ['clave' => '13', 'nombre' => 'RECOBA HUERTA JOSUE'],
    'AFQ00012' => ['clave' => '15', 'nombre' => 'CORTES SANTES LAURA'],
    'AFQ00013' => ['clave' => '6452', 'nombre' => 'OSCAR EMANUEL PASCUAL SANCHEZ'],
    'AFQ00014' => ['clave' => '17', 'nombre' => 'MENDOZA ESCOBEDO JUANA'],
    'AFQ00015' => ['clave' => '18', 'nombre' => 'GONZALEZ YESSICA DEL CARMEN'],
    'ATQ0002'  => ['clave' => '16', 'nombre' => 'HERNANDEZ VELASCO ROCIO'],
    'ATQ0005'  => ['clave' => '19', 'nombre' => 'ITZANA LUZ VALENTE HEREDIA'],
];

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFile'])) {
    $file = $_FILES['excelFile']['tmp_name'];
    
    if ($file) {
        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        
        // Leer los datos
        $data = $sheet->toArray();
        
        // Transformar las fechas en la fila 5, columnas 5 a 18
        if (!empty($data[4])) { // Fila 5 (índice 4)
            for ($col = 4; $col <= 17; $col++) { // Columnas 5 a 18 (índices 4 a 17)
                if (!empty($data[4][$col]) && preg_match('/^\d{2}-\d{2}$/', $data[4][$col])) {
                    // Separar mes y día
                    list($month, $day) = explode('-', $data[4][$col]);
                    
                    // Validar si es una fecha válida
                    if (checkdate($month, $day, 2025)) {
                        // Crear un objeto DateTime con el año 2025
                        $date = DateTime::createFromFormat('m-d-Y', "$month-$day-2025");
                        if ($date) {
                            // Formatear la fecha como DD/MM/AAAA
                            $data[4][$col] = $date->format('d/m/Y');
                        }
                    }
                }
            }
        }
        
        // Añadir las columnas extras (Clave del Trabajador y Nombre completo) a partir de la fila 8
        for ($i = 7; $i < count($data); $i++) { // Fila 8 en adelante (índice 7)
            $id = $data[$i][0]; // ID del trabajador (primera columna)
            if (isset($trabajadores[$id])) {
                // Añadir Clave del Trabajador y Nombre completo
                $data[$i][] = $trabajadores[$id]['clave'];
                $data[$i][] = $trabajadores[$id]['nombre'];
            } else {
                // Si no se encuentra el ID, dejar las columnas vacías
                $data[$i][] = '';
                $data[$i][] = '';
            }
        }
        
        $_SESSION['excel_data'] = $data;
        
        echo "<h3 class='mt-4 mb-3'>Vista previa del archivo Excel</h3>";
        echo "<div class='table-responsive'>"; // Contenedor para hacer la tabla responsive
        echo "<table class='table table-bordered table-striped table-hover table-custom'>"; // Clases de Bootstrap para diseño
        echo "<thead class='thead-dark'>"; // Cabecera oscura
        echo "<tr>";
        // Encabezados de la tabla (incluyendo las nuevas columnas)
        $headers = array_merge($data[0], ['Clave del Trabajador', 'Nombre completo']);
        foreach ($headers as $header) {
            echo "<th>" . htmlspecialchars($header ?? '') . "</th>";
        }
        echo "</tr>";
        echo "</thead>";
        echo "<tbody>";
        // Mostrar los datos (a partir de la fila 1 para omitir los encabezados)
        for ($i = 1; $i < count($data); $i++) {
            echo "<tr>";
            foreach ($data[$i] as $cell) {
                echo "<td>" . htmlspecialchars($cell ?? '') . "</td>";
            }
            echo "</tr>";
        }
        echo "</tbody>";
        echo "</table>";
        echo "</div>"; // Cierre del contenedor responsive
        echo "<button class='btn btn-primary mt-3' id='exportButton'>Exportar a XLSX</button>";
    } else {
        echo "<div class='alert alert-danger'>Error al subir el archivo.</div>";
    }
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['exportCsv'])) {
    if (!isset($_SESSION['excel_data'])) {
        echo "No hay datos para exportar.";
        exit;
    }
    
    $data = $_SESSION['excel_data'];
    
    // Crear un nuevo archivo XLSX
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    
    // Agregar los datos al archivo XLSX
    foreach ($data as $rowIndex => $row) {
        foreach ($row as $colIndex => $cell) {
            $sheet->setCellValueByColumnAndRow($colIndex + 1, $rowIndex + 1, $cell);
        }
    }
    
    // Guardar el archivo XLSX
    $writer = new Xlsx($spreadsheet);
    $xlsxFile = 'exported_excel.xlsx';
    $writer->save($xlsxFile);
    
    // Descargar el archivo XLSX
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . basename($xlsxFile) . '"');
    header('Cache-Control: max-age=0');
    readfile($xlsxFile);
    exit;
}
?>