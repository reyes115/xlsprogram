<?php
require 'vendor/autoload.php'; // Asegúrate de instalar PhpSpreadsheet con Composer

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Conexión a la base de datos
$servername = "localhost";
$username = "u470151145_ab_forti";
$password = "A8#BfO2r0T4i!";
$database = "u470151145_ceers_2_0";

$conn = new mysqli($servername, $username, $password, $database);
if ($conn->connect_error) {
    die("Error de conexión: " . $conn->connect_error);
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['excelFile'])) {
    $file = $_FILES['excelFile']['tmp_name'];
    
    if ($file) {
        $spreadsheet = IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
        
        // Leer los datos
        $data = $sheet->toArray();
        
        // Nueva estructura de datos
        $newData = [];
        $newData[] = ["Clave trabajador", "Tipo falta", "% falta", "Certificado IMSS", "% pagado por IMSS", "Fecha inicio", "Fecha fin", "Clave tipo de incapacidad", "Observaciones"];
        
        foreach ($data as $row) {
            if (!isset($row[0]) || empty($row[0]) || $row[0] == "ID") {
                continue; // Saltar encabezados y filas vacías
            }
            
            // Buscar la clave del trabajador en la base de datos
            $id = $row[0];
            $query = "SELECT clave_trabajador FROM empleados WHERE id = '$id'";
            $result = $conn->query($query);
            $clave_trabajador = ($result->num_rows > 0) ? $result->fetch_assoc()['clave_trabajador'] : "Desconocido";
            
            // Agregar fila con datos estructurados
            $newData[] = [
                $clave_trabajador,
                $row[1] ?? '',
                $row[2] ?? '',
                $row[3] ?? '',
                $row[4] ?? '',
                $row[5] ?? '',
                $row[6] ?? '',
                $row[7] ?? '',
                $row[8] ?? ''
            ];
        }
        
        // Crear nuevo archivo Excel
        $newSpreadsheet = new Spreadsheet();
        $newSheet = $newSpreadsheet->getActiveSheet();
        $newSheet->fromArray($newData, null, 'A1');
        
        // Guardar el archivo modificado
        $outputFile = 'modified_excel.xlsx';
        $writer = new Xlsx($newSpreadsheet);
        $writer->save($outputFile);
        
        echo "<a href='$outputFile' download>Descargar archivo modificado</a>";
    } else {
        echo "Error al subir el archivo.";
    }
}
$conn->close();
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Subir y procesar Excel</title>
</head>
<body>
    <h2>Subir un archivo Excel</h2>
    <form action="" method="POST" enctype="multipart/form-data">
        <input type="file" name="excelFile" accept=".xls,.xlsx" required>
        <button type="submit">Subir y procesar</button>
    </form>
</body>
</html>
