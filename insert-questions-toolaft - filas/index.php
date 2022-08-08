<?php

// ? Este script lee e itera las filas de una columna obteniendo su valor y enviandolo a la base de datos

require 'vendor/autoload.php';

// Incializar clase Dotenv para variables de entorno
$dotenv = Dotenv\Dotenv::createImmutable(__DIR__ . '/src/config');
$dotenv->load();

# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;
// Clase de conexión a la base de datos
use Tesla\InsertQuestionsToolaft\lib\Database;
use TdTrung\Chalk\Chalk;

$chalk = new Chalk();
$database = new Database();

$rutaArchivo = "preguntas-seg-elec.xlsx";
$documento = IOFactory::load($rutaArchivo);

# Obtener conteo de hojas
$totalDeHojas = $documento->getSheetCount();
// Escogemos la hoja en que trabajaremos (en excel las hojas comienzan desde 1, pero el id es 0)
$hojaAEscoger = 0;

$hojaActual = $documento->getSheet($hojaAEscoger);
print $chalk->bold->blue("Empezamos en la hoja con índice $hojaAEscoger\n");

# Calcular el máximo valor de la fila como entero, es decir, el limit de nuestro ciclo
$numeroMayorDeFila = $hojaActual->getHighestRow(); // Numérico
$letraMayorDeColumna = $hojaActual->getHighestColumn(); // Letra
# Convertir la letra al número de columna correspondiente
$numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

// Columna a iterar, 1 seria igual a la columna 'A'
$indiceColumna = 1;

// Conteo dinamico
$conteoPreguntas = 1;
$conteoComentarios = 1;
$seccionPreguntas = 0;

# Iterar filas con ciclo for e índices
for ($indiceFila = 1; $indiceFila <= $numeroMayorDeFila; $indiceFila++) {
    # Obtener celda por columna y fila
    $celda = $hojaActual->getCellByColumnAndRow($indiceColumna, $indiceFila);
    # Y ahora que tenemos una celda trabajamos con ella igual que antes
    # El valor, así como está en el documento
    $valorRaw = $celda->getValue();

    $preguntaString = $valorRaw;

    # Fila, que comienza en 1, luego 2 y así...
    $fila = $celda->getRow();
    # Columna, que es la A, B, C y así...
    $columna = $celda->getColumn();
        
    print "En $columna$fila tenemos el valor => $preguntaString\n"; 

    // Secciones de preguntas
    if($conteoPreguntas >= 1 and $conteoPreguntas <= 28){
        $seccionPreguntas = 30;
    }else if($conteoPreguntas >= 29 and $conteoPreguntas <= 92){
        $seccionPreguntas = 31;
    }else if($conteoPreguntas >= 93 and $conteoPreguntas <= 227){
        $seccionPreguntas = 32;
    }else if($conteoPreguntas >= 228 and $conteoPreguntas <= 237){
        $seccionPreguntas = 33;
    }else if($conteoPreguntas >= 238 and $conteoPreguntas <= 241){
        $seccionPreguntas = 34;
    }else if($conteoPreguntas >= 242 and $conteoPreguntas <= 254){
        $seccionPreguntas = 35;
    }else if($conteoPreguntas >= 255 and $conteoPreguntas <= 268){
        $seccionPreguntas = 36;
    }

    try {

        // Preparar consulta, y en vez de pasarle los valores directamente, se pasan unos placeholder
        $query = $database->connect()->prepare('INSERT INTO preguntas (id_orden_pregunta, pregunta, anotacion, tipo_pregunta, concepto_negativo_pregunta, fecha_creacion, id_grupo_preguntas, status_preguntas_ayuda, id_temario_preguntas) VALUES (:id_orden_pregunta, :pregunta, :anotacion, :tipo_pregunta, :concepto_negativo_pregunta, :fecha_creacion, :id_grupo_preguntas, :status_preguntas_ayuda, :id_temario_preguntas)');

        // Ejecutar el query reemplazando los placeholder por los valores
        $query->execute([
            'id_orden_pregunta' => $conteoPreguntas,
            'pregunta' => $preguntaString,
            'anotacion' => null,
            'tipo_pregunta' => null,
            'concepto_negativo_pregunta' => 'No cumple',
            'fecha_creacion' => date('d-m-Y H:i:s'),
            'id_grupo_preguntas' => 4,
            'status_preguntas_ayuda' => 0,
            'id_temario_preguntas' => $seccionPreguntas
        ]);

        print $chalk->bold->green("Pregunta añadida con temario $seccionPreguntas ✅\n\n");

    } catch (PDOException $e) {

        print $e->getMessage();
        return false;

    }

    $conteoPreguntas++;

}

