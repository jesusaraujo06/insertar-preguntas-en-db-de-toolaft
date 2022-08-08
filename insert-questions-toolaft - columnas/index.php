<?php

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

$rutaArchivo = "appseguridad.xlsx";
$documento = IOFactory::load($rutaArchivo);

# obtener conteo de hojas e iterar
$totalDeHojas = $documento->getSheetCount();

for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {
    # Obtener hoja en el índice que vaya del ciclo
    $hojaActual = $documento->getSheet($indiceHoja);
    print $chalk->bold->blue("Empecamos en la hoja con índice $indiceHoja\n");

    # Calcular el máximo valor de la fila como entero, es decir, el
    # límite de nuestro ciclo
    $numeroMayorDeFila = $hojaActual->getHighestRow(); // Numérico
    $letraMayorDeColumna = $hojaActual->getHighestColumn(); // Letra
    # Convertir la letra al número de columna correspondiente
    $numeroMayorDeColumna = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($letraMayorDeColumna);

    // Escojer solo la fila numero 1
    $indiceFila = 1;
    
    // Conteo dinamico
    $conteoPreguntas = 1;
    $conteoComentarios = 1;
    $seccionPreguntas = 0;


    
    # Iterar columnas con ciclo for e índices
    for ($indiceColumna = 6; $indiceColumna <= $numeroMayorDeColumna-1; $indiceColumna++) {
        # Obtener celda por columna y fila
        $celda = $hojaActual->getCellByColumnAndRow($indiceColumna, $indiceFila);
        # Y ahora que tenemos una celda trabajamos con ella igual que antes
        # El valor, así como está en el documento
        $valorRaw = $celda->getValue();

        # Fila, que comienza en 1, luego 2 y así...
        $fila = $celda->getRow();
        # Columna, que es la A, B, C y así...
        $columna = $celda->getColumn();

        if($valorRaw != "comentario$conteoComentarios"){

            // Dividimos el string en un array
            $extarr = explode("$conteoPreguntas. ", $valorRaw);

            // Obtenemos la pregunta, la cual se encuentra en la posicion 1 del array
            $preguntaString = empty($extarr[1]) ? null : $extarr[1];
            
            print "En $columna$fila tenemos el valor => $preguntaString\n"; 

            // Secciones de preguntas
            if($conteoPreguntas >= 1 and $conteoPreguntas <= 10){
                $seccionPreguntas = 19;
            }else if($conteoPreguntas >= 11 and $conteoPreguntas <= 20){
                $seccionPreguntas = 20;
            }else if($conteoPreguntas >= 21 and $conteoPreguntas <= 38){
                $seccionPreguntas = 21;
            }else if($conteoPreguntas >= 39 and $conteoPreguntas <= 62){
                $seccionPreguntas = 22;
            }else if($conteoPreguntas >= 63 and $conteoPreguntas <= 103){
                $seccionPreguntas = 23;
            }else if($conteoPreguntas >= 104 and $conteoPreguntas <= 124){
                $seccionPreguntas = 24;
            }else if($conteoPreguntas >= 125 and $conteoPreguntas <= 154){
                $seccionPreguntas = 25;
            }else if($conteoPreguntas >= 155 and $conteoPreguntas <= 166){
                $seccionPreguntas = 26;
            }else if($conteoPreguntas >= 167 and $conteoPreguntas <= 194){
                $seccionPreguntas = 27;
            }else if($conteoPreguntas >= 195 and $conteoPreguntas <= 217){
                $seccionPreguntas = 28;
            }else if($conteoPreguntas >= 218 and $conteoPreguntas <= 226){
                $seccionPreguntas = 29;
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
                    'id_grupo_preguntas' => 3,
                    'status_preguntas_ayuda' => 0,
                    'id_temario_preguntas' => $seccionPreguntas
                ]);

                print $chalk->bold->green("Pregunta añadida con temario $seccionPreguntas ✅\n\n");
    
            } catch (PDOException $e) {
    
                print $e->getMessage();
                return false;
    
            }

            $conteoPreguntas++;

        }else{

            $conteoComentarios++;

        }
        
    }
    
}