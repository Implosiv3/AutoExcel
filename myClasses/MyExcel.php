<?php
require_once $_SERVER ['DOCUMENT_ROOT'] . '/MyExcel_Pro/Classes/PHPExcel.php';


/**
 *  Static class made by Daniel Alcalá Valera (www.aldava.es)
 */
class MyExcel {
    function __construct() {

    }

    public static function getContent($file, $sheet) {
        $content = [];
        $references = [];
        
        $xls = PHPExcel_IOFactory::load($file);
        $xls->setActiveSheetIndex($sheet);
        $sheet = $xls->getActiveSheet();

        $mostrar = 0;

        foreach($sheet->getRowIterator() as $row) {
            foreach($row->getCellIterator() as $key => $cell) {
                if ($cell->getCalculatedValue() != null) {
                    $content[count($content)] = $cell ->getCalculatedValue();
                    $references[count($references)] = $key . $cell->getRow();                    
                }
            }
        }

        for ($i = 0; $i < count($content); $i++) {
            echo 'Found ' . $content[$i] . ' in ' . $references[$i] . '<br />';
        }
    }

    private static function round_down($number, $precision = 2) {
        $fig = (int) str_pad('1', $precision, '0');
        return (floor($number * $fig) / $fig);
    }

    private static function generateCellColumn($number) {
        $chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
        $options = [];

        $EXCEL_MAX_COLUMNS = 16384;

        /*for ($i = 0; $i < $EXCEL_MAX_COLUMNS; $i++) {
            // From 0 to 25^1 => 'A' to 'Z'
            // From 26 to (26 + 25^2) => 'AA' to 'ZZ'
            for ($x = 0; $x < ($i ))
        }*/

        echo 'Entered: ' . $number;

        $pow = 1;
        for ($i = 0; $i < 100; $i++) {
            if ($number < (count($chars) ** $pow))
                break;
            else
                $pow++;
        }

        echo 'Necesitaremos ' . $pow . ' caracteres';

        $column = '';

        $numAux = $number;
        // We get the number of letters needed
        for ($i = $pow - 1; $i >= 0; $i--) {

            /*$timesDivisible = MyExcel::round_down($number / (count($chars) ** $pow), 0);
            echo 'TimesDisible: ' . $timesDivisible;
            $nextNumber = $number % (count($chars) ** $pow);
            echo 'NextNumber: ' . $nextNumber;
            $column = $chars[$nextNumber] . $column;
            $number = $nextNumber;*/
        }

        return $column;
    }

    /**
     *  Gets the full table.
     */
    public static function readTable($file, $sheet, $firstColumn, $firstRow, $lastColumn, $lastRow) {
        $sheet = MyExcel::loadFile($file)->setActiveSheetIndex($sheet);
        
        $headers = MyExcel::readRow($sheet, $firstRow, $firstColumn, $lastColumn);
        $data = array();
        $index = 0;
        for ($i = $firstRow + 1; $i <= $lastRow; $i++) {
            $data[$index] = MyExcel::readRow($sheet, $i, $firstColumn, $lastColumn);
            $index++;
        }
   
        array_unshift($data, $headers);

        return $data;
    }



















    
    /**
     *  Reads one table from a given $sheet of the $file specified.
     */
    public static function readOneTable($file, $sheet, $firstColumn, $firstRow, $lastColumn, $lastRow) {
        $sheetLoaded = MyExcel::loadFile($file)->setActiveSheetIndex($sheet);
        return MyExcel::readTable($sheet, $firstColumn, $firstRow, $lastColumn, $lastRow);
    }

    /**
     *  Reads and format to HTML one table from a given $sheet of the $file specified.
     */
    public static function readOneTableAndFormatToHTML($file, $sheet, $firstColumn, $firstRow, $lastColumn, $lastRow) {
        $sheetLoaded = MyExcel::loadFile($file)->setActiveSheetIndex($sheet);
        return MyExcel::tableToHTML(MyExcel::getHeadersFromTable($sheetLoaded, $firstColumn, $firstRow, $lastColumn, $lastRow), MyExcel::getDataFromTable($sheetLoaded, $firstColumn, $firstRow, $lastColumn, $lastRow), "promedioControladores");
    }







    /**
     * Load the provided Excel file. Path must be like 'C:\Users\PC06\Google Drive\Taller\Denuncias'.
     */
    private static function loadFile($file) {
        $fileReadException = false;

        try {
            $filetype = PHPExcel_IOFactory::identify($file);
        } catch (Exception $e) {
            $fileReadException = true;
        }

        if ($fileReadException) {
            $file = substr($file, 0, -1);
            try {
                $filetype = PHPExcel_IOFactory::identify($file);
            } catch (Exception $e) {
                echo "No existe el archivo.";
            }
        }
        $objReader = PHPExcel_IOFactory::createReader($filetype);

        return $objReader->load($file);
    }






    /**
     *  Returns the HEADERS from a given Table (specified by first cell ($firstColumn + $firstRow) and last cell ($lastColumn + $lastRow))
     */
    public static function getHeadersFromTable($file, $sheet, $firstColumn, $firstRow, $lastColumn, $lastRow) {
        $sheet = MyExcel::loadFile($file)->setActiveSheetIndex($sheet);
        return MyExcel::readRow($sheet, $firstRow, $firstColumn, $lastColumn);
    }
    
    /**
     *  Returns the DATA (as an asociative array) from a given Table (specified by first cell ($firstColumn + $firstRow) and last cell ($lastColumn + $lastRow))
     */
    public static function getDataFromTable($sheet, $firstColumn, $firstRow, $lastColumn, $lastRow) {
        $data = array();
        
        $index = 0;
        for ($i = $firstRow + 1; $i <= $lastRow; $i++) {
            $data[$index] = MyExcel::readRow($sheet, $i, $firstColumn, $lastColumn);
            $index++;
        }
        
        $dataAsociativo = [];
        $headers = MyExcel::getHeadersFromTable($sheet, $firstColumn, $firstRow, $lastColumn, $lastRow);
        for ($i = 0; $i < count($data); $i++) {
            for ($j = 0; $j < count($data[0]); $j++) {
                $dataAsociativo[$i][$headers[$j]] = $data[$i][$j];
            }
        }
        
        // Devolvemos los resultados en array asociativo. De esta forma podremos obtener el valor del campo "Nombre" del segundo registro con la línea $data[1]["Nombre"];
        return $dataAsociativo;
    }
    
    /**
     *  Returns the plain text in the cell. If it has a formula, it is returned.
     */
    private function readCellFormula($sheet, $cellCoordinate) {
        $readCell = $sheet->getCell($cellCoordinate)->getValue();
        return $readCell;
    }
    
    /**
     *   Returns the indicated cell value.
     */
    private static function readCell($sheet, $cellCoordinate) {
        $readCell = null;
        
        // Depending on the value in cell, it can fail. That's why I try to get 2 times.
        try {
            $readCell = $sheet->getCell($cellCoordinate)->getCalculatedValue();
        } catch (Exception $e) {
            
        }
        
        if (!$readCell) {
            try {
                $readCell = $sheet->getCell($cellCoordinate)->getOldCalculatedValue();
            } catch (Exception $e) {
                
            }
        }
        
        return $readCell;
    }
    
    /**
     *  Returns the provided Row as an array (if possible)
     */
    private static function readRow($sheet, $row, $firstColumn, $lastColumn) {
        $readData = [];
        // [MUST] Test if possible ($firstIndex < $lastIndex)
        for ($i = $firstColumn; $i <= $lastColumn; $i++) {
            array_push($readData, MyExcel::readCell($sheet, $i . $row));
        }
        
        return $readData;
    }
    
    /**
     *  Returns the provided Column as an array (if possible)
     */
    private static function readColumn($sheet, $column, $firstRow, $lastRow) {
        $readData = [];
        // [MUST] Test if possible ($firstRow < $lastRow)
        for ($i = $firstRow; $i <= $lastRow; $i++) {
            array_push($readData, MyExcel::readCell($sheet, $column . $i));
        }
        
        return $readData;
    }
    
    
    
    
    
    
    /****************************/
    /*         WRITING          */
    /****************************/
    // Crea un nuevo objeto PHPExcel
    private function createExcel() {
        $objPHPExcel = new PHPExcel();

        // Establecer propiedades
        $objPHPExcel->getProperties()
        ->setCreator("Cattivo")
        ->setLastModifiedBy("Cattivo")
        ->setTitle("Documento Excel de Prueba")
        ->setSubject("Documento Excel de Prueba")
        ->setDescription("Demostracion sobre como crear archivos de Excel desde PHP.")
        ->setKeywords("Excel Office 2007 openxml php")
        ->setCategory("Pruebas de Excel");

        // Agregar Informacion
        $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', 'Valor 1')
        ->setCellValue('B1', 'Valor 2')
        ->setCellValue('C1', 'Total')
        ->setCellValue('A2', '10')
        ->setCellValue('C2', '=sum(A2:B2)');

        // Renombrar Hoja
        $objPHPExcel->getActiveSheet()->setTitle('Tecnologia Simple');

        // Establecer la hoja activa, para que cuando se abra el documento se muestre primero.
        $objPHPExcel->setActiveSheetIndex(0);

        // Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="pruebaReal.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
    }
}
