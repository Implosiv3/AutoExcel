<?php

class MyExcelFormatter {
    public static function getRowFromValue($data, $headers, $header, $value) {
        echo 'Looking for ' . $value;
        $result = null;
        for ($i = 0; $i < count($data); $i++) {
            if ($data[$i][$header] == $value) {
                $result = $data[$i];
                break;
            }
        }
        return $result;
    }

    public static function getRowIndexFromValue($data, $headers, $header, $value) {
        echo 'Looking for ' . $value;
        $result = null;
        for ($i = 0; $i < count($data); $i++) {
            if ($data[$i][$header] == $value) {
                $result = $i;
                break;
            }
        }
        return $i;
    }




    public static function getRowResultFromTable($table, $fieldToFindIndex, $valueToFind) {
        $rowResult = null;

        for ($i = 0; $i < count($table); $i++) {
            if ($table[$i][$fieldToFindIndex] == $valueToFind) {
                // We found what we were looking for
                $rowResult = $table[$i];
                break;
            }
        }

        return $rowResult;
    }

    public static function getRowsResultFromTable($table, $fieldToFindIndex, $valuesToFind) {
        $rowsResult = array();

        for ($i = 0; $i < count($table); $i++) {
            for ($x = 0; $x < count($valuesToFind); $x++) {
                if ($table[$i][$fieldToFindIndex] == $valuesToFind[$x]) {
                    // We found one of them
                    array_push($rowsResult, $table[$i]);
                    echo 'Found ' . $valuesToFind[$x];
                    array_splice($valuesToFind, $x, 1);
                    break;
                }
            }
            if (count($valuesToFind) == 0)
                break;
        }

        return $rowsResult;
    }

    public static function getTableWithLessColumns($table, $columns = null) {
        if ($columns == null)
            return $table;

        $reducedTable = array();

        
        $rowNumber = 0;
        $colNumber = 0;
        for ($i = 0; $i < count($table); $i++) {
            $colNumber = 0;
            for ($j = 0; $j < count($table[0]); $j++) {
                for ($column = 0; $column < count($columns); $column++) {
                    if ($j == $columns[$column]) {
                        $reducedTable[$rowNumber][$colNumber] = $table[$i][$j];
                        $colNumber++;
                        break;
                    }
                }
            }
            $rowNumber++;
        }

        return $reducedTable;
    }

    /**
     *  Returns a table as HTML providing his $headers and $data (and an optional id field).
     */
    // Esto de $onlyThisRow no debe estar aquí, es de prueba para mostrar solo una fila. Está aquí metio con calzador.
    public static function tableToHTML($headers, $data, $id = null, $classes = null, $caption = null) {
        $tableHTML = '<div id="d1" class="d1"><table style="text-align: center;"';
        if ($id)
            $tableHTML .= ' id="' . $id . '"';
        if ($classes) {
            $tableHTML .= ' class="';
            for ($i = 0; $i < count($classes); $i++) {
                $tableHTML .= $classes[$i];
                if ($i != (count($classes) - 1))
                    $tableHTML .= ' ';
            }
            $tableHTML .= '"';
        }
        $tableHTML .= '>';
        if ($caption)
            $tableHTML .= '<caption>' . $caption . '</caption>';
        $tableHTML .= '<thead class="head"><tr>';
        for ($column = 0; $column < count($headers); $column++) {
            $tableHTML .= '<th>' . $headers[$column] . '</th>';
        }
        $tableHTML .= '</tr></thead><tbody>';
        for ($row = 0; $row < count($data); $row++) {
            $tableHTML .= '<tr>';
            for ($column = 0; $column < count($data[0]); $column++) {
                $tableHTML .= '<td>' . $data[$row][$headers[$column]] . '</td>';
            }
            $tableHTML .= '</tr>';
        }
        
        $tableHTML .= '</tbody></table></div>';
        
        return $tableHTML;
    }

    /**
     *  Returns a table as HTML providing his $headers and $data (and an optional id field).
     */
    // Esto de $onlyThisRow no debe estar aquí, es de prueba para mostrar solo una fila. Está aquí metio con calzador.
    public static function tableToHTML_index($headers, $data, $id = null, $classes = null, $caption = null) {
        $tableHTML = '<table style="text-align: center;"';
        if ($id)
            $tableHTML .= ' id="' . $id . '"';
        if ($classes) {
            $tableHTML .= ' class="';
            for ($i = 0; $i < count($classes); $i++) {
                $tableHTML .= $classes[$i];
                if ($i != (count($classes) - 1))
                    $tableHTML .= ' ';
            }
            $tableHTML .= '"';
        }
        $tableHTML .= '>';
        if ($caption)
            $tableHTML .= '<caption>' . $caption . '</caption>';
        $tableHTML .= '<thead class="head"><tr>';
        for ($column = 0; $column < count($headers); $column++) {
            $tableHTML .= '<th>' . $headers[$column] . '</th>';
        }
        $tableHTML .= '</tr></thead><tbody>';
        for ($row = 0; $row < count($data); $row++) {
            $tableHTML .= '<tr>';
            for ($column = 0; $column < count($data[0]); $column++) {
                $tableHTML .= '<td>' . $data[$row][$column] . '</td>';
            }
            $tableHTML .= '</tr>';
        }
        
        $tableHTML .= '</tbody></table>';
        
        return $tableHTML;
    }
}