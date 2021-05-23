<?php
$valid_section_headers = [
    'REGISTRO DE ACERVO',
    'DADOS TÉCNICOS',
    'PROCEDÊNCIA',
    'DIMENSÕES',
    'FORMA DE AQUISIÇÃO',
    'ESTADO DE CONSERVAÇÃO',
    'DADOS HISTÓRICOS'
];
$ignore_headers = ['Cm', 'Menor', 'Maior', 'Fotografia'];
$valid_table_headers = [
    'Comprimento',
    'Espessura',
    'Diâmetro',
    'Altura',
    'Circunferência',
    'Profundidade',
    'Peso'
];

function read_docx($filename) {
    $striped_content = '';
    $content = '';
    //print($filename);
    $zip = zip_open($filename);
    if (!$zip || is_numeric($zip)) return false;
    while ($zip_entry = zip_read($zip)) {
        if (zip_entry_open($zip, $zip_entry) == FALSE) continue;
        if (zip_entry_name($zip_entry) != "word/document.xml") continue;
        $content .= zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));
        zip_entry_close($zip_entry);
    }
    zip_close($zip);
    $content = str_replace('</w:r></w:p></w:tc><w:tc>', "\n", $content);
    $content = str_replace('</w:r></w:p>', "\n", $content);
    $content = str_replace('</w:rPr></w:pPr>', "\n", $content);
    $striped_content = strip_tags($content);
    //print($striped_content);
    return $striped_content;
}

$text = read_docx("BUSTO.docx");
$exploded_text = explode("\n", $text);

$properties = array();
$is_parsing_table = false;
$lines_to_skip = 0;
$current_section_header = 'REGISTRO DE ACERVO';
foreach ($exploded_text as $line) {
    $exploded_line = explode(":", $line);
    $trimmed_header = trim($exploded_line[0]);

    if ($lines_to_skip > 0) {
        $lines_to_skip--;
        continue;
    }

    if (in_array($trimmed_header, $valid_section_headers)) {
        $current_section_header = ucfirst(mb_strtolower($trimmed_header));
    }

    if (!$is_parsing_table) {
        if (count($exploded_line) >= 2) {
            $properties[$current_section_header][$trimmed_header] = trim(implode("", array_slice($exploded_line, 1)));
            continue;
        }

        if ($trimmed_header == 'DIMENSÕES') {
            $is_parsing_table = true;
            continue;
        }
    } else {
        if ($trimmed_header == 'FORMA DE AQUISIÇÃO') {
            $is_parsing_table = false;
            continue;
        }

        if (in_array($trimmed_header, $ignore_headers)) {
            continue;
        }

        if ($trimmed_header == 'FORMA DE AQUISIÇÃO') {
            $is_parsing_table = true;
            continue;
        }

        if (in_array($trimmed_header, $valid_table_headers)) {
            $current_table_header = $trimmed_header;
            $lines_to_skip = 2;
            continue;
        }

        if (!isset($current_table_header)) {
            continue;
        }

        if (!isset($properties[$current_section_header][$current_table_header]['menor'])) {
            $properties[$current_section_header][$current_table_header]['menor'] = trim($line);
            continue;
        }

        if (!isset($properties[$current_section_header][$current_table_header]['maior'])) {
            $properties[$current_section_header][$current_table_header]['maior'] = trim($line);
            continue;
        }
    }
}

header('Content-Type: application/json');
//echo json_encode($properties, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
//$list = ----your array

$csv = "file.csv";
//$file = json_encode($properties);
$file = $properties;
print(gettype($file));
//$file = [    "foo" => "bar",    "bar" => "foo",];
//echo($file);
//echo json_encode($file, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
//echo($properties);

jsonToCSV($file, $csv);

function jsonToCSV($jfilename, $cfilename)
{
    $data = ["Registro de acervo", "Dados técnicos", "Procedência", "Dimensões", "Forma de aquisição", "Estado de conservação", "Dados históricos"];
    $fp = fopen("file.csv", 'w');
    fputcsv($fp, $data);
    foreach ($jfilename as $fields)
    {
      print($jfilename);
      fputcsv($fp, $fields);
      foreach ($fields as $key) {
        fputcsv($fp, $key);
      }

    }
    fclose($fp);
}

//$object = json_decode($properties);

//$list = get_field($object, $field, $default);
//$fp = fopen('file.csv', 'w');

//foreach ($list as $fields) {
//    fputcsv($fp,get_object_vars($fields));
//}

//fclose($fp);
