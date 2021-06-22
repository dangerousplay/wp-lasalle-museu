<?php

$valid_section_headers = [
    'REGISTRO DE ACERVO',
    'DADOS TÉCNICOS',
    'PROCEDÊNCIA',
    'DIMENSÕES',
    'FORMA DE AQUISIÇÃO',
    'ESTADO DE CONSERVAÇÃO',
    'DADOS HISTÓRICOS',
    'PARECER'
];

$ignore_table_headers = ['Cm', 'Menor', 'Maior', 'Fotografia'];
$valid_table_headers = [
    'Comprimento',
    'Espessura',
    'Diâmetro',
    'Altura',
    'Circunferência',
    'Profundidade',
    'Peso'
];

$ignore_location_headers = ['Localização', 'Saída', 'Retornar', 'Responsável'];

function read_docx($filename) {
    $current_image = '';
    $current_image_name = '';
    $current_image_size = 0;
    $content = '';
    $zip = new ZipArchive;
    if (true === $zip->open($filename)) {
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $zip_element = $zip->statIndex($i);
            $name = $zip_element['name'];
            $size = $zip_element['size'];

            if(preg_match("([^\s]+(\.(?i)(jpg|jpeg|png|gif|bmp))$)", $name) && $size > $current_image_size) {
                $current_image = $zip->getFromIndex($i);
                $current_image_name = $name;
                $current_image_size = $zip_element['size'];
                continue;
            }
            if ($name != "word/document.xml") {
                continue;
            }
            $content .= $zip->getFromIndex($i);
        }
        $zip->close();
    }
    $content = str_replace('</w:r></w:p></w:tc><w:tc>', "\n", $content);
    $content = str_replace('</w:r></w:p>', "\n", $content);
    $content = str_replace('</w:rPr></w:pPr>', "\n", $content);

    return array(
        "text" => strip_tags($content),
        "image" => [
            'data' => $current_image,
            'name' => $current_image_name
        ]
    );
}

$document_data = read_docx("test.docx");
$exploded_text = explode("\n", $document_data['text']);

$properties = array();
$is_parsing_table = false;
$is_parsing_locations = false;

$lines_to_skip = 0;
$location_header_line = 0;

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

    if (!$is_parsing_table && !$is_parsing_locations) {
        if (count($exploded_line) >= 2) {
            if (isset($current_section_header)) {
                $properties[$current_section_header][$trimmed_header] = trim(implode("", array_slice($exploded_line, 1)));
            } else {
                $properties[$trimmed_header] = trim(implode("", array_slice($exploded_line, 1)));
            }
            continue;
        }

        if ($trimmed_header == 'DIMENSÕES') {
            $is_parsing_table = true;
            continue;
        }

        if ($trimmed_header == 'PARECER') {
            $is_parsing_locations = true;
            $properties[$current_section_header] = array();
            continue;
        }
    }
    
    if ($is_parsing_table) {
        if ($trimmed_header == 'FORMA DE AQUISIÇÃO') {
            $is_parsing_table = false;
            continue;
        }

        if (in_array($trimmed_header, $ignore_table_headers)) {
            continue;
        }

        if (in_array($trimmed_header, $valid_table_headers)) {
            $current_table_header = $trimmed_header;
            $lines_to_skip = 3;
            continue;
        }

        if (!isset($current_table_header)) {
            continue;
        }

        if (!isset($properties[$current_section_header][$current_table_header . ' menor'])) {
            $properties[$current_section_header][$current_table_header . ' menor'] = trim($line);
            $lines_to_skip = 1;
            continue;
        }

        if (!isset($properties[$current_section_header][$current_table_header . ' maior'])) {
            $properties[$current_section_header][$current_table_header .' maior'] = trim($line);
            continue;
        }
    }

    if ($is_parsing_locations) {
        if ($trimmed_header == 'Referências Bibliográficas/ Fontes') {
            unset($current_section_header);
            $properties[$trimmed_header] = trim(implode("", array_slice($exploded_line, 1)));
            $is_parsing_locations = false;
            continue;
        }

        if (in_array($trimmed_header, $ignore_location_headers) || ($location_header_line == 0 && $trimmed_header == '')) {
            continue;
        }

        if ($location_header_line == 0) {
            array_push($properties[$current_section_header], array(
                $ignore_location_headers[$location_header_line] => $trimmed_header
            ));
        } else {
            $properties[$current_section_header][count($properties[$current_section_header]) - 1][$ignore_location_headers[$location_header_line]] = $trimmed_header;
        }
        $lines_to_skip = 1;

        if ($location_header_line == 3) {
            $location_header_line = 0;
            continue;
        }

        $location_header_line++;
    }
}

header('Content-Type: application/json');
echo json_encode($properties, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);