<?php
require_once __DIR__ . "/vendor/autoload.php";

$document = \PhpOffice\PhpWord\IOFactory::load("BUSTO.docx");

$elements = $document->getSection(0)->getElements();

$tables = array_filter($elements, function ($item) {
    return $item instanceof \PhpOffice\PhpWord\Element\Table;
});

class DocxParser {
    public static $REGISTER_TYPE = "REGISTER";
    public static $TECHNICAL_DATA_TYPE = "TECHNICAL";
    public static $ORIGIN_DATA_TYPE = "ORIGIN";
    public static $DIMENSIONS_TYPE = "DIMENSIONS";
    public static $CONSERVATION_TYPE = "CONSERVATION";

    private static function find_table_by_type(array $textsGroup): string|null
    {
        global $tables_by_title;

        foreach ($textsGroup as $texts) {
            $title = $texts[0]->getText();
            foreach ($tables_by_title as $key => $value) {
                if(!strcmp($value, $title)) {
                    return $key;
                }
            }
        }

        return null;
    }

    private static function build_text($texts) {
        return array_reduce(array_slice($texts, 1, count($texts)), function ($a, $b) {
            return $a . $b->getText();
        }, "");
    }

    private static function get_texts_group_from_table($table): array {
        return array_filter(
            array_map(function ($item) {
                return $item->getCells()[0]->getElements()[0]->getElements();
            }, $table->getRows()),
            function ($texts) {
               return count($texts) > 0;
            }
        );
    }

    public static function process_table(\PhpOffice\PhpWord\Element\Table $table) {
        $textsGroup = self::get_texts_group_from_table($table);
        $type = DocxParser::find_table_by_type($textsGroup);

        return match ($type) {
            DocxParser::$REGISTER_TYPE => self::process_register_data($textsGroup),
            DocxParser::$TECHNICAL_DATA_TYPE => self::process_technical_data($textsGroup),
            DocxParser::$ORIGIN_DATA_TYPE => self::process_origin_data($textsGroup),
            DocxParser::$DIMENSIONS_TYPE => self::process_dimensions_data($textsGroup)
        };
    }

    private static function process_register_data(array $textsGroup) {
        return 1;
    }

    private static function process_technical_data(array $textsGroup) {
        return 1;
    }

    private static function process_origin_data(array $textsGroup) {
        return 1;
    }

    private static function process_dimensions_data(array $textsGroup) {
        return 1;
    }
}

global $tables_by_title;

$tables_by_title = [
    DocxParser::$REGISTER_TYPE => "Objeto:",
    DocxParser::$TECHNICAL_DATA_TYPE => "Título",
    DocxParser::$ORIGIN_DATA_TYPE => "Localidade:",
    DocxParser::$DIMENSIONS_TYPE => "Comprimento",
    DocxParser::$CONSERVATION_TYPE => "Conservação:"
];

foreach ($tables as $table) {
    DocxParser::process_table($table);
}

$table1 = $tables[8];

$row = $table1->getRows()[0];

$a = 123;