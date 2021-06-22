<?php

namespace Lasalle\Tainacan;

use Tainacan\Entities\Collection;
use Tainacan\Entities\Item;
use Tainacan\Entities\Item_Metadata_Entity;
use Tainacan\Entities\Metadatum;
use Tainacan\Entities\Taxonomy;
use Tainacan\Importer\Importer;
use Tainacan\Media;
use Tainacan\Repositories\Collections;
use Tainacan\Repositories\Item_Metadata;
use Tainacan\Repositories\Items;
use Tainacan\Repositories\Metadata;
use Tainacan\Repositories\Terms;
use WP_Query;
use ZipArchive;


if (!function_exists('str_putcsv')) {
    function str_putcsv($input, $delimiter = ',', $enclosure = '"')
    {
        $fp = fopen('php://temp', 'r+b');
        fputcsv($fp, $input, $delimiter, $enclosure);
        rewind($fp);
        $data = rtrim(stream_get_contents($fp), "\n");
        fclose($fp);
        return $data;
    }
}

class Docx_Importer extends Importer
{
    private static $valid_section_headers = [
        'REGISTRO DE ACERVO',
        'DADOS TÉCNICOS',
        'PROCEDÊNCIA',
        'DIMENSÕES',
        'FORMA DE AQUISIÇÃO',
        'ESTADO DE CONSERVAÇÃO',
        'DADOS HISTÓRICOS',
        'PARECER'
    ];

    private static $ignore_table_headers = ['Cm', 'Menor', 'Maior', 'Fotografia'];
    private static $valid_table_headers = [
        'Comprimento',
        'Espessura',
        'Diâmetro',
        'Altura',
        'Circunferência',
        'Profundidade',
        'Peso'
    ];

    private static $ignore_location_headers = ['Localização', 'Saída', 'Retornar', 'Responsável'];

    private static $PRIVATE_SUFFIX = "- PRIVADO";

    private static $METADATUM_MAPPING = [
        'Registro de acervo'    => [
            'Nª do livro Tombo'                  => [ 'type' => 'text', 'private' => true ],
            'Nª de Registro'                     => [ 'type' => 'text', 'private' => true ],
            'Outros números'                     => [ 'type' => 'text', 'private' => true ],
            'Localização no Museu'               => [ 'type' => 'text', 'private' => true ]
        ],
        'Dados técnicos'                         => [
            'Data da confecção do material'      => [ 'type' => 'date', 'private' => false ],
            'Autor/Autoridade'                   => [ 'type' => 'text', 'private' => true ],
            'Descrição intrínseca'               => [ 'type' => 'text', 'private' => true ],
            'Matéria Prima'                      => [ 'type' => 'text', 'private' => true ],
            'Inscrição/ Marcas/ Títulos'         => [ 'type' => 'text', 'private' => true ],
            'Técnica de manufatura'              => [ 'type' => 'text', 'private' => true ],
            'Técnica decorativa'                 => [ 'type' => 'text', 'private' => true ],
            'Representação/ Decoração'           => [ 'type' => 'text', 'private' => true ],
            'Observações/Outras Características' => [ 'type' => 'textarea', 'private' => true ]
        ],
        'Procedência' => [
            'Município'                          => [ 'type' => 'text', 'private' => true ],
            'Sítio'                              => [ 'type' => 'text', 'private' => true ],
            'Localidade'                         => [ 'type' => 'text', 'private' => true ],
            'Estado'                             => [ 'type' => 'text', 'private' => true ],
            'Região'                             => [ 'type' => 'text', 'private' => true ],
            'Proprietário'                       => [ 'type' => 'text', 'private' => true ],
        ],
        'Forma de aquisição' => [
            'Data da Aquisição'                  => [ 'type' => 'date', 'private' => true ],
            'Doador'                             => [ 'type' => 'text', 'private' => true ],
            'Último Proprietário'                => [ 'type' => 'text', 'private' => true ],
            'Personalidade/ Pessoa'              => [ 'type' => 'text', 'private' => true ],
            'Outras Informações'                 => [ 'type' => 'textarea', 'private' => true ],
        ],
        'Estado de conservação' => [
            'Descrição'                          => [ 'type' => 'textarea', 'private' => true ]
        ],
        'Dados históricos' => [
            'Histórico'                          => [ 'type' => 'textarea', 'private' => false ]
        ],
        'Parecer' => [
            'Localização'                        => [ 'type' => 'text', 'private' => true ],
            'Saída'                              => [ 'type' => 'date', 'private' => true ],
            'Retornar'                           => [ 'type' => 'date', 'private' => true ],
            'Responsável'                        => [ 'type' => 'text', 'private' => true ],
        ],
        'Outros' => [
            'Referências Bibliográficas/ Fontes' => [ 'type' => 'text', 'private' => true ],
            'Repetidos/ Duplos'                  => [ 'type' => 'text', 'private' => true ]
        ]
    ];

    private $items_repo;

    public function __construct($attributes = array())
    {
        parent::__construct($attributes);
        $this->items_repo = Items::get_instance();
        $this->set_default_options([
            'delimiter' => ',',
            'multivalued_delimiter' => '||',
            'encode' => 'utf8',
            'enclosure' => '"'
        ]);
    }

    private static function create_image_file(array $image): string
    {
        $extension = explode(".", $image['name'])[1];

        $tmp_image_filename = tempnam("/tmp", "IMP") . "." . $extension;

        $tmp_image = fopen($tmp_image_filename, 'w');

        fwrite($tmp_image, $image['data']);
        fclose($tmp_image);

        return $tmp_image_filename;
    }

    /**
     * alter the default options
     */
    public function set_option($key, $value)
    {
        $this->default_options[$key] = $value;
    }

    /**
     * @inheritdoc
     */
    public function get_source_metadata()
    {
        if (($handle = fopen($this->tmp_file, "r")) !== false) {
            if ($this->get_option('enclosure') && strlen($this->get_option('enclosure')) > 0) {
                $rawColumns = $this->handle_enclosure($handle);
            } else {
                $rawColumns = fgetcsv($handle, 0, $this->get_option('delimiter'));
            }

            $columns = [];

            if ($rawColumns) {
                foreach ($rawColumns as $index => $rawColumn) {
                    if (strpos($rawColumn, 'special_') === 0) {
                        if ($rawColumn === 'special_document') {
                            $this->set_option('document_index', $index);
                        } else if ($rawColumn === 'special_attachments' ||
                            $rawColumn === 'special_attachments|APPEND' ||
                            $rawColumn === 'special_attachments|REPLACE') {
                            $this->set_option('attachment_index', $index);
                            $attachment_type = explode('|', $rawColumn);
                            $this->set_option('attachment_operation_type', sizeof($attachment_type) == 2 ? $attachment_type[1] : 'APPEND');
                        } else if ($rawColumn === 'special_item_status') {
                            $this->set_option('item_status_index', $index);
                        } else if ($rawColumn === 'special_item_id') {
                            $this->set_option('item_id_index', $index);
                        } else if ($rawColumn === 'special_comment_status') {
                            $this->set_option('item_comment_status_index', $index);
                        }
                    } else {
                        if (preg_match('/.*\|compound\(.*\)/', $rawColumn)) {
                            $data = preg_split("/[()]+/", $rawColumn, -1, PREG_SPLIT_NO_EMPTY);
                            $parent = $data[0] . (isset($data[2]) ? $data[2] : '');
                            $columns[] = [$parent => explode($this->get_option('delimiter'), $data[1])];
                        } else {
                            $columns[] = $rawColumn;
                        }
                    }
                }
                return $columns;
            }
        }
        return [];
    }

    public function get_source_special_fields()
    {
        if (($handle = fopen($this->tmp_file, "r")) !== false) {
            if ($this->get_option('enclosure') && strlen($this->get_option('enclosure')) > 0) {
                $rawColumns = $this->handle_enclosure($handle);
            } else {
                $rawColumns = fgetcsv($handle, 0, $this->get_option('delimiter'));
            }

            $columns = [];

            if ($rawColumns) {
                foreach ($rawColumns as $index => $rawColumn) {
                    if (strpos($rawColumn, 'special_') === 0) {
                        if (in_array($rawColumn, ['special_document', 'special_attachments', 'special_item_status', 'special_item_id', 'special_comment_status', 'special_attachments|APPEND', 'special_attachments|REPLACE'])) {
                            $columns[] = $rawColumn;
                        }
                    }
                }
                if (!empty($columns))
                    return $columns;
            }
        }
        return false;
    }

    /**
     *
     * returns all header including special
     */
    public function raw_source_metadata()
    {
        if (($handle = fopen($this->tmp_file, "r")) !== false) {
            if ($this->get_option('enclosure') && strlen($this->get_option('enclosure')) > 0) {
                return $this->handle_enclosure($handle);
            } else {
                return fgetcsv($handle, 0, $this->get_option('delimiter'));
            }
        }
        return false;
    }

    /**
     * @inheritdoc
     */
    public function process_item($index, $collection_definition)
    {
        $processedItem = [];
        $compoundHeaders = [];
        $headers = array_map(function ($header) use (&$compoundHeaders) {
            if (preg_match('/.*\|compound\(.*\)/', $header)) {
                $data = preg_split("/[()]+/", $header, -1, PREG_SPLIT_NO_EMPTY);
                $header = $data[0] . (isset($data[2]) ? $data[2] : '');
                $compoundHeaders[$header] = $data[1];
                return $header;
            }
            return $header;
        }, $this->raw_source_metadata());

        $item_line = (int)$index + 2;

        $this->add_log('Processing item on line ' . $item_line);
        $this->add_log('Target collection: ' . $collection_definition['id']);

        if (($handle = fopen($this->tmp_file, "r")) !== false) {
            $file = $handle;
        } else {
            $this->add_error_log(' Error reading the file ');
            return false;
        }

        if ($index === 0) {
            // moves the pointer forward
            fgetcsv($file, 0, $this->get_option('delimiter'));
        } else {
            //get the pointer
            $csv_pointer = $this->get_transient('csv_pointer');
            if ($csv_pointer) {
                fseek($file, $csv_pointer);
            }
        }

        $this->add_transient('csv_last_pointer', ftell($file)); // add reference to post_process item in after_inserted_item()

        if ($this->get_option('enclosure') && strlen($this->get_option('enclosure')) > 0) {
            $values = $this->handle_enclosure($file);
        } else {
            $values = fgetcsv($file, 0, $this->get_option('delimiter'));
        }

        $this->add_transient('csv_pointer', ftell($file)); // add reference for insert

        if (count($headers) !== count($values)) {
            $string = (is_array($values)) ? implode('::', $values) : $values;

            $this->add_error_log(' Mismatch count headers and row columns ');
            $this->add_error_log(' Headers count: ' . count($headers));
            $this->add_error_log(' Values count: ' . count($values));
            $this->add_error_log(' enclosure : ' . $this->get_option('enclosure'));
            $this->add_error_log(' Headers value:' . wp_json_encode($headers));
            $this->add_error_log(' Values: ' . wp_json_encode($values));
            $this->add_error_log(' Values string: ' . $string);
            return false;
        }

        if (is_numeric($this->get_option('item_id_index'))) {
            $this->handle_item_id($values);
        }
        foreach ($collection_definition['mapping'] as $metadatum_id => $header) {
            $column = null;
            foreach ($headers as $indexRaw => $headerRaw) {
                if ((is_array($header) && $headerRaw === key($header)) || ($headerRaw === $header)) {
                    $column = $indexRaw;
                }
            }

            if (is_null($column))
                continue;

            $valueToInsert = $this->handle_encoding($values[$column]);

            $metadatum = new Metadatum($metadatum_id);
            if ($metadatum->get_metadata_type() == 'Tainacan\Metadata_Types\Compound') {
                $valueToInsert = $metadatum->is_multiple()
                    ? explode($this->get_option('multivalued_delimiter'), $valueToInsert)
                    : [$valueToInsert];

                $key = key($header);
                $returnValue = [];
                foreach ($valueToInsert as $index => $metadatumValue) {
                    $childrenHeaders = str_getcsv($compoundHeaders[$key], $this->get_option('delimiter'), $this->get_option('enclosure'));
                    $childrenValue = str_getcsv($metadatumValue, $this->get_option('delimiter'), $this->get_option('enclosure'));

                    if (sizeof($childrenHeaders) != sizeof($childrenValue)) {
                        $this->add_error_log("Children headers: " . wp_json_encode($childrenHeaders));
                        $this->add_error_log('Mismatch count headers childrens and row columns. file value:' . $metadatumValue);
                        return false;
                    }
                    $tmp = [];
                    foreach ($childrenValue as $i => $value) {
                        $tmp[$childrenHeaders[$i]] = $value;
                    }
                    $returnValue[] = $tmp;
                }
                $processedItem[$key] = $returnValue;
            } else {
                $processedItem[$header] = ($metadatum->is_multiple()) ?
                    explode($this->get_option('multivalued_delimiter'), $valueToInsert) : $valueToInsert;
            }
        }
        if (!empty($this->get_option('document_index'))) $processedItem['special_document'] = '';
        if (!empty($this->get_option('attachment_index'))) $processedItem['special_attachments'] = '';
        if (!empty($this->get_option('item_status_index'))) $processedItem['special_item_status'] = '';
        if (!empty($this->get_option('item_comment_status_index'))) $processedItem['special_comment_status'] = '';

        $this->add_log('Success processing index: ' . $index);
        return $processedItem;
    }

    /**
     * @inheritdoc
     */
    public function after_inserted_item($inserted_item, $collection_index)
    {
        $column_document = $this->get_option('document_index');
        $column_attachment = $this->get_option('attachment_index');
        $column_item_status = $this->get_option('item_status_index');
        $column_item_comment_status = $this->get_option('item_comment_status_index');

        if (!empty($column_document) || !empty($column_attachment) || !empty($column_item_status)) {

            if (($handle = fopen($this->tmp_file, "r")) !== false) {
                $file = $handle;
            } else {
                $this->add_error_log(' Error reading the file ');
                return false;
            }

            $csv_pointer = $this->get_transient('csv_last_pointer');
            fseek($file, $csv_pointer);

            if ($this->get_option('enclosure') && strlen($this->get_option('enclosure')) > 0) {
                $values = $this->handle_enclosure($file);
            } else {
                $values = fgetcsv($file, 0, $this->get_option('delimiter'));
            }

            if (is_array($values) && !empty($column_document)) {
                $this->handle_document($values[$column_document], $inserted_item);
            }

            if (is_array($values) && !empty($column_attachment)) {
                $this->handle_attachment($values[$column_attachment], $inserted_item);
            }

            if (is_array($values) && !empty($column_item_status)) {
                $this->handle_item_status($values[$column_item_status], $inserted_item);
            }

            if (is_array($values) && !empty($column_item_comment_status)) {
                $this->handle_item_comment_status($values[$column_item_comment_status], $inserted_item);
            }
        }
    }

    /**
     * @inheritdoc
     */
    public function get_source_number_of_items()
    {
        if (isset($this->tmp_file) && file_exists($this->tmp_file) && ($handle = fopen($this->tmp_file, "r")) !== false) {
            $cont = 0;
            while (($data = fgetcsv($handle, 0, $this->get_option('delimiter'))) !== false) {
                $cont++;
            }
            // does not count the header
            return $cont - 1;
        }
        return false;
    }

    public function options_form()
    {

    }

    /**
     * get the encode option and return as expected
     */
    private function handle_encoding($string)
    {
        switch ($this->get_option('encode')) {
            case 'utf8':
                return $string;
            case 'iso88591':
                return utf8_encode($string);
            default:
                return $string;
        }
    }

    /**
     * method responsible to insert the item document
     */
    private function handle_document($column_value, $item_inserted)
    {
        $TainacanMedia = Media::get_instance();
        $this->items_repo->disable_logs();

        if (strpos($column_value, 'url:') === 0) {
            $correct_value = trim(substr($column_value, 4));
            $item_inserted->set_document($correct_value);
            $item_inserted->set_document_type('url');

            if ($item_inserted->validate()) {
                $item_inserted = $this->items_repo->update($item_inserted);
            }
        } else if (strpos($column_value, 'text:') === 0) {
            $correct_value = trim(substr($column_value, 5));
            $item_inserted->set_document($correct_value);
            $item_inserted->set_document_type('text');

            if ($item_inserted->validate()) {
                $item_inserted = $this->items_repo->update($item_inserted);
            }
        } else if (strpos($column_value, 'file:') === 0) {
            $correct_value = trim(substr($column_value, 5));
            if (isset(parse_url($correct_value)['scheme'])) {
                $id = $TainacanMedia->insert_attachment_from_url($correct_value, $item_inserted->get_id());

                if (!$id) {
                    $this->add_error_log('Error in Document file imported from URL ' . $correct_value);
                    return false;
                }

                $item_inserted->set_document($id);
                $item_inserted->set_document_type('attachment');
                $this->add_log('Document file URL imported from ' . $correct_value);

                if ($item_inserted->validate()) {
                    $item_inserted = $this->items_repo->update($item_inserted);
                }
            } else {
                $server_path_files = trailingslashit($this->get_option('server_path'));
                $id = $TainacanMedia->insert_attachment_from_file($server_path_files . $correct_value, $item_inserted->get_id());

                if (!$id) {
                    $this->add_error_log('Error in Document file imported from server ' . $correct_value);
                    return false;
                }

                $item_inserted->set_document($id);
                $item_inserted->set_document_type('attachment');
                $this->add_log('Document file in Server imported from ' . $correct_value);

                if ($item_inserted->validate()) {
                    $item_inserted = $this->items_repo->update($item_inserted);
                }
            }
        }

        $thumb_id = $this->items_repo->get_thumbnail_id_from_document($item_inserted);
        if (!is_null($thumb_id)) {
            $this->add_log('Setting item thumbnail: ' . $thumb_id);
            set_post_thumbnail($item_inserted->get_id(), (int)$thumb_id);
        }

        $this->items_repo->enable_logs();
        return true;
    }

    /**
     * method responsible to insert the item document
     */
    private function handle_attachment($column_value, $item_inserted)
    {
        $TainacanMedia = Media::get_instance();
        $this->items_repo->disable_logs();

        switch ($this->get_option('attachment_operation_type')) {
            case 'APPEND':
                $this->add_log('Attachment APPEND file ');
                break;
            case 'REPLACE':
                $this->add_log('Attachment REPLACE file ');
                $args['post_parent'] = $item_inserted->get_id();
                $args['post_type'] = 'attachment';
                $args['post_status'] = 'any';
                $args['post__not_in'] = [$item_inserted->get_document()];
                $posts_query = new WP_Query();
                $query_result = $posts_query->query($args);
                foreach ($query_result as $post) {
                    wp_delete_attachment($post->ID, true);
                }
                break;
        }

        $attachments = explode($this->get_option('multivalued_delimiter'), $column_value);
        if ($attachments) {
            foreach ($attachments as $attachment) {
                if (empty($attachment)) continue;
                if (isset(parse_url($attachment)['scheme'])) {
                    $id = $TainacanMedia->insert_attachment_from_url($attachment, $item_inserted->get_id());
                    if (!$id) {
                        $this->add_error_log('Error in Attachment file imported from URL ' . $attachment);
                        return false;
                    }
                    $this->add_log('Attachment file URL imported from ' . $attachment);
                    continue;
                }

                $server_path_files = trailingslashit($this->get_option('server_path'));
                $id = $TainacanMedia->insert_attachment_from_file($server_path_files . $attachment, $item_inserted->get_id());

                if (!$id) {
                    $this->add_log('Error in Attachment file imported from server ' . $attachment);
                    continue;
                }

                $this->add_log('Attachment file in Server imported from ' . $attachment);
            }
        }
        $this->items_repo->enable_logs();
    }

    /**
     * @param $file resource the csv file uploaded
     */
    private function handle_enclosure(&$file)
    {
        return fgetcsv($file, 0, $this->get_option('delimiter'), $this->get_option('enclosure'));
    }

    /**
     * @param $status string the item status
     */
    private function handle_item_status($status, $item_inserted)
    {
        //if ( in_array( $status, array( 'auto-draft', 'draft', 'pending', 'future', 'publish', 'trash', 'inherit' ) ) ) {

        $status = ($status == 'public') ? 'publish' : $status;
        $item_inserted->set_status($status);
        if ($item_inserted->validate()) {
            $item_inserted = $this->items_repo->update($item_inserted);
        }
        //}
    }

    /**
     * @param $comment_status string the item comment status
     */
    private function handle_item_comment_status($comment_status, $item_inserted)
    {
        if (!in_array($comment_status, array('open', 'closed'))) {
            $comment_status = 'closed';
        }

        $item_inserted->set_comment_status($comment_status);
        if ($item_inserted->validate()) {
            $item_inserted = $this->items_repo->update($item_inserted);
        }
    }

    /**
     * @param $status string the item ID
     */
    private function handle_item_id($values)
    {
        $item_id_index = $this->get_option('item_id_index');
        if (is_numeric($item_id_index) && isset($values[intval($item_id_index)])) {
            $this->add_transient('item_id', $values[intval($item_id_index)]);
            $this->add_transient('item_action', $this->get_option('repeated_item'));
        }
    }

    /**
     * insert processed item from source to Tainacan
     *
     * @param array $processed_item Associative array with metadatum source's as index with
     *                              its value or values
     * @param integer $collection_index The index in the $this->collections array of the collection the item is being inserted into
     *
     * @return Item Item inserted
     */
    public function insert($processed_item, $collection_index)
    {
        remove_action('post_updated', 'wp_save_post_revision');
        $collections = $this->get_collections();
        $collection_definition = isset($collections[$collection_index]) ? $collections[$collection_index] : false;
        if (!$collection_definition || !is_array($collection_definition) || !isset($collection_definition['id']) || !isset($collection_definition['mapping'])) {
            $this->add_error_log('Collection misconfigured');
            return false;
        }

        $collection = Collections::get_instance()->fetch($collection_definition['id']);

        if ($collection instanceof Collection && !$collection->user_can('edit_items')) {
            $this->add_error_log(__("You don't have permission to create items in this collection.", 'tainacan'));
            return false;
        }

        $Tainacan_Metadata = Metadata::get_instance();
        $Tainacan_Item_Metadata = Item_Metadata::get_instance();
        $Tainacan_Items = Items::get_instance();

        // $Tainacan_Items->disable_logs();
        // $Tainacan_Metadata->disable_logs();
        // $Tainacan_Item_Metadata->disable_logs();

        $itemMetadataArray = [];

        $updating_item = false;

        $Tainacan_Item_Metadata->disable_logs();
        if (is_numeric($this->get_transient('item_id'))) {
            $item = $Tainacan_Items->fetch((int)$this->get_transient('item_id'));
            if ($item instanceof Item && ($item->get_collection() == null || $item->get_collection()->get_id() != $collection->get_id())) {
                $this->add_log('item with ID ' . $this->get_transient('item_id') . ' not found in collection ' . $collection->get_name());
                $item = new Item();
            }
        } else {
            $item = new Item();
        }

        if (is_numeric($this->get_transient('item_id'))) {
            if ($item instanceof Item && $item->get_id() == $this->get_transient('item_id')) {
                if (!$item->can_edit()) {
                    $this->add_error_log("You don't have permission to edit item:" . $item->get_id());
                    return $item;
                }
                $this->add_log('item will be updated ID:' . $item->get_id());
                $updating_item = true;
                // When creating a new item, disable log for each metadata to speed things up
                $Tainacan_Item_Metadata->enable_logs();
            } else {
                $this->add_log('item with ID ' . $this->get_transient('item_id') . ' not found. Unable to update. Creating a new one.');
                $item = new Item();
            }

        }

        if ($this->get_transient('item_id') && $item instanceof Item && is_numeric($item->get_id()) && $item->get_id() > 0 && $this->get_transient('item_action') == 'ignore') {
            $this->add_log('Ignoring repeated Item');
            return $item;
        }

        if (is_array($processed_item)) {
            foreach ($processed_item as $metadatum_source => $values) {

                if ($metadatum_source == 'special_document' ||
                    $metadatum_source == 'special_attachments' ||
                    $metadatum_source == 'special_item_status' ||
                    $metadatum_source == 'special_comment_status') {
                    $special_columns = true;
                    continue;
                }

                foreach ($collection_definition['mapping'] as $id => $value) {
                    if ((is_array($value) && key($value) == $metadatum_source) || ($value == $metadatum_source))
                        $tainacan_metadatum_id = $id;
                }
                $metadatum = $Tainacan_Metadata->fetch($tainacan_metadatum_id);

                if ($this->is_empty_value($values)) continue;

                if ($metadatum instanceof Metadatum) {
                    $singleItemMetadata = new Item_Metadata_Entity($item, $metadatum); // *empty item will be replaced by inserted in the next foreach
                    if ($metadatum->get_metadata_type() == 'Tainacan\Metadata_Types\Taxonomy') {
                        if (!is_array($values)) {
                            $tmp = $this->insert_hierarchy($metadatum, $values);
                            if ($tmp !== false) {
                                $singleItemMetadata->set_value($tmp);
                            }
                        } else {
                            $terms = [];
                            foreach ($values as $k => $v) {
                                $tmp = $this->insert_hierarchy($metadatum, $v);
                                if ($tmp !== false) {
                                    $terms[] = $tmp;
                                }
                            }
                            $singleItemMetadata->set_value($terms);
                        }
                    } elseif ($metadatum->get_metadata_type() == 'Tainacan\Metadata_Types\Compound') {
                        $children_mapping = $collection_definition['mapping'][$tainacan_metadatum_id][$metadatum_source];
                        $singleItemMetadata = [];
                        foreach ($values as $compoundValue) {
                            $tmp = [];
                            foreach ($children_mapping as $tainacan_children_metadatum_id => $tainacan_children_header) {
                                $metadatumChildren = $Tainacan_Metadata->fetch($tainacan_children_metadatum_id, 'OBJECT');
                                $compoundItemMetadata = new Item_Metadata_Entity($item, $metadatumChildren);
                                $compoundItemMetadata->set_value($compoundValue[$tainacan_children_header]);
                                $tmp[] = $compoundItemMetadata;
                            }
                            $singleItemMetadata[] = $tmp;
                        }
                    } else {
                        $singleItemMetadata->set_value($values);
                    }
                    $itemMetadataArray[] = $singleItemMetadata;
                } else {
                    $this->add_error_log('Metadata ' . $metadatum_source . ' not found');
                }
            }
        }

        if ((!empty($itemMetadataArray) || $special_columns) && $collection instanceof Collection) {
            $item->set_collection($collection);
            if ($item->validate()) {
                $insertedItem = $Tainacan_Items->insert($item);
            } else {
                $this->add_error_log('Error inserting Item Title: ' . $item->get_title());
                $this->add_error_log($item->get_errors());
                return false;
            }

            foreach ($itemMetadataArray as $itemMetadata) {
                if ($itemMetadata instanceof Item_Metadata_Entity) {
                    $itemMetadata->set_item($insertedItem);  // *I told you
                    if ($itemMetadata->validate()) {
                        $result = $Tainacan_Item_Metadata->insert($itemMetadata);
                    } else {
                        $this->add_error_log('Error saving value for ' . $itemMetadata->get_metadatum()->get_name() . " in item " . $insertedItem->get_title());
                        $this->add_error_log($itemMetadata->get_errors());
                        continue;
                    }
                } elseif (is_array($itemMetadata)) {
                    if ($updating_item == true) {
                        $this->deleteAllValuesCompoundItemMetadata($insertedItem, $itemMetadata[0][0]->get_metadatum()->get_parent());
                    }
                    foreach ($itemMetadata as $compoundItemMetadata) {
                        $parent_meta_id = null;
                        foreach ($compoundItemMetadata as $itemChildren) {
                            $itemChildren->set_parent_meta_id($parent_meta_id);
                            if ($itemChildren->validate()) {
                                $item_children_metadata = $Tainacan_Item_Metadata->insert($itemChildren);
                                $parent_meta_id = $item_children_metadata->get_parent_meta_id();
                            } else {
                                $this->add_error_log('Error saving value for ' . $itemChildren->get_metadatum()->get_name() . " in item " . $insertedItem->get_title());
                                $this->add_error_log($itemChildren->get_errors());
                                continue;
                            }
                        }
                    }
                }

                //if( $result ){
                //	$values = ( is_array( $itemMetadata->get_value() ) ) ? implode( PHP_EOL, $itemMetadata->get_value() ) : $itemMetadata->get_value();
                //    $this->add_log( 'Item ' . $insertedItem->get_id() .
                //        ' has inserted the values: ' . $values . ' on metadata: ' . $itemMetadata->get_metadatum()->get_name() );
                //} else {
                //    $this->add_error_log( 'Item ' . $insertedItem->get_id() . ' has an error' );
                //}
            }

            if (!$updating_item) {
                $insertedItem->set_status('publish');
            }

            if ($insertedItem->validate()) {
                $insertedItem = $Tainacan_Items->update($insertedItem);
                $this->after_inserted_item($insertedItem, $collection_index);
            } else {
                $this->add_error_log('Error publishing, Item Title: ' . $insertedItem->get_title());
                $this->add_error_log('Error publishing, Item ID: ' . $insertedItem->get_id());
                $this->add_error_log($insertedItem->get_errors());
                return false;
            }
            return $insertedItem;
        } else {
            $this->add_error_log('Collection not set');
            return false;
        }
    }

    private function deleteAllValuesCompoundItemMetadata($item, $compoundMetadataID)
    {
        $Tainacan_Metadata = Metadata::get_instance();
        $Tainacan_Item_Metadata = Item_Metadata::get_instance();
        $compound_metadata = $Tainacan_Metadata->fetch($compoundMetadataID, 'OBJECT');
        $compound_item_metadata = new Entities\Item_Metadata_Entity($item, $compound_metadata);
        $compound_item_metadata_value = $compound_item_metadata->get_value();
        foreach ($compound_item_metadata_value as $item_metadata_value) {
            foreach ($item_metadata_value as $itemMetadata) {
                $Tainacan_Item_Metadata->remove_compound_value($item, $compound_metadata, $itemMetadata->get_parent_meta_id());
            }
        }
    }

    /**
     * @param $value
     * @return bool
     */
    public function is_empty_value($value)
    {
        if (is_array($value)) {
            return (empty(array_filter($value)));
        } else {
            return (trim($value) === '');
        }
    }

    /**
     * @param $metadatum the metadata
     * @param $values the categories names
     *
     * @return array empty with no category or array with IDs
     */
    private function insert_hierarchy($metadatum, $values)
    {

        if (empty($values)) {
            return false;
        }

        $Tainacan_Terms = Terms::get_instance();
        $taxonomy = new Taxonomy($metadatum->get_metadata_type_options()['taxonomy_id']);

        if (strpos($values, '>>') === false) {
            return $values;
        }

        $exploded_values = explode(">>", $values);

        if (empty($exploded_values)) {
            return false;
        }

        if (is_array($exploded_values)) {
            $parent = 0;
            foreach ($exploded_values as $key => $value) {
                $value = trim($value);
                if ($value == '') {
                    $this->add_error_log('Malformed term hierarchy for Item ' . $this->get_current_collection_item() . '. Term skipped. Value: ' . $values);
                    return false;
                }
                $exists = $Tainacan_Terms->term_exists($value, $taxonomy->get_db_identifier(), $parent, true);
                if (false !== $exists && isset($exists->term_id)) {
                    $parent = $exists->term_id;
                } else {
                    $this->add_log('New term created: ' . $value . ' in tax_id: ' . $taxonomy->get_db_identifier() . '; parent: ' . $parent);
                    $term = new Entities\Term();
                    $term->set_name($value);
                    $term->set_parent($parent);
                    $term->set_taxonomy($taxonomy->get_db_identifier());
                    if ($term->validate()) {
                        $term = $Tainacan_Terms->insert($term);
                        $parent = $term->get_id();
                    } else {
                        $this->add_error_log('Invalid Term for Item ' . $this->get_current_collection_item() . ' on Metadatum ' . $metadatum->get_name() . '. Term skipped. Value: ' . $values);
                        $this->add_error_log(implode(',', $term->get_errors()));
                        return false;
                    }

                }
            }
            return $parent !== 0 ? (int)$parent : false;
        } else {
            return false;
        }
    }

    /**
     * @param $collection_id int the collection id
     * @param $mapping array the headers-metadata mapping
     */
    public function save_mapping($collection_id, $mapping)
    {
        update_post_meta($collection_id, 'metadata_mapping', $mapping);
    }

    /**
     * @param $collection_id
     *
     * @return array/bool false if has no mapping or associated array with metadata id and header
     */
    public function get_mapping($collection_id)
    {
        $mapping = get_post_meta($collection_id, 'metadata_mapping', true);
        return ($mapping) ? $mapping : false;
    }


    /**
     * @inheritdoc
     *
     * allow save mapping
     */
    public function add_collection(array $collection)
    {
        if (isset($collection['id'])) {

            if (isset($collection['mapping']) && is_array($collection['mapping'])) {

                foreach ($collection['mapping'] as $metadatum_id => $header) {

                    if (!is_numeric($metadatum_id)) {
                        $repo_key = "create_repository_metadata";
                        $collection_id = $collection['id'];
                        if (strpos($metadatum_id, $repo_key) !== false) {
                            $collection_id = "default";
                        }
                        $metadatum = $this->create_new_metadata($header, $collection_id);

                        if ($metadatum == false) {
                            $this->add_error_log(__("Error while creating metadatum, please review the metadatum description.", 'tainacan'));
                            $this->abort();
                            return false;
                        }

                        if (is_object($metadatum) && $metadatum instanceof Metadatum) {
                            $collection['mapping'][$metadatum->get_id()] = $header;
                        } elseif (is_array($metadatum) && sizeof($metadatum) == 2) {
                            $parent_header = key($header);
                            $collection['mapping'][$metadatum[0]->get_id()] = [$parent_header => $metadatum[1]];
                        }
                        unset($collection['mapping'][$metadatum_id]);
                    }
                }

                $this->save_mapping($collection['id'], $collection['mapping']);

                $coll = Collections::get_instance()->fetch($collection['id']);
                if (empty($coll->get_metadata_order())) {
                    $metadata_order = array_map(
                        function ($meta) {
                            return ["enabled" => true, "id" => $meta];
                        },
                        array_keys($collection['mapping'])
                    );
                    $coll->set_metadata_order($metadata_order);
                }

                if ($coll->validate()) {
                    Collections::get_instance()->update($coll);
                } else {
                    $this->add_error_log(__("Don't save metadata order collection.", 'tainacan'));
                }

            }

            $this->remove_collection($collection['id']);
            $this->collections[] = $collection;
            return true;
        }
    }

    private function get_collections_names()
    {
        $collections_names = [];
        foreach ($this->collections as $col) {
            $collection = Collections::get_instance()->fetch((int)$col['id'], 'OBJECT');
            $collections_names[] = $collection->get_name();
        }
        return $collections_names;
    }

    /**
     * Called when the process is finished. returns the final message to the user with a
     * short description of what happened. May contain HTML code and links
     *
     * @return string
     */
    public function get_output()
    {
        $imported_file = basename($this->get_tmp_file());
        $current_user = wp_get_current_user();
        $author = $current_user->user_login;

        $message = __('imported file:', 'tainacan');
        $message .= " <b> ${imported_file} </b><br/>";
        $message .= __('target collections:', 'tainacan');
        $message .= " <b>" . implode(", ", $this->get_collections_names()) . "</b><br/>";
        $message .= __('Imported by:', 'tainacan');
        $message .= " <b> ${author} </b><br/>";

        return $message;
    }

    public function add_file($file)
    {
        if (parent::add_file($file)) {
            $properties = self::process_document($this->tmp_file);

            $attachment = $properties["attachment"];
            unset($properties["attachment"]);

            $headers = [];

            $processed_properties = self::split_properties_private($properties);

            $processed_properties["Título"] = $properties["Dados técnicos"]["Título"];

            foreach ($processed_properties as $key => $value) {
                $is_private = strpos($key, self::$PRIVATE_SUFFIX) !== false;
                $mapping = array_key_exists($key, self::$METADATUM_MAPPING) ? self::$METADATUM_MAPPING[$key] : null;
                $type = $mapping != null && array_key_exists('type', $mapping) ? $mapping['type'] : "text";

                $headers += self::is_compound_value($value) ? [$key => $this->compound_header($key, $value, !$is_private)] : [$key => "$key|" . $type];
            }

            $headers += ['special_document'];
            $processed_properties['special_document'] = "file:$attachment";

            return $this->document_to_csv($processed_properties, $headers, $this->tmp_file);
        }
        return false;
    }

    private static function split_properties_private($properties): array {
        $processed_properties = [];

        foreach ($properties as $key => $value) {
            if(array_key_exists($key, self::$METADATUM_MAPPING)) {
                $metadata_mapping = self::$METADATUM_MAPPING[$key];

                if(self::is_array_associative($value)) {
                    foreach ($value as $inner_key => $inner_value) {
                        if(array_key_exists($inner_key, $metadata_mapping)) {
                            $processed_key = $metadata_mapping[$inner_key]['private'] ? $key . self::$PRIVATE_SUFFIX : $key;
                            self::append_nested_key($processed_properties, $processed_key, $inner_value, $inner_key);
                        } else {
                            self::append_nested_key($processed_properties, $key, $inner_value, $inner_key);
                        }
                    }
                    continue;
                }

                $processed_key = $metadata_mapping['private'] ? $key . self::$PRIVATE_SUFFIX : $key;
                $processed_properties[$processed_key] = $value;
            } else {
                $processed_properties[$key] = $value;
            }
        }

        return $processed_properties;
    }


    private static function append_nested_key(array &$properties, string $key, $inner_value, $inner_key)
    {
        if(array_key_exists($key, $properties)) {
            $properties[$key] += [$inner_key => $inner_value];
        }
        else {
            $properties[$key] = [$inner_key => $inner_value];
        }
    }

    private static function is_compound_value($value): bool {
        return is_array($value) && (self::has_string_keys($value) || self::has_array_values($value));
    }

    private static function is_array_associative($array): bool {
        return is_array($array) && self::has_string_keys($array);
    }

    private static function has_array_values(array $array): bool {
        return count(array_filter($array, 'is_array')) > 0;
    }

    private static function has_string_keys(array $array): bool {
        return count(array_filter(array_keys($array), 'is_string')) > 0;
    }

    private function sanitize_header($header_name): string
    {
        return str_replace(",", "", $header_name);
    }

    private function remove_private_prefix($header_name): string {
        return str_replace(self::$PRIVATE_SUFFIX, "", $header_name);
    }

    private function compound_header($header_name, array $properties, $display = false): string
    {
        $sanitized_header = $this->sanitize_header($header_name);
        $is_multi_valued = self::has_array_values($properties);

        if($is_multi_valued) {
            $properties = array_reduce($properties, function ($a,$b) {
                return $a + $b;
            }, []);
        }


        $compound = "$sanitized_header|compound(";
        $size = count($properties);
        $count = 0;

        $metadata_header_name = $display ? $header_name : $this->remove_private_prefix($header_name);

        $header_mapping = array_key_exists($metadata_header_name, self::$METADATUM_MAPPING) ? self::$METADATUM_MAPPING[$metadata_header_name] : null;

        foreach ($properties as $key => $value) {
            $sanitized_key = $this->sanitize_header($key);
            $data_type = $header_mapping != null && array_key_exists($key, $header_mapping)
                ?  $header_mapping[$key]['type']
                : "text";

            $compound .= "$sanitized_key|$data_type";
            $count++;

            if ($count < $size) {
                $compound .= ",";
            }
        }
        $compound .= ")";
        $compound .= $is_multi_valued ? '|multiple' : '';
        $compound .= $display ? '|display_yes' : '|display_no';

        return $compound;
    }

    function document_to_csv(array $properties, $headers, $filename = "php://temp", $separator = ",", $max_properties = 20)
    {
        $multi_delimiter = $this->get_option('multivalued_delimiter');
        $fp = fopen($filename, 'w');

        if (!$fp) {
            return false;
        }

        fputcsv($fp, array_slice(array_values($headers), 0, $max_properties), $separator);

        $data = [];

        $count = 0;
        foreach ($properties as $key => $value) {
            if ($count++ >= $max_properties) {
                break;
            }

            $is_compound = self::is_compound_value($value);
            if($is_compound) {
                $data[] = self::has_array_values($value)
                    ? implode($multi_delimiter,
                        array_map(
                            function ($a) use ($separator) {
                                return str_putcsv($a, $separator);
                            }, $value
                        )
                    )
                    : str_putcsv($value, $separator);
            } else {
                $data[] = $value;
            }
        }

        fputcsv($fp, $data, $separator);

        fclose($fp);

        return true;
    }

    private static function read_docx($filename)
    {
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

    private static function process_document($file): array
    {
        $document_data = self::read_docx($file);
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

            if (in_array($trimmed_header, self::$valid_section_headers)) {
                $current_section_header = ucfirst(mb_strtolower($trimmed_header));
            }

            if (!$is_parsing_table && !$is_parsing_locations) {
                if (count($exploded_line) >= 2) {
                    $properties[$current_section_header][$trimmed_header] = trim(implode("", array_slice($exploded_line, 1)));
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

                if (in_array($trimmed_header, self::$ignore_table_headers)) {
                    continue;
                }

                if (in_array($trimmed_header, self::$valid_table_headers)) {
                    $current_table_header = $trimmed_header;
                    $lines_to_skip = 3;
                    continue;
                }

                if (!isset($current_table_header)) {
                    continue;
                }

                if (!isset($properties[$current_section_header][$current_table_header]['menor'])) {
                    $properties[$current_section_header][$current_table_header]['menor'] = trim($line);
                    $lines_to_skip = 1;
                    continue;
                }
        
                if (!isset($properties[$current_section_header][$current_table_header]['maior'])) {
                    $properties[$current_section_header][$current_table_header]['maior'] = trim($line);
                    continue;
                }
            }
        
            if ($is_parsing_locations) {
                if ($trimmed_header == 'Referências Bibliográficas/ Fontes') {
                    $current_section_header = 'Outros';
                    $properties[$current_section_header][$trimmed_header] = trim(implode("", array_slice($exploded_line, 1)));
                    $is_parsing_locations = false;
                    continue;
                }
        
                if (in_array($trimmed_header, self::$ignore_location_headers) || ($location_header_line == 0 && $trimmed_header == '')) {
                    continue;
                }
        
                if ($location_header_line == 0) {
                    array_push($properties[$current_section_header], array(
                        self::$ignore_location_headers[$location_header_line] => $trimmed_header
                    ));
                } else {
                    $properties[$current_section_header][count($properties[$current_section_header]) - 1][self::$ignore_location_headers[$location_header_line]] = $trimmed_header;
                }
                $lines_to_skip = 1;
        
                if ($location_header_line == 3) {
                    $location_header_line = 0;
                    continue;
                }
        
                $location_header_line++;
            }
        }

        $tmp_image_filename = self::create_image_file($document_data["image"]);

        $properties["attachment"] = $tmp_image_filename;

        return $properties;
    }



}
