<?php
/*
Plugin Name: Museu LaSalle
Plugin URI: tainacan.org
Description: Plugin for Lasalle Museu customizations
Author: Media Lab / LaSalle
Version: 0.0.1
Text Domain: tainacan-lasalle
License: GPLv2 or later
License URI: http://www.gnu.org/licenses/gpl-3.0.html
*/
namespace TainacanLasalle;

class Plugin {

    public function __construct() {
        add_action("tainacan-register-importers", [$this, "register_importer"]);
    }

    function register_importer() {
        global $Tainacan_Importer_Handler;

        require_once( plugin_dir_path(__FILE__) . 'document_importer.php' );

        $Tainacan_Importer_Handler->register_importer([
            'name' => 'Documento Word',
            'description' => __('Import items from a DOCX document', 'tainacan'),
            'slug' => 'docx_importer',
            'class_name' => '\TainacanLasalle\Docx_importer',
            'manual_collection' => true,
            'manual_mapping' => false,
        ]);
    }
}

$TainacanLasalle = new \TainacanLasalle\Plugin();