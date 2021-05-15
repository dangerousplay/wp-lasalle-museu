<?php

namespace TainacanLasalle;

class Docx_importer extends \Tainacan\Importer\CSV {
    public function __construct($attributes = array()) {
        parent::__construct($attributes);
    }

    public function process_item($index, $collection_id)
    {
        // TODO: Implement process_item() method.
    }

    function options_form() {
        $form = '<div class="field">';
        $form .= '<label class="label">' . __('My Importer Option 1', 'tainacan') . '</label>';
        $form .= '<div class="control">';
        $form .= '<input type="text" class="input" name="my_importer_option_1" value="2" />';
        $form .= '</div>';
        $form .= '</div>';

        return $form;
    }
}