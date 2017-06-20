<?php

namespace Builder\Interfaces;

/**
 * Interface BuilderTestInterface
 * @package Builder\Interfaces
 */
interface BuilderTestInterface
{
    public function builder_is_correct_builder();

    public function can_create_single_sheet_spreadsheet();

    public function can_create_multi_sheet_spreadsheet();
}
