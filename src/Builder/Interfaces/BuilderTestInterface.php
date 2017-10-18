<?php

namespace Builder\Interfaces;

/**
 * Interface BuilderTestInterface
 */
interface BuilderTestInterface
{
    /**
     * This test should be used to verify that the builder returned from the Service Provider
     * is the builder you specified in app options during bootstrapping.
     */
    public function builder_is_correct_builder();

    /**
     * This test should be used to verify that a single sheet spreadsheet can be generated.
     */
    public function can_create_single_sheet_spreadsheet();

    /**
     * This test should be used to verify that a multi-sheet spreadsheet can be generated.
     */
    public function can_create_multi_sheet_spreadsheet();
}
