<?php

namespace Builder\Traits;

/**
 * Trait InitialisationStateTrait
 */
trait InitialisationStateTrait
{
    /**
     * @var bool
     */
    private $initialised = false;

    /**
     * Mark the Builder as having been initialised.
     *
     * @return void
     */
    private function setAsInitialised()
    {
        $this->initialised = true;
    }

    /**
     * Returns true if the Builder has been initialised.
     *
     * @return bool
     */
    private function isInitialised()
    {
        return $this->initialised === true;
    }

    /**
     * Returns true if the Builder has not been initialised.
     *
     * @return bool
     */
    private function isNotInitialised()
    {
        return $this->initialised === false;
    }
}
