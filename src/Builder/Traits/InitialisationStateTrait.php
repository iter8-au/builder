<?php

namespace Iter8\Builder\Traits;

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
        return true === $this->initialised;
    }

    /**
     * Returns true if the Builder has not been initialised.
     *
     * @return bool
     */
    private function isNotInitialised()
    {
        return false === $this->initialised;
    }
}
