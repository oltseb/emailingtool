<?php

/**
 * Class Helper
 */
class Helper
{

    /**
     * @param array $arguments
     * @param int $index
     * @return mixed|string
     */
    public function getReadFileName(array $arguments, int $index)
    {
        return isset($arguments[$index]) ? $arguments[$index] : 'defaul.xlsx';
    }

    /**
     * @param array $arguments
     * @param int $index
     * @return mixed|string
     */
    public function getEmailColumnNumber(array $arguments, int $index)
    {
        return isset($arguments[$index]) ? $arguments[$index] : '1';
    }

    /**
     * @param array $arguments
     * @param int $index
     * @return mixed|string
     */
    public function getWriteFileName(array $arguments, int $index)
    {
        return isset($arguments[$index]) ? $arguments[$index] : 'defaulOutput.xlsx';
    }
}