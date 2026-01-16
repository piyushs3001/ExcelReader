<?php

namespace Piyush\ExcelImporter\Data;

/**
 * Represents an Excel workbook
 */
class Workbook implements \IteratorAggregate, \Countable
{
    /**
     * @var Worksheet[] Array of worksheets
     */
    private $sheets = [];

    /**
     * @var array Sheet names indexed by position
     */
    private $sheetNames = [];

    /**
     * @var string The file path of the workbook
     */
    private $filePath;

    /**
     * Create a new Workbook instance
     *
     * @param string $filePath The file path
     */
    public function __construct($filePath = '')
    {
        $this->filePath = $filePath;
    }

    /**
     * Add a worksheet to the workbook
     *
     * @param Worksheet $sheet The worksheet to add
     * @return self
     */
    public function addSheet(Worksheet $sheet)
    {
        $index = $sheet->getIndex();
        $this->sheets[$index] = $sheet;
        $this->sheetNames[$index] = $sheet->getName();

        return $this;
    }

    /**
     * Get a worksheet by name or index
     *
     * @param string|int $nameOrIndex The sheet name or index
     * @return Worksheet|null
     */
    public function getSheet($nameOrIndex)
    {
        // If integer, get by index
        if (is_int($nameOrIndex)) {
            return isset($this->sheets[$nameOrIndex]) ? $this->sheets[$nameOrIndex] : null;
        }

        // If string, search by name
        foreach ($this->sheets as $sheet) {
            if (strcasecmp($sheet->getName(), $nameOrIndex) === 0) {
                return $sheet;
            }
        }

        return null;
    }

    /**
     * Get the first worksheet
     *
     * @return Worksheet|null
     */
    public function getFirstSheet()
    {
        return $this->getSheet(0);
    }

    /**
     * Get all worksheets
     *
     * @return Worksheet[]
     */
    public function getSheets()
    {
        ksort($this->sheets);
        return $this->sheets;
    }

    /**
     * Get all sheet names
     *
     * @return string[]
     */
    public function getSheetNames()
    {
        ksort($this->sheetNames);
        return array_values($this->sheetNames);
    }

    /**
     * Get the number of sheets
     *
     * @return int
     */
    public function getSheetCount()
    {
        return count($this->sheets);
    }

    /**
     * Get the file path
     *
     * @return string
     */
    public function getFilePath()
    {
        return $this->filePath;
    }

    /**
     * Check if a sheet exists
     *
     * @param string|int $nameOrIndex The sheet name or index
     * @return bool
     */
    public function hasSheet($nameOrIndex)
    {
        return $this->getSheet($nameOrIndex) !== null;
    }

    /**
     * Get the number of sheets (Countable interface)
     *
     * @return int
     */
    public function count()
    {
        return $this->getSheetCount();
    }

    /**
     * Get iterator for sheets (IteratorAggregate interface)
     *
     * @return \ArrayIterator
     */
    public function getIterator()
    {
        ksort($this->sheets);
        return new \ArrayIterator($this->sheets);
    }
}
