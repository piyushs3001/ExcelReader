<?php

namespace Piyush\ExcelImporter\Data;

/**
 * Represents a row of cells in a worksheet
 */
class Row implements \IteratorAggregate, \Countable
{
    /**
     * @var Cell[] Array of cells in this row
     */
    private $cells = [];

    /**
     * @var int The row index (1-based)
     */
    private $rowIndex;

    /**
     * Create a new Row instance
     *
     * @param int $rowIndex The row index (1-based)
     */
    public function __construct($rowIndex)
    {
        $this->rowIndex = $rowIndex;
    }

    /**
     * Add a cell to the row
     *
     * @param Cell $cell The cell to add
     * @param int $columnIndex The column index (0-based)
     * @return self
     */
    public function addCell(Cell $cell, $columnIndex)
    {
        $this->cells[$columnIndex] = $cell;
        return $this;
    }

    /**
     * Get a cell by column index
     *
     * @param int $columnIndex The column index (0-based)
     * @return Cell|null
     */
    public function getCell($columnIndex)
    {
        return isset($this->cells[$columnIndex]) ? $this->cells[$columnIndex] : null;
    }

    /**
     * Get a cell by column letter
     *
     * @param string $column The column letter (A, B, C, etc.)
     * @return Cell|null
     */
    public function getCellByColumn($column)
    {
        $columnIndex = self::columnLetterToIndex($column);
        return $this->getCell($columnIndex);
    }

    /**
     * Get all cells in the row
     *
     * @return Cell[]
     */
    public function getCells()
    {
        ksort($this->cells);
        return $this->cells;
    }

    /**
     * Get the row index (1-based)
     *
     * @return int
     */
    public function getRowIndex()
    {
        return $this->rowIndex;
    }

    /**
     * Get the number of cells in the row
     *
     * @return int
     */
    public function count()
    {
        return count($this->cells);
    }

    /**
     * Get iterator for cells
     *
     * @return \ArrayIterator
     */
    public function getIterator()
    {
        ksort($this->cells);
        return new \ArrayIterator($this->cells);
    }

    /**
     * Convert row to array of values
     *
     * @return array
     */
    public function toArray()
    {
        $result = [];
        ksort($this->cells);

        if (empty($this->cells)) {
            return $result;
        }

        $maxColumn = max(array_keys($this->cells));

        for ($i = 0; $i <= $maxColumn; $i++) {
            $result[$i] = isset($this->cells[$i]) ? $this->cells[$i]->getValue() : null;
        }

        return $result;
    }

    /**
     * Check if row is empty
     *
     * @return bool
     */
    public function isEmpty()
    {
        foreach ($this->cells as $cell) {
            if (!$cell->isEmpty()) {
                return false;
            }
        }
        return true;
    }

    /**
     * Convert column letter to index (A=0, B=1, ..., Z=25, AA=26, etc.)
     *
     * @param string $column The column letter
     * @return int
     */
    public static function columnLetterToIndex($column)
    {
        $column = strtoupper($column);
        $length = strlen($column);
        $index = 0;

        for ($i = 0; $i < $length; $i++) {
            $index = $index * 26 + (ord($column[$i]) - ord('A') + 1);
        }

        return $index - 1;
    }

    /**
     * Convert column index to letter (0=A, 1=B, ..., 25=Z, 26=AA, etc.)
     *
     * @param int $index The column index (0-based)
     * @return string
     */
    public static function columnIndexToLetter($index)
    {
        $letter = '';
        $index++;

        while ($index > 0) {
            $index--;
            $letter = chr(ord('A') + ($index % 26)) . $letter;
            $index = intval($index / 26);
        }

        return $letter;
    }
}
