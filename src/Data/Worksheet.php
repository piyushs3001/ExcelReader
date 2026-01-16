<?php

namespace Piyush\ExcelImporter\Data;

/**
 * Represents a single worksheet in a workbook
 */
class Worksheet implements \IteratorAggregate, \Countable
{
    /**
     * @var string The worksheet name
     */
    private $name;

    /**
     * @var int The worksheet index (0-based)
     */
    private $index;

    /**
     * @var Row[] Array of rows in this worksheet
     */
    private $rows = [];

    /**
     * @var int The highest row number
     */
    private $highestRow = 0;

    /**
     * @var int The highest column index
     */
    private $highestColumn = 0;

    /**
     * Create a new Worksheet instance
     *
     * @param string $name The worksheet name
     * @param int $index The worksheet index (0-based)
     */
    public function __construct($name, $index = 0)
    {
        $this->name = $name;
        $this->index = $index;
    }

    /**
     * Get the worksheet name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Get the worksheet index
     *
     * @return int
     */
    public function getIndex()
    {
        return $this->index;
    }

    /**
     * Add a row to the worksheet
     *
     * @param Row $row The row to add
     * @param int $rowIndex The row index (1-based)
     * @return self
     */
    public function addRow(Row $row, $rowIndex)
    {
        $this->rows[$rowIndex] = $row;

        if ($rowIndex > $this->highestRow) {
            $this->highestRow = $rowIndex;
        }

        foreach ($row->getCells() as $colIndex => $cell) {
            if ($colIndex > $this->highestColumn) {
                $this->highestColumn = $colIndex;
            }
        }

        return $this;
    }

    /**
     * Get a row by index
     *
     * @param int $rowIndex The row index (1-based)
     * @return Row|null
     */
    public function getRow($rowIndex)
    {
        return isset($this->rows[$rowIndex]) ? $this->rows[$rowIndex] : null;
    }

    /**
     * Get all rows
     *
     * @return Row[]
     */
    public function getRows()
    {
        ksort($this->rows);
        return $this->rows;
    }

    /**
     * Get a specific cell
     *
     * @param string $coordinate Cell coordinate (e.g., A1, B2)
     * @return Cell|null
     */
    public function getCell($coordinate)
    {
        preg_match('/^([A-Z]+)(\d+)$/i', $coordinate, $matches);

        if (count($matches) !== 3) {
            return null;
        }

        $column = strtoupper($matches[1]);
        $rowIndex = (int) $matches[2];

        $row = $this->getRow($rowIndex);

        if ($row === null) {
            return null;
        }

        return $row->getCellByColumn($column);
    }

    /**
     * Get the highest row number
     *
     * @return int
     */
    public function getHighestRow()
    {
        return $this->highestRow;
    }

    /**
     * Get the highest column index
     *
     * @return int
     */
    public function getHighestColumn()
    {
        return $this->highestColumn;
    }

    /**
     * Get the highest column letter
     *
     * @return string
     */
    public function getHighestColumnLetter()
    {
        return Row::columnIndexToLetter($this->highestColumn);
    }

    /**
     * Get the number of rows
     *
     * @return int
     */
    public function count()
    {
        return count($this->rows);
    }

    /**
     * Get iterator for rows
     *
     * @return \ArrayIterator
     */
    public function getIterator()
    {
        ksort($this->rows);
        return new \ArrayIterator($this->rows);
    }

    /**
     * Convert worksheet to 2D array
     *
     * @param bool $includeEmptyRows Whether to include empty rows
     * @return array
     */
    public function toArray($includeEmptyRows = false)
    {
        $result = [];
        ksort($this->rows);

        if ($includeEmptyRows && $this->highestRow > 0) {
            for ($i = 1; $i <= $this->highestRow; $i++) {
                $row = $this->getRow($i);
                $result[$i] = $row !== null ? $row->toArray() : [];
            }
        } else {
            foreach ($this->rows as $rowIndex => $row) {
                if (!$row->isEmpty()) {
                    $result[$rowIndex] = $row->toArray();
                }
            }
        }

        return $result;
    }

    /**
     * Get rows as associative arrays using first row as headers
     *
     * @return array
     */
    public function toAssociativeArray()
    {
        $rows = $this->toArray();

        if (empty($rows)) {
            return [];
        }

        // Get headers from first row
        $headers = array_shift($rows);

        if (empty($headers)) {
            return [];
        }

        $result = [];

        foreach ($rows as $rowIndex => $row) {
            $assocRow = [];
            foreach ($headers as $colIndex => $header) {
                $key = $header !== null ? (string) $header : 'column_' . $colIndex;
                $assocRow[$key] = isset($row[$colIndex]) ? $row[$colIndex] : null;
            }
            $result[] = $assocRow;
        }

        return $result;
    }
}
