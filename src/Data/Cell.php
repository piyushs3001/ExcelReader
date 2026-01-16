<?php

namespace Piyush\ExcelImporter\Data;

/**
 * Represents a single cell in a worksheet
 */
class Cell
{
    /**
     * Cell type constants
     */
    const TYPE_STRING = 'string';
    const TYPE_NUMBER = 'number';
    const TYPE_DATE = 'date';
    const TYPE_BOOLEAN = 'boolean';
    const TYPE_FORMULA = 'formula';
    const TYPE_EMPTY = 'empty';

    /**
     * @var mixed The raw value of the cell
     */
    private $value;

    /**
     * @var string The data type of the cell
     */
    private $type;

    /**
     * @var string The cell coordinate (e.g., A1, B2)
     */
    private $coordinate;

    /**
     * @var string|null The formatted value as string
     */
    private $formattedValue;

    /**
     * Create a new Cell instance
     *
     * @param mixed $value The cell value
     * @param string $type The cell type
     * @param string $coordinate The cell coordinate
     * @param string|null $formattedValue Optional formatted value
     */
    public function __construct($value, $type = self::TYPE_STRING, $coordinate = '', $formattedValue = null)
    {
        $this->value = $value;
        $this->type = $type;
        $this->coordinate = $coordinate;
        $this->formattedValue = $formattedValue;
    }

    /**
     * Get the raw value of the cell
     *
     * @return mixed
     */
    public function getValue()
    {
        return $this->value;
    }

    /**
     * Get the data type of the cell
     *
     * @return string
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * Get the cell coordinate
     *
     * @return string
     */
    public function getCoordinate()
    {
        return $this->coordinate;
    }

    /**
     * Get the formatted value as string
     *
     * @return string
     */
    public function getFormattedValue()
    {
        if ($this->formattedValue !== null) {
            return $this->formattedValue;
        }

        if ($this->value === null) {
            return '';
        }

        if ($this->type === self::TYPE_BOOLEAN) {
            return $this->value ? 'TRUE' : 'FALSE';
        }

        if ($this->type === self::TYPE_DATE && $this->value instanceof \DateTime) {
            return $this->value->format('Y-m-d H:i:s');
        }

        return (string) $this->value;
    }

    /**
     * Check if cell is empty
     *
     * @return bool
     */
    public function isEmpty()
    {
        return $this->type === self::TYPE_EMPTY || $this->value === null || $this->value === '';
    }

    /**
     * Check if cell contains a date
     *
     * @return bool
     */
    public function isDate()
    {
        return $this->type === self::TYPE_DATE;
    }

    /**
     * Check if cell contains a number
     *
     * @return bool
     */
    public function isNumber()
    {
        return $this->type === self::TYPE_NUMBER;
    }

    /**
     * Check if cell contains a string
     *
     * @return bool
     */
    public function isString()
    {
        return $this->type === self::TYPE_STRING;
    }

    /**
     * Check if cell contains a boolean
     *
     * @return bool
     */
    public function isBoolean()
    {
        return $this->type === self::TYPE_BOOLEAN;
    }
}
