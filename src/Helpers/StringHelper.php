<?php

namespace Piyush\ExcelImporter\Helpers;

/**
 * Helper class for string operations in Excel files
 */
class StringHelper
{
    /**
     * @var array Shared strings lookup table
     */
    private $sharedStrings = [];

    /**
     * @var int Count of shared strings
     */
    private $count = 0;

    /**
     * Add a shared string to the table
     *
     * @param string $string The string to add
     * @return int The index of the string
     */
    public function addString($string)
    {
        $this->sharedStrings[$this->count] = $string;
        return $this->count++;
    }

    /**
     * Get a shared string by index
     *
     * @param int $index The string index
     * @return string|null
     */
    public function getString($index)
    {
        return isset($this->sharedStrings[$index]) ? $this->sharedStrings[$index] : null;
    }

    /**
     * Get all shared strings
     *
     * @return array
     */
    public function getStrings()
    {
        return $this->sharedStrings;
    }

    /**
     * Get the count of shared strings
     *
     * @return int
     */
    public function getCount()
    {
        return $this->count;
    }

    /**
     * Clear all shared strings
     *
     * @return void
     */
    public function clear()
    {
        $this->sharedStrings = [];
        $this->count = 0;
    }

    /**
     * Parse cell coordinate to get column and row
     *
     * @param string $coordinate The cell coordinate (e.g., A1, AB123)
     * @return array ['column' => 'A', 'row' => 1, 'columnIndex' => 0]
     */
    public static function parseCoordinate($coordinate)
    {
        preg_match('/^([A-Z]+)(\d+)$/i', strtoupper($coordinate), $matches);

        if (count($matches) !== 3) {
            return [
                'column' => '',
                'row' => 0,
                'columnIndex' => -1
            ];
        }

        return [
            'column' => $matches[1],
            'row' => (int) $matches[2],
            'columnIndex' => self::columnToIndex($matches[1])
        ];
    }

    /**
     * Convert column letter to index (A=0, B=1, ..., Z=25, AA=26)
     *
     * @param string $column The column letter
     * @return int
     */
    public static function columnToIndex($column)
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
     * Convert column index to letter (0=A, 1=B, ..., 25=Z, 26=AA)
     *
     * @param int $index The column index (0-based)
     * @return string
     */
    public static function indexToColumn($index)
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

    /**
     * Build coordinate from column and row
     *
     * @param int $columnIndex Column index (0-based)
     * @param int $rowIndex Row index (1-based)
     * @return string
     */
    public static function buildCoordinate($columnIndex, $rowIndex)
    {
        return self::indexToColumn($columnIndex) . $rowIndex;
    }

    /**
     * Decode XML entities and special characters
     *
     * @param string $string The string to decode
     * @return string
     */
    public static function decodeXmlString($string)
    {
        if (empty($string)) {
            return '';
        }

        // Decode HTML entities
        $string = html_entity_decode($string, ENT_QUOTES | ENT_XML1, 'UTF-8');

        // Replace _x????_ encoded characters (Excel XML encoding)
        $string = preg_replace_callback(
            '/_x([0-9A-Fa-f]{4})_/',
            function ($matches) {
                return mb_convert_encoding(
                    pack('H*', $matches[1]),
                    'UTF-8',
                    'UTF-16BE'
                );
            },
            $string
        );

        return $string;
    }

    /**
     * Clean a string value (trim and normalize whitespace)
     *
     * @param string $string The string to clean
     * @return string
     */
    public static function cleanString($string)
    {
        if (empty($string)) {
            return '';
        }

        // Normalize line endings
        $string = str_replace(["\r\n", "\r"], "\n", $string);

        // Trim whitespace
        $string = trim($string);

        return $string;
    }

    /**
     * Check if a string looks like a number
     *
     * @param string $string The string to check
     * @return bool
     */
    public static function isNumeric($string)
    {
        if (empty($string) && $string !== '0') {
            return false;
        }

        return is_numeric($string);
    }

    /**
     * Check if a string represents a boolean value
     *
     * @param string $string The string to check
     * @return bool
     */
    public static function isBoolean($string)
    {
        $lower = strtolower(trim($string));
        return in_array($lower, ['true', 'false', '1', '0'], true);
    }

    /**
     * Parse a boolean string to boolean value
     *
     * @param string $string The string to parse
     * @return bool
     */
    public static function parseBoolean($string)
    {
        $lower = strtolower(trim($string));
        return in_array($lower, ['true', '1'], true);
    }
}
