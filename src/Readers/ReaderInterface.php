<?php

namespace Piyush\ExcelImporter\Readers;

use Piyush\ExcelImporter\Data\Workbook;

/**
 * Interface for Excel file readers
 */
interface ReaderInterface
{
    /**
     * Check if the reader can read the given file
     *
     * @param string $filePath The file path to check
     * @return bool
     */
    public function canRead($filePath);

    /**
     * Load and parse an Excel file
     *
     * @param string $filePath The file path to load
     * @return Workbook
     * @throws \RuntimeException If the file cannot be read
     */
    public function load($filePath);

    /**
     * Get the list of sheet names without loading full content
     *
     * @param string $filePath The file path
     * @return string[]
     * @throws \RuntimeException If the file cannot be read
     */
    public function getSheetNames($filePath);

    /**
     * Load only specific sheets from the file
     *
     * @param string $filePath The file path to load
     * @param array $sheetNames Names of sheets to load
     * @return Workbook
     * @throws \RuntimeException If the file cannot be read
     */
    public function loadSheets($filePath, array $sheetNames);
}
