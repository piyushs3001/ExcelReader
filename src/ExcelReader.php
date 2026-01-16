<?php

namespace Piyush\ExcelImporter;

use Piyush\ExcelImporter\Data\Workbook;
use Piyush\ExcelImporter\Readers\ReaderInterface;
use Piyush\ExcelImporter\Readers\XlsxReader;
use Piyush\ExcelImporter\Readers\XlsReader;

/**
 * Main entry point for reading Excel files
 *
 * Usage:
 *   $excel = ExcelReader::load('/path/to/file.xlsx');
 *   $sheet = $excel->getSheet(0);
 *   foreach ($sheet->getRows() as $row) {
 *       foreach ($row->getCells() as $cell) {
 *           echo $cell->getValue();
 *       }
 *   }
 */
class ExcelReader
{
    /**
     * @var ReaderInterface[] Available readers
     */
    private static $readers = [];

    /**
     * Load an Excel file and return a Workbook
     *
     * @param string $filePath The path to the Excel file
     * @return Workbook
     * @throws \RuntimeException If the file cannot be read
     * @throws \InvalidArgumentException If no suitable reader is found
     */
    public static function load($filePath)
    {
        $reader = self::getReaderForFile($filePath);
        return $reader->load($filePath);
    }

    /**
     * Load only specific sheets from an Excel file
     *
     * @param string $filePath The path to the Excel file
     * @param array $sheetNames Names of sheets to load
     * @return Workbook
     * @throws \RuntimeException If the file cannot be read
     * @throws \InvalidArgumentException If no suitable reader is found
     */
    public static function loadSheets($filePath, array $sheetNames)
    {
        $reader = self::getReaderForFile($filePath);
        return $reader->loadSheets($filePath, $sheetNames);
    }

    /**
     * Get the list of sheet names from an Excel file without loading data
     *
     * @param string $filePath The path to the Excel file
     * @return string[]
     * @throws \RuntimeException If the file cannot be read
     * @throws \InvalidArgumentException If no suitable reader is found
     */
    public static function getSheetNames($filePath)
    {
        $reader = self::getReaderForFile($filePath);
        return $reader->getSheetNames($filePath);
    }

    /**
     * Check if a file can be read by any available reader
     *
     * @param string $filePath The path to the file
     * @return bool
     */
    public static function canRead($filePath)
    {
        try {
            self::getReaderForFile($filePath);
            return true;
        } catch (\InvalidArgumentException $e) {
            return false;
        }
    }

    /**
     * Get the appropriate reader for a file
     *
     * @param string $filePath The path to the file
     * @return ReaderInterface
     * @throws \InvalidArgumentException If no suitable reader is found
     */
    public static function getReaderForFile($filePath)
    {
        if (!file_exists($filePath)) {
            throw new \InvalidArgumentException("File not found: {$filePath}");
        }

        // Initialize readers if not done
        if (empty(self::$readers)) {
            self::$readers = [
                new XlsxReader(),
                new XlsReader()
            ];
        }

        // Find a suitable reader
        foreach (self::$readers as $reader) {
            if ($reader->canRead($filePath)) {
                return $reader;
            }
        }

        $extension = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));

        // Provide helpful error for .xlsx when zip is not available
        if ($extension === 'xlsx' && !XlsxReader::isZipExtensionAvailable()) {
            throw new \RuntimeException(
                "Cannot read .xlsx file: The PHP zip extension is not enabled. " .
                "Please enable the zip extension or convert your file to .xls format."
            );
        }

        throw new \InvalidArgumentException(
            "No suitable reader found for file: {$filePath} (extension: {$extension})"
        );
    }

    /**
     * Create an XLSX reader explicitly
     *
     * @return XlsxReader
     */
    public static function createXlsxReader()
    {
        return new XlsxReader();
    }

    /**
     * Create an XLS reader explicitly
     *
     * @return XlsReader
     */
    public static function createXlsReader()
    {
        return new XlsReader();
    }

    /**
     * Register a custom reader
     *
     * @param ReaderInterface $reader The reader to register
     * @return void
     */
    public static function registerReader(ReaderInterface $reader)
    {
        // Initialize readers if not done
        if (empty(self::$readers)) {
            self::$readers = [
                new XlsxReader(),
                new XlsReader()
            ];
        }

        // Add to beginning so custom readers take precedence
        array_unshift(self::$readers, $reader);
    }

    /**
     * Clear all registered readers (useful for testing)
     *
     * @return void
     */
    public static function clearReaders()
    {
        self::$readers = [];
    }
}
