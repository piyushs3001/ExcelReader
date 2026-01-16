<?php

namespace Piyush\ExcelImporter\Readers;

use Piyush\ExcelImporter\Data\Workbook;
use Piyush\ExcelImporter\Data\Worksheet;
use Piyush\ExcelImporter\Data\Row;
use Piyush\ExcelImporter\Data\Cell;
use Piyush\ExcelImporter\Helpers\DateHelper;
use Piyush\ExcelImporter\Helpers\StringHelper;

/**
 * Reader for .xlsx files (Office Open XML format)
 *
 * .xlsx files are ZIP archives containing XML files:
 * - xl/workbook.xml - Sheet names and structure
 * - xl/sharedStrings.xml - All text strings (shared pool)
 * - xl/worksheets/sheet1.xml - Cell data for each sheet
 * - xl/styles.xml - Number formats (for date detection)
 *
 * Note: Requires PHP zip extension to be enabled.
 */
class XlsxReader implements ReaderInterface
{
    /**
     * Check if the zip extension is available
     *
     * @return bool
     */
    public static function isZipExtensionAvailable()
    {
        return extension_loaded('zip') && class_exists('ZipArchive');
    }

    /**
     * @var \ZipArchive The ZIP archive
     */
    private $zip;

    /**
     * @var StringHelper Shared strings helper
     */
    private $sharedStrings;

    /**
     * @var array Number format styles for date detection
     */
    private $numberFormats = [];

    /**
     * @var array Cell format indices
     */
    private $cellFormats = [];

    /**
     * @var array Sheet info from workbook.xml
     */
    private $sheetInfo = [];

    /**
     * @var array Relationships from workbook.xml.rels
     */
    private $relationships = [];

    /**
     * Check if the reader can read the given file
     *
     * @param string $filePath The file path to check
     * @return bool
     */
    public function canRead($filePath)
    {
        if (!file_exists($filePath)) {
            return false;
        }

        $extension = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));

        if ($extension !== 'xlsx') {
            return false;
        }

        // Check if zip extension is available
        if (!self::isZipExtensionAvailable()) {
            return false;
        }

        // Verify it's a valid ZIP file
        $zip = new \ZipArchive();
        $result = $zip->open($filePath);

        if ($result !== true) {
            return false;
        }

        // Check for required files
        $hasWorkbook = $zip->locateName('xl/workbook.xml') !== false;
        $zip->close();

        return $hasWorkbook;
    }

    /**
     * Load and parse an Excel file
     *
     * @param string $filePath The file path to load
     * @return Workbook
     * @throws \RuntimeException If the file cannot be read
     */
    public function load($filePath)
    {
        return $this->loadSheets($filePath, []);
    }

    /**
     * Get the list of sheet names without loading full content
     *
     * @param string $filePath The file path
     * @return string[]
     * @throws \RuntimeException If the file cannot be read
     */
    public function getSheetNames($filePath)
    {
        $this->openZip($filePath);
        $this->parseWorkbook();
        $this->closeZip();

        return array_column($this->sheetInfo, 'name');
    }

    /**
     * Load only specific sheets from the file
     *
     * @param string $filePath The file path to load
     * @param array $sheetNames Names of sheets to load (empty = all)
     * @return Workbook
     * @throws \RuntimeException If the file cannot be read
     */
    public function loadSheets($filePath, array $sheetNames)
    {
        $this->openZip($filePath);

        // Parse workbook structure
        $this->parseRelationships();
        $this->parseWorkbook();
        $this->parseSharedStrings();
        $this->parseStyles();

        // Create workbook
        $workbook = new Workbook($filePath);

        // Load each sheet
        foreach ($this->sheetInfo as $index => $info) {
            // Skip if not in requested sheets
            if (!empty($sheetNames) && !in_array($info['name'], $sheetNames, true)) {
                continue;
            }

            $worksheet = $this->parseWorksheet($info, $index);
            $workbook->addSheet($worksheet);
        }

        $this->closeZip();

        return $workbook;
    }

    /**
     * Open the ZIP archive
     *
     * @param string $filePath The file path
     * @throws \RuntimeException
     */
    private function openZip($filePath)
    {
        if (!file_exists($filePath)) {
            throw new \RuntimeException("File not found: {$filePath}");
        }

        if (!self::isZipExtensionAvailable()) {
            throw new \RuntimeException(
                "The PHP zip extension is required to read .xlsx files. " .
                "Please enable the zip extension or use .xls format instead."
            );
        }

        $this->zip = new \ZipArchive();
        $result = $this->zip->open($filePath);

        if ($result !== true) {
            throw new \RuntimeException("Failed to open ZIP file: {$filePath} (Error code: {$result})");
        }

        $this->sharedStrings = new StringHelper();
        $this->numberFormats = [];
        $this->cellFormats = [];
        $this->sheetInfo = [];
        $this->relationships = [];
    }

    /**
     * Close the ZIP archive
     */
    private function closeZip()
    {
        if ($this->zip !== null) {
            $this->zip->close();
            $this->zip = null;
        }
    }

    /**
     * Parse relationships from xl/_rels/workbook.xml.rels
     */
    private function parseRelationships()
    {
        $content = $this->getZipContent('xl/_rels/workbook.xml.rels');

        if ($content === false) {
            return;
        }

        $xml = $this->parseXml($content);

        if ($xml === false) {
            return;
        }

        foreach ($xml->Relationship as $rel) {
            $id = (string) $rel['Id'];
            $target = (string) $rel['Target'];
            $type = (string) $rel['Type'];

            $this->relationships[$id] = [
                'target' => $target,
                'type' => $type
            ];
        }
    }

    /**
     * Parse workbook.xml to get sheet information
     */
    private function parseWorkbook()
    {
        $content = $this->getZipContent('xl/workbook.xml');

        if ($content === false) {
            throw new \RuntimeException('Cannot read workbook.xml');
        }

        $xml = $this->parseXml($content);

        if ($xml === false) {
            throw new \RuntimeException('Cannot parse workbook.xml');
        }

        // Get sheets
        $sheets = $xml->sheets->sheet;

        if ($sheets === null) {
            return;
        }

        $index = 0;
        foreach ($sheets as $sheet) {
            $name = (string) $sheet['name'];
            $sheetId = (string) $sheet['sheetId'];

            // Get relationship ID
            $rId = null;
            foreach ($sheet->attributes('r', true) as $attr => $value) {
                if ($attr === 'id') {
                    $rId = (string) $value;
                    break;
                }
            }

            // Determine target file
            $target = null;
            if ($rId !== null && isset($this->relationships[$rId])) {
                $target = 'xl/' . $this->relationships[$rId]['target'];
            } else {
                // Fallback to default naming
                $target = 'xl/worksheets/sheet' . ($index + 1) . '.xml';
            }

            $this->sheetInfo[$index] = [
                'name' => $name,
                'sheetId' => $sheetId,
                'rId' => $rId,
                'target' => $target
            ];

            $index++;
        }
    }

    /**
     * Parse sharedStrings.xml
     */
    private function parseSharedStrings()
    {
        $content = $this->getZipContent('xl/sharedStrings.xml');

        if ($content === false) {
            return; // File may not exist if no strings in workbook
        }

        $xml = $this->parseXml($content);

        if ($xml === false) {
            return;
        }

        foreach ($xml->si as $si) {
            $text = $this->extractStringValue($si);
            $this->sharedStrings->addString($text);
        }
    }

    /**
     * Extract string value from si element (handles rich text)
     *
     * @param \SimpleXMLElement $si The si element
     * @return string
     */
    private function extractStringValue($si)
    {
        // Simple text
        if (isset($si->t)) {
            return StringHelper::decodeXmlString((string) $si->t);
        }

        // Rich text (multiple r elements)
        if (isset($si->r)) {
            $text = '';
            foreach ($si->r as $r) {
                if (isset($r->t)) {
                    $text .= (string) $r->t;
                }
            }
            return StringHelper::decodeXmlString($text);
        }

        return '';
    }

    /**
     * Parse styles.xml for number formats
     */
    private function parseStyles()
    {
        $content = $this->getZipContent('xl/styles.xml');

        if ($content === false) {
            return;
        }

        $xml = $this->parseXml($content);

        if ($xml === false) {
            return;
        }

        // Built-in number formats for dates
        $builtInFormats = [
            14 => 'mm-dd-yy',
            15 => 'd-mmm-yy',
            16 => 'd-mmm',
            17 => 'mmm-yy',
            18 => 'h:mm AM/PM',
            19 => 'h:mm:ss AM/PM',
            20 => 'h:mm',
            21 => 'h:mm:ss',
            22 => 'm/d/yy h:mm',
            45 => 'mm:ss',
            46 => '[h]:mm:ss',
            47 => 'mmss.0',
        ];

        $this->numberFormats = $builtInFormats;

        // Parse custom number formats
        if (isset($xml->numFmts)) {
            foreach ($xml->numFmts->numFmt as $numFmt) {
                $id = (int) $numFmt['numFmtId'];
                $code = (string) $numFmt['formatCode'];
                $this->numberFormats[$id] = $code;
            }
        }

        // Parse cell formats (xf elements)
        if (isset($xml->cellXfs)) {
            $index = 0;
            foreach ($xml->cellXfs->xf as $xf) {
                $numFmtId = (int) $xf['numFmtId'];
                $this->cellFormats[$index] = $numFmtId;
                $index++;
            }
        }
    }

    /**
     * Parse a worksheet XML file
     *
     * @param array $info Sheet info
     * @param int $index Sheet index
     * @return Worksheet
     */
    private function parseWorksheet($info, $index)
    {
        $worksheet = new Worksheet($info['name'], $index);

        $content = $this->getZipContent($info['target']);

        if ($content === false) {
            return $worksheet;
        }

        $xml = $this->parseXml($content);

        if ($xml === false) {
            return $worksheet;
        }

        // Parse sheet data
        if (!isset($xml->sheetData)) {
            return $worksheet;
        }

        foreach ($xml->sheetData->row as $rowXml) {
            $rowIndex = (int) $rowXml['r'];
            $row = new Row($rowIndex);

            foreach ($rowXml->c as $cellXml) {
                $cell = $this->parseCell($cellXml);
                $coord = StringHelper::parseCoordinate($cell->getCoordinate());
                $row->addCell($cell, $coord['columnIndex']);
            }

            $worksheet->addRow($row, $rowIndex);
        }

        return $worksheet;
    }

    /**
     * Parse a cell element
     *
     * @param \SimpleXMLElement $cellXml The cell XML element
     * @return Cell
     */
    private function parseCell($cellXml)
    {
        $coordinate = (string) $cellXml['r'];
        $type = (string) $cellXml['t'];
        $styleIndex = isset($cellXml['s']) ? (int) $cellXml['s'] : 0;

        $value = null;
        $cellType = Cell::TYPE_EMPTY;
        $formattedValue = null;

        // Get the value
        $rawValue = isset($cellXml->v) ? (string) $cellXml->v : null;

        // Handle inline strings
        if ($type === 'inlineStr') {
            if (isset($cellXml->is)) {
                $value = $this->extractStringValue($cellXml->is);
            }
            $cellType = Cell::TYPE_STRING;
        }
        // Handle shared strings
        elseif ($type === 's') {
            $stringIndex = (int) $rawValue;
            $value = $this->sharedStrings->getString($stringIndex);
            $cellType = Cell::TYPE_STRING;
        }
        // Handle booleans
        elseif ($type === 'b') {
            $value = $rawValue === '1';
            $cellType = Cell::TYPE_BOOLEAN;
        }
        // Handle errors
        elseif ($type === 'e') {
            $value = $rawValue;
            $cellType = Cell::TYPE_STRING;
        }
        // Handle strings (direct text)
        elseif ($type === 'str') {
            $value = StringHelper::decodeXmlString($rawValue);
            $cellType = Cell::TYPE_STRING;
        }
        // Handle numbers and dates
        elseif ($rawValue !== null && $rawValue !== '') {
            $numericValue = (float) $rawValue;

            // Check if this is a date format
            if ($this->isDateFormat($styleIndex)) {
                $dateTime = DateHelper::excelToDateTime($numericValue);
                if ($dateTime !== null) {
                    $value = $dateTime;
                    $cellType = Cell::TYPE_DATE;
                    $formattedValue = $dateTime->format('Y-m-d H:i:s');
                } else {
                    $value = $numericValue;
                    $cellType = Cell::TYPE_NUMBER;
                }
            } else {
                // Regular number
                $value = $numericValue;
                $cellType = Cell::TYPE_NUMBER;

                // Convert to int if it's a whole number
                if (floor($numericValue) == $numericValue && abs($numericValue) < PHP_INT_MAX) {
                    $value = (int) $numericValue;
                }
            }
        }

        // Check for formula
        if (isset($cellXml->f)) {
            // We store the calculated value, not the formula
            if ($cellType === Cell::TYPE_EMPTY && $value !== null) {
                $cellType = Cell::TYPE_FORMULA;
            }
        }

        return new Cell($value, $cellType, $coordinate, $formattedValue);
    }

    /**
     * Check if a style index indicates a date format
     *
     * @param int $styleIndex The style index
     * @return bool
     */
    private function isDateFormat($styleIndex)
    {
        if (!isset($this->cellFormats[$styleIndex])) {
            return false;
        }

        $numFmtId = $this->cellFormats[$styleIndex];

        // Check built-in date format codes
        if (DateHelper::isDateFormatCode($numFmtId)) {
            return true;
        }

        // Check custom format string
        if (isset($this->numberFormats[$numFmtId])) {
            return DateHelper::isDateFormatString($this->numberFormats[$numFmtId]);
        }

        return false;
    }

    /**
     * Get content from ZIP archive
     *
     * @param string $name The file name within the archive
     * @return string|false
     */
    private function getZipContent($name)
    {
        if ($this->zip === null) {
            return false;
        }

        // Try exact name first
        $content = $this->zip->getFromName($name);

        if ($content !== false) {
            return $content;
        }

        // Try with leading slash
        return $this->zip->getFromName('/' . $name);
    }

    /**
     * Parse XML content with namespace handling
     *
     * @param string $content The XML content
     * @return \SimpleXMLElement|false
     */
    private function parseXml($content)
    {
        // Suppress errors for invalid XML
        $previousValue = libxml_use_internal_errors(true);

        // Remove namespace prefixes for easier parsing
        $content = preg_replace('/xmlns[^=]*="[^"]*"/i', '', $content);
        $content = preg_replace('/<([a-zA-Z0-9]+):/', '<', $content);
        $content = preg_replace('/<\/([a-zA-Z0-9]+):/', '</', $content);

        $xml = simplexml_load_string($content);

        libxml_clear_errors();
        libxml_use_internal_errors($previousValue);

        return $xml;
    }
}
