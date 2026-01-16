<?php

namespace Piyush\ExcelImporter\Readers;

use Piyush\ExcelImporter\Data\Workbook;
use Piyush\ExcelImporter\Data\Worksheet;
use Piyush\ExcelImporter\Data\Row;
use Piyush\ExcelImporter\Data\Cell;
use Piyush\ExcelImporter\Helpers\DateHelper;
use Piyush\ExcelImporter\Helpers\StringHelper;

/**
 * Reader for .xls files (BIFF - Binary Interchange File Format)
 *
 * Supports BIFF5 (Excel 5.0/7.0) and BIFF8 (Excel 97-2003)
 */
class XlsReader implements ReaderInterface
{
    // BIFF Record Types
    const RECORD_BOF = 0x0809;
    const RECORD_EOF = 0x000A;
    const RECORD_BOUNDSHEET = 0x0085;
    const RECORD_SST = 0x00FC;
    const RECORD_LABELSST = 0x00FD;
    const RECORD_NUMBER = 0x0203;
    const RECORD_RK = 0x027E;
    const RECORD_MULRK = 0x00BD;
    const RECORD_LABEL = 0x0204;
    const RECORD_BOOLERR = 0x0205;
    const RECORD_FORMULA = 0x0006;
    const RECORD_STRING = 0x0207;
    const RECORD_BLANK = 0x0201;
    const RECORD_MULBLANK = 0x00BE;
    const RECORD_CONTINUE = 0x003C;
    const RECORD_FORMAT = 0x041E;
    const RECORD_XF = 0x00E0;
    const RECORD_DATEMODE = 0x0022;

    // BIFF Versions
    const BIFF5 = 0x0500;
    const BIFF8 = 0x0600;

    /**
     * @var string File handle
     */
    private $fileHandle;

    /**
     * @var int BIFF version
     */
    private $biffVersion;

    /**
     * @var array Shared strings table
     */
    private $sharedStrings = [];

    /**
     * @var array Sheet information
     */
    private $sheets = [];

    /**
     * @var array Number formats
     */
    private $numberFormats = [];

    /**
     * @var array XF (cell format) records
     */
    private $xfRecords = [];

    /**
     * @var bool Date mode (0 = 1900, 1 = 1904)
     */
    private $dateMode = 0;

    /**
     * @var string File content
     */
    private $data;

    /**
     * @var int Current position in data
     */
    private $pos;

    /**
     * @var int Data length
     */
    private $dataLength;

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

        if ($extension !== 'xls') {
            return false;
        }

        // Check for OLE compound document signature
        $handle = fopen($filePath, 'rb');
        if ($handle === false) {
            return false;
        }

        $signature = fread($handle, 8);
        fclose($handle);

        // OLE2 signature: D0 CF 11 E0 A1 B1 1A E1
        $oleSignature = pack('H*', 'D0CF11E0A1B11AE1');

        return $signature === $oleSignature;
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
        $this->readOleStream($filePath);
        $this->parseGlobalStream();

        return array_column($this->sheets, 'name');
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
        $this->readOleStream($filePath);
        $this->parseGlobalStream();

        $workbook = new Workbook($filePath);

        foreach ($this->sheets as $index => $sheetInfo) {
            if (!empty($sheetNames) && !in_array($sheetInfo['name'], $sheetNames, true)) {
                continue;
            }

            $worksheet = $this->parseSheet($sheetInfo, $index);
            $workbook->addSheet($worksheet);
        }

        return $workbook;
    }

    /**
     * Read the OLE compound document and extract the Workbook stream
     *
     * @param string $filePath The file path
     * @throws \RuntimeException
     */
    private function readOleStream($filePath)
    {
        if (!file_exists($filePath)) {
            throw new \RuntimeException("File not found: {$filePath}");
        }

        $oleData = file_get_contents($filePath);

        if ($oleData === false) {
            throw new \RuntimeException("Cannot read file: {$filePath}");
        }

        // Verify OLE signature
        $signature = substr($oleData, 0, 8);
        $oleSignature = pack('H*', 'D0CF11E0A1B11AE1');

        if ($signature !== $oleSignature) {
            throw new \RuntimeException("Invalid XLS file format");
        }

        // Parse OLE header
        $sectorSize = pow(2, $this->getUInt16($oleData, 30));
        $miniSectorSize = pow(2, $this->getUInt16($oleData, 32));
        $fatSectors = $this->getUInt32($oleData, 44);
        $directorySectorStart = $this->getUInt32($oleData, 48);
        $miniFatSectorStart = $this->getUInt32($oleData, 60);
        $miniFatSectors = $this->getUInt32($oleData, 64);
        $difatSectorStart = $this->getUInt32($oleData, 68);
        $difatSectors = $this->getUInt32($oleData, 72);

        // Read DIFAT (first 109 entries are in header)
        $difat = [];
        for ($i = 0; $i < 109; $i++) {
            $sector = $this->getUInt32($oleData, 76 + $i * 4);
            if ($sector < 0xFFFFFFFE) {
                $difat[] = $sector;
            }
        }

        // Build FAT
        $fat = [];
        foreach ($difat as $fatSector) {
            $offset = 512 + $fatSector * $sectorSize;
            for ($i = 0; $i < $sectorSize / 4; $i++) {
                $fat[] = $this->getUInt32($oleData, $offset + $i * 4);
            }
        }

        // Read directory entries
        $directoryData = $this->readSectorChain($oleData, $directorySectorStart, $fat, $sectorSize);
        $entries = $this->parseDirectoryEntries($directoryData);

        // Find Workbook or Book stream
        $workbookEntry = null;
        foreach ($entries as $entry) {
            $name = strtolower($entry['name']);
            if ($name === 'workbook' || $name === 'book') {
                $workbookEntry = $entry;
                break;
            }
        }

        if ($workbookEntry === null) {
            throw new \RuntimeException("Cannot find Workbook stream in XLS file");
        }

        // Read workbook stream
        if ($workbookEntry['size'] < 4096 && $miniFatSectors > 0) {
            // Use mini stream
            $rootEntry = $entries[0];
            $miniStream = $this->readSectorChain($oleData, $rootEntry['startSector'], $fat, $sectorSize);
            $miniFat = $this->readSectorChain($oleData, $miniFatSectorStart, $fat, $sectorSize);
            $miniFatArray = [];
            for ($i = 0; $i < strlen($miniFat) / 4; $i++) {
                $miniFatArray[] = $this->getUInt32($miniFat, $i * 4);
            }
            $this->data = $this->readMiniSectorChain($miniStream, $workbookEntry['startSector'], $miniFatArray, $miniSectorSize, $workbookEntry['size']);
        } else {
            $this->data = $this->readSectorChain($oleData, $workbookEntry['startSector'], $fat, $sectorSize);
            $this->data = substr($this->data, 0, $workbookEntry['size']);
        }

        $this->pos = 0;
        $this->dataLength = strlen($this->data);

        // Reset parsed data
        $this->sharedStrings = [];
        $this->sheets = [];
        $this->numberFormats = [];
        $this->xfRecords = [];
        $this->dateMode = 0;
    }

    /**
     * Read a chain of sectors
     *
     * @param string $data The OLE data
     * @param int $startSector Starting sector
     * @param array $fat FAT array
     * @param int $sectorSize Sector size
     * @return string
     */
    private function readSectorChain($data, $startSector, $fat, $sectorSize)
    {
        $result = '';
        $sector = $startSector;
        $maxIterations = 100000; // Prevent infinite loops

        while ($sector < 0xFFFFFFFE && $maxIterations-- > 0) {
            $offset = 512 + $sector * $sectorSize;
            if ($offset + $sectorSize <= strlen($data)) {
                $result .= substr($data, $offset, $sectorSize);
            }

            if (!isset($fat[$sector])) {
                break;
            }
            $sector = $fat[$sector];
        }

        return $result;
    }

    /**
     * Read a chain of mini sectors
     *
     * @param string $miniStream The mini stream data
     * @param int $startSector Starting sector
     * @param array $miniFat Mini FAT array
     * @param int $sectorSize Mini sector size
     * @param int $size Total size to read
     * @return string
     */
    private function readMiniSectorChain($miniStream, $startSector, $miniFat, $sectorSize, $size)
    {
        $result = '';
        $sector = $startSector;
        $remaining = $size;
        $maxIterations = 100000;

        while ($sector < 0xFFFFFFFE && $remaining > 0 && $maxIterations-- > 0) {
            $offset = $sector * $sectorSize;
            $toRead = min($sectorSize, $remaining);

            if ($offset + $toRead <= strlen($miniStream)) {
                $result .= substr($miniStream, $offset, $toRead);
                $remaining -= $toRead;
            }

            if (!isset($miniFat[$sector])) {
                break;
            }
            $sector = $miniFat[$sector];
        }

        return $result;
    }

    /**
     * Parse directory entries
     *
     * @param string $data Directory data
     * @return array
     */
    private function parseDirectoryEntries($data)
    {
        $entries = [];
        $entrySize = 128;
        $count = strlen($data) / $entrySize;

        for ($i = 0; $i < $count; $i++) {
            $offset = $i * $entrySize;
            $nameLen = $this->getUInt16($data, $offset + 64);

            if ($nameLen === 0) {
                continue;
            }

            $name = '';
            for ($j = 0; $j < ($nameLen - 2) / 2; $j++) {
                $char = $this->getUInt16($data, $offset + $j * 2);
                if ($char > 0) {
                    $name .= chr($char);
                }
            }

            $entries[] = [
                'name' => $name,
                'type' => ord($data[$offset + 66]),
                'startSector' => $this->getUInt32($data, $offset + 116),
                'size' => $this->getUInt32($data, $offset + 120)
            ];
        }

        return $entries;
    }

    /**
     * Parse the global workbook stream
     */
    private function parseGlobalStream()
    {
        $this->pos = 0;

        while ($this->pos < $this->dataLength) {
            $record = $this->readRecord();

            if ($record === null) {
                break;
            }

            switch ($record['type']) {
                case self::RECORD_BOF:
                    $this->biffVersion = $this->getUInt16($record['data'], 0);
                    break;

                case self::RECORD_DATEMODE:
                    $this->dateMode = $this->getUInt16($record['data'], 0);
                    break;

                case self::RECORD_BOUNDSHEET:
                    $this->parseBoundSheet($record['data']);
                    break;

                case self::RECORD_SST:
                    $this->parseSST($record['data']);
                    break;

                case self::RECORD_FORMAT:
                    $this->parseFormat($record['data']);
                    break;

                case self::RECORD_XF:
                    $this->parseXF($record['data']);
                    break;

                case self::RECORD_EOF:
                    break 2;
            }
        }
    }

    /**
     * Read a BIFF record
     *
     * @return array|null
     */
    private function readRecord()
    {
        if ($this->pos + 4 > $this->dataLength) {
            return null;
        }

        $type = $this->getUInt16($this->data, $this->pos);
        $length = $this->getUInt16($this->data, $this->pos + 2);
        $this->pos += 4;

        if ($this->pos + $length > $this->dataLength) {
            return null;
        }

        $data = substr($this->data, $this->pos, $length);
        $this->pos += $length;

        // Handle CONTINUE records
        while ($this->pos + 4 <= $this->dataLength) {
            $nextType = $this->getUInt16($this->data, $this->pos);
            if ($nextType !== self::RECORD_CONTINUE) {
                break;
            }

            $nextLength = $this->getUInt16($this->data, $this->pos + 2);
            $this->pos += 4;

            if ($this->pos + $nextLength > $this->dataLength) {
                break;
            }

            $data .= substr($this->data, $this->pos, $nextLength);
            $this->pos += $nextLength;
        }

        return [
            'type' => $type,
            'data' => $data
        ];
    }

    /**
     * Parse BOUNDSHEET record
     *
     * @param string $data Record data
     */
    private function parseBoundSheet($data)
    {
        $offset = $this->getUInt32($data, 0);
        $hidden = ord($data[4]);
        $type = ord($data[5]);

        if ($this->biffVersion === self::BIFF8) {
            $nameLen = ord($data[6]);
            $optionFlags = ord($data[7]);

            if ($optionFlags & 0x01) {
                // UTF-16LE
                $name = $this->decodeUtf16($data, 8, $nameLen);
            } else {
                // ASCII
                $name = substr($data, 8, $nameLen);
            }
        } else {
            $nameLen = ord($data[6]);
            $name = substr($data, 7, $nameLen);
        }

        $this->sheets[] = [
            'name' => $name,
            'offset' => $offset,
            'hidden' => $hidden,
            'type' => $type
        ];
    }

    /**
     * Parse SST (Shared String Table) record
     *
     * @param string $data Record data
     */
    private function parseSST($data)
    {
        $totalStrings = $this->getUInt32($data, 0);
        $uniqueStrings = $this->getUInt32($data, 4);
        $pos = 8;
        $dataLen = strlen($data);

        for ($i = 0; $i < $uniqueStrings && $pos < $dataLen; $i++) {
            $result = $this->readString($data, $pos);
            $this->sharedStrings[] = $result['string'];
            $pos = $result['pos'];
        }
    }

    /**
     * Read a string from BIFF data
     *
     * @param string $data The data
     * @param int $pos Starting position
     * @return array ['string' => string, 'pos' => int]
     */
    private function readString($data, $pos)
    {
        $dataLen = strlen($data);

        if ($pos + 3 > $dataLen) {
            return ['string' => '', 'pos' => $pos];
        }

        $charCount = $this->getUInt16($data, $pos);
        $optionFlags = ord($data[$pos + 2]);
        $pos += 3;

        $isUnicode = ($optionFlags & 0x01) !== 0;
        $hasExtString = ($optionFlags & 0x04) !== 0;
        $hasRichText = ($optionFlags & 0x08) !== 0;

        $richTextRuns = 0;
        $extStringLen = 0;

        if ($hasRichText && $pos + 2 <= $dataLen) {
            $richTextRuns = $this->getUInt16($data, $pos);
            $pos += 2;
        }

        if ($hasExtString && $pos + 4 <= $dataLen) {
            $extStringLen = $this->getUInt32($data, $pos);
            $pos += 4;
        }

        $string = '';
        if ($isUnicode) {
            $byteLen = $charCount * 2;
            if ($pos + $byteLen <= $dataLen) {
                $string = $this->decodeUtf16($data, $pos, $charCount);
                $pos += $byteLen;
            }
        } else {
            if ($pos + $charCount <= $dataLen) {
                $string = substr($data, $pos, $charCount);
                $pos += $charCount;
            }
        }

        // Skip rich text formatting runs
        $pos += $richTextRuns * 4;

        // Skip extended string data
        $pos += $extStringLen;

        return ['string' => $string, 'pos' => $pos];
    }

    /**
     * Parse FORMAT record
     *
     * @param string $data Record data
     */
    private function parseFormat($data)
    {
        $formatIndex = $this->getUInt16($data, 0);

        if ($this->biffVersion === self::BIFF8) {
            $strLen = $this->getUInt16($data, 2);
            $optionFlags = ord($data[4]);

            if ($optionFlags & 0x01) {
                $formatString = $this->decodeUtf16($data, 5, $strLen);
            } else {
                $formatString = substr($data, 5, $strLen);
            }
        } else {
            $strLen = ord($data[2]);
            $formatString = substr($data, 3, $strLen);
        }

        $this->numberFormats[$formatIndex] = $formatString;
    }

    /**
     * Parse XF record
     *
     * @param string $data Record data
     */
    private function parseXF($data)
    {
        $formatIndex = $this->getUInt16($data, 2);
        $this->xfRecords[] = $formatIndex;
    }

    /**
     * Parse a worksheet
     *
     * @param array $sheetInfo Sheet info
     * @param int $index Sheet index
     * @return Worksheet
     */
    private function parseSheet($sheetInfo, $index)
    {
        $worksheet = new Worksheet($sheetInfo['name'], $index);
        $this->pos = $sheetInfo['offset'];

        $inSheet = false;
        $lastFormulaRow = 0;
        $lastFormulaCol = 0;

        while ($this->pos < $this->dataLength) {
            $record = $this->readRecord();

            if ($record === null) {
                break;
            }

            switch ($record['type']) {
                case self::RECORD_BOF:
                    $inSheet = true;
                    break;

                case self::RECORD_EOF:
                    if ($inSheet) {
                        return $worksheet;
                    }
                    break;

                case self::RECORD_NUMBER:
                    $this->parseNumberCell($record['data'], $worksheet);
                    break;

                case self::RECORD_RK:
                    $this->parseRKCell($record['data'], $worksheet);
                    break;

                case self::RECORD_MULRK:
                    $this->parseMulRKCell($record['data'], $worksheet);
                    break;

                case self::RECORD_LABELSST:
                    $this->parseLabelSSTCell($record['data'], $worksheet);
                    break;

                case self::RECORD_LABEL:
                    $this->parseLabelCell($record['data'], $worksheet);
                    break;

                case self::RECORD_BOOLERR:
                    $this->parseBoolErrCell($record['data'], $worksheet);
                    break;

                case self::RECORD_FORMULA:
                    $result = $this->parseFormulaCell($record['data'], $worksheet);
                    $lastFormulaRow = $result['row'];
                    $lastFormulaCol = $result['col'];
                    break;

                case self::RECORD_STRING:
                    // String result of formula
                    $this->parseStringResult($record['data'], $worksheet, $lastFormulaRow, $lastFormulaCol);
                    break;

                case self::RECORD_BLANK:
                    // Blank cell - ignore
                    break;

                case self::RECORD_MULBLANK:
                    // Multiple blank cells - ignore
                    break;
            }
        }

        return $worksheet;
    }

    /**
     * Parse NUMBER cell
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     */
    private function parseNumberCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $col = $this->getUInt16($data, 2);
        $xfIndex = $this->getUInt16($data, 4);
        $value = $this->getDouble($data, 6);

        $cell = $this->createNumericCell($value, $xfIndex, $row, $col);
        $this->addCellToWorksheet($worksheet, $cell, $row, $col);
    }

    /**
     * Parse RK cell (compressed number)
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     */
    private function parseRKCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $col = $this->getUInt16($data, 2);
        $xfIndex = $this->getUInt16($data, 4);
        $rk = $this->getUInt32($data, 6);
        $value = $this->decodeRK($rk);

        $cell = $this->createNumericCell($value, $xfIndex, $row, $col);
        $this->addCellToWorksheet($worksheet, $cell, $row, $col);
    }

    /**
     * Parse MULRK cell (multiple RK values)
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     */
    private function parseMulRKCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $colFirst = $this->getUInt16($data, 2);
        $colLast = $this->getUInt16($data, strlen($data) - 2);

        $pos = 4;
        for ($col = $colFirst; $col <= $colLast; $col++) {
            $xfIndex = $this->getUInt16($data, $pos);
            $rk = $this->getUInt32($data, $pos + 2);
            $value = $this->decodeRK($rk);

            $cell = $this->createNumericCell($value, $xfIndex, $row, $col);
            $this->addCellToWorksheet($worksheet, $cell, $row, $col);

            $pos += 6;
        }
    }

    /**
     * Parse LABELSST cell (shared string)
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     */
    private function parseLabelSSTCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $col = $this->getUInt16($data, 2);
        $sstIndex = $this->getUInt32($data, 6);

        $value = isset($this->sharedStrings[$sstIndex]) ? $this->sharedStrings[$sstIndex] : '';
        $coord = StringHelper::buildCoordinate($col, $row);

        $cell = new Cell($value, Cell::TYPE_STRING, $coord);
        $this->addCellToWorksheet($worksheet, $cell, $row, $col);
    }

    /**
     * Parse LABEL cell (inline string)
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     */
    private function parseLabelCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $col = $this->getUInt16($data, 2);

        if ($this->biffVersion === self::BIFF8) {
            $result = $this->readString($data, 6);
            $value = $result['string'];
        } else {
            $strLen = $this->getUInt16($data, 6);
            $value = substr($data, 8, $strLen);
        }

        $coord = StringHelper::buildCoordinate($col, $row);
        $cell = new Cell($value, Cell::TYPE_STRING, $coord);
        $this->addCellToWorksheet($worksheet, $cell, $row, $col);
    }

    /**
     * Parse BOOLERR cell
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     */
    private function parseBoolErrCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $col = $this->getUInt16($data, 2);
        $value = ord($data[6]);
        $isError = ord($data[7]) === 1;

        $coord = StringHelper::buildCoordinate($col, $row);

        if ($isError) {
            $cell = new Cell('#ERROR', Cell::TYPE_STRING, $coord);
        } else {
            $cell = new Cell($value === 1, Cell::TYPE_BOOLEAN, $coord);
        }

        $this->addCellToWorksheet($worksheet, $cell, $row, $col);
    }

    /**
     * Parse FORMULA cell
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     * @return array
     */
    private function parseFormulaCell($data, $worksheet)
    {
        $row = $this->getUInt16($data, 0) + 1;
        $col = $this->getUInt16($data, 2);
        $xfIndex = $this->getUInt16($data, 4);

        $coord = StringHelper::buildCoordinate($col, $row);

        // Check result type
        $byte6 = ord($data[12]);
        $byte7 = ord($data[13]);

        if ($byte6 === 0xFF && $byte7 === 0xFF) {
            // Special value
            $type = ord($data[6]);
            switch ($type) {
                case 0: // String - will be in following STRING record
                    return ['row' => $row, 'col' => $col];
                case 1: // Boolean
                    $value = ord($data[8]) === 1;
                    $cell = new Cell($value, Cell::TYPE_BOOLEAN, $coord);
                    break;
                case 2: // Error
                    $cell = new Cell('#ERROR', Cell::TYPE_STRING, $coord);
                    break;
                case 3: // Empty
                    $cell = new Cell(null, Cell::TYPE_EMPTY, $coord);
                    break;
                default:
                    $cell = new Cell(null, Cell::TYPE_EMPTY, $coord);
            }
        } else {
            // Numeric value
            $value = $this->getDouble($data, 6);
            $cell = $this->createNumericCell($value, $xfIndex, $row, $col);
        }

        $this->addCellToWorksheet($worksheet, $cell, $row, $col);

        return ['row' => $row, 'col' => $col];
    }

    /**
     * Parse STRING result (formula string result)
     *
     * @param string $data Record data
     * @param Worksheet $worksheet
     * @param int $row Row index
     * @param int $col Column index
     */
    private function parseStringResult($data, $worksheet, $row, $col)
    {
        if ($this->biffVersion === self::BIFF8) {
            $result = $this->readString($data, 0);
            $value = $result['string'];
        } else {
            $strLen = $this->getUInt16($data, 0);
            $value = substr($data, 2, $strLen);
        }

        $coord = StringHelper::buildCoordinate($col, $row);
        $cell = new Cell($value, Cell::TYPE_STRING, $coord);
        $this->addCellToWorksheet($worksheet, $cell, $row, $col);
    }

    /**
     * Create a numeric cell with date detection
     *
     * @param float $value The numeric value
     * @param int $xfIndex The XF index
     * @param int $row Row index
     * @param int $col Column index
     * @return Cell
     */
    private function createNumericCell($value, $xfIndex, $row, $col)
    {
        $coord = StringHelper::buildCoordinate($col, $row);

        // Check if this is a date format
        if ($this->isDateFormat($xfIndex)) {
            $dateTime = DateHelper::excelToDateTime($value);
            if ($dateTime !== null) {
                return new Cell($dateTime, Cell::TYPE_DATE, $coord, $dateTime->format('Y-m-d H:i:s'));
            }
        }

        // Regular number
        if (floor($value) == $value && abs($value) < PHP_INT_MAX) {
            $value = (int) $value;
        }

        return new Cell($value, Cell::TYPE_NUMBER, $coord);
    }

    /**
     * Check if XF index indicates a date format
     *
     * @param int $xfIndex The XF index
     * @return bool
     */
    private function isDateFormat($xfIndex)
    {
        if (!isset($this->xfRecords[$xfIndex])) {
            return false;
        }

        $formatIndex = $this->xfRecords[$xfIndex];

        // Check built-in date formats
        if (DateHelper::isDateFormatCode($formatIndex)) {
            return true;
        }

        // Check custom format
        if (isset($this->numberFormats[$formatIndex])) {
            return DateHelper::isDateFormatString($this->numberFormats[$formatIndex]);
        }

        return false;
    }

    /**
     * Add a cell to the worksheet
     *
     * @param Worksheet $worksheet
     * @param Cell $cell
     * @param int $rowIndex
     * @param int $colIndex
     */
    private function addCellToWorksheet($worksheet, $cell, $rowIndex, $colIndex)
    {
        $row = $worksheet->getRow($rowIndex);

        if ($row === null) {
            $row = new Row($rowIndex);
            $worksheet->addRow($row, $rowIndex);
        }

        $row->addCell($cell, $colIndex);
    }

    /**
     * Decode RK compressed number
     *
     * @param int $rk The RK value
     * @return float
     */
    private function decodeRK($rk)
    {
        $isInteger = ($rk & 0x02) !== 0;
        $isDivided = ($rk & 0x01) !== 0;

        if ($isInteger) {
            $value = ($rk >> 2);
            // Sign extend if negative
            if ($value >= 0x20000000) {
                $value -= 0x40000000;
            }
        } else {
            // IEEE 754 with top 30 bits
            $packed = pack('V', 0) . pack('V', ($rk & 0xFFFFFFFC));
            $value = unpack('d', $packed)[1];
        }

        if ($isDivided) {
            $value /= 100;
        }

        return $value;
    }

    /**
     * Decode UTF-16LE string
     *
     * @param string $data The data
     * @param int $offset Starting offset
     * @param int $charCount Number of characters
     * @return string
     */
    private function decodeUtf16($data, $offset, $charCount)
    {
        $str = substr($data, $offset, $charCount * 2);
        return mb_convert_encoding($str, 'UTF-8', 'UTF-16LE');
    }

    /**
     * Get unsigned 16-bit integer
     *
     * @param string $data The data
     * @param int $offset The offset
     * @return int
     */
    private function getUInt16($data, $offset)
    {
        return unpack('v', substr($data, $offset, 2))[1];
    }

    /**
     * Get unsigned 32-bit integer
     *
     * @param string $data The data
     * @param int $offset The offset
     * @return int
     */
    private function getUInt32($data, $offset)
    {
        return unpack('V', substr($data, $offset, 4))[1];
    }

    /**
     * Get IEEE 754 double
     *
     * @param string $data The data
     * @param int $offset The offset
     * @return float
     */
    private function getDouble($data, $offset)
    {
        return unpack('d', substr($data, $offset, 8))[1];
    }
}
