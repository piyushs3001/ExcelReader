<?php

namespace Piyush\ExcelImporter\Helpers;

/**
 * Helper class for Excel date conversions
 */
class DateHelper
{
    /**
     * Unix timestamp of Excel base date (1900-01-01)
     * Excel incorrectly considers 1900 as a leap year, so we use 1899-12-30
     */
    const EXCEL_BASE_DATE = -2209161600;

    /**
     * Number of seconds in a day
     */
    const SECONDS_PER_DAY = 86400;

    /**
     * Excel date format codes that indicate a date
     */
    private static $dateFormatCodes = [
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58
    ];

    /**
     * Date format patterns to detect date formats
     */
    private static $dateFormatPatterns = [
        '/y{1,4}/i',      // Year
        '/m{1,5}/i',      // Month (but not minutes if preceded by h or followed by s)
        '/d{1,4}/i',      // Day
        '/h{1,2}/i',      // Hour
        '/s{1,2}/i',      // Second
    ];

    /**
     * Convert Excel serial date to PHP DateTime
     *
     * @param float|int $excelDate The Excel serial date number
     * @param string|null $timezone Optional timezone
     * @return \DateTime|null
     */
    public static function excelToDateTime($excelDate, $timezone = null)
    {
        if (!is_numeric($excelDate)) {
            return null;
        }

        // Excel dates before 1900-01-01 are invalid
        if ($excelDate < 1) {
            return null;
        }

        // Adjust for Excel's incorrect leap year bug (1900 was not a leap year)
        if ($excelDate < 60) {
            // Dates before March 1, 1900
            $excelDate += 1;
        }

        // Calculate days and time fraction
        $days = floor($excelDate);
        $timeFraction = $excelDate - $days;

        // Excel base date is December 30, 1899
        $baseTimestamp = self::EXCEL_BASE_DATE;

        // Add days to base timestamp
        $timestamp = $baseTimestamp + ($days * self::SECONDS_PER_DAY);

        // Add time portion
        $timestamp += round($timeFraction * self::SECONDS_PER_DAY);

        try {
            // Suppress any errors during DateTime creation
            $previousHandler = set_error_handler(function() { return true; });
            $previousLevel = error_reporting(0);

            // Create DateTime from timestamp using @ notation
            // This creates DateTime in UTC without requiring timezone database
            $dateTime = date_create('@' . (int) $timestamp);

            error_reporting($previousLevel);
            restore_error_handler();

            if ($dateTime === false) {
                return null;
            }

            if ($timezone !== null) {
                try {
                    $tz = @timezone_open($timezone);
                    if ($tz !== false) {
                        date_timezone_set($dateTime, $tz);
                    }
                } catch (\Exception $e) {
                    // Ignore timezone errors
                }
            }

            return $dateTime;
        } catch (\Exception $e) {
            return null;
        }
    }

    /**
     * Convert Excel serial date to Unix timestamp
     *
     * @param float|int $excelDate The Excel serial date number
     * @return int|null
     */
    public static function excelToTimestamp($excelDate)
    {
        $dateTime = self::excelToDateTime($excelDate);
        return $dateTime !== null ? $dateTime->getTimestamp() : null;
    }

    /**
     * Convert Excel serial date to formatted string
     *
     * @param float|int $excelDate The Excel serial date number
     * @param string $format PHP date format string
     * @return string|null
     */
    public static function excelToFormatted($excelDate, $format = 'Y-m-d H:i:s')
    {
        $dateTime = self::excelToDateTime($excelDate);
        return $dateTime !== null ? $dateTime->format($format) : null;
    }

    /**
     * Check if a format code indicates a date format
     *
     * @param int $formatCode The Excel number format code
     * @return bool
     */
    public static function isDateFormatCode($formatCode)
    {
        return in_array((int) $formatCode, self::$dateFormatCodes, true);
    }

    /**
     * Check if a format string indicates a date format
     *
     * @param string $formatString The Excel number format string
     * @return bool
     */
    public static function isDateFormatString($formatString)
    {
        if (empty($formatString)) {
            return false;
        }

        // Remove escaped characters and quoted strings
        $format = preg_replace('/\[[^\]]*\]/', '', $formatString);
        $format = preg_replace('/"[^"]*"/', '', $format);
        $format = preg_replace("/\\\\.'/", '', $format);

        // Check for date/time components
        // Exclude formats that are just numbers or currency
        if (preg_match('/^[#0.,\$%\s\-\(\)]+$/', $format)) {
            return false;
        }

        // Check for date patterns
        foreach (self::$dateFormatPatterns as $pattern) {
            if (preg_match($pattern, $format)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Check if an Excel value looks like a date (reasonable date range)
     *
     * @param float|int $value The numeric value
     * @return bool
     */
    public static function isValidExcelDate($value)
    {
        if (!is_numeric($value)) {
            return false;
        }

        // Excel dates start from 1 (January 1, 1900)
        // Upper limit around year 9999
        return $value >= 1 && $value <= 2958465;
    }

    /**
     * Extract just the date portion from an Excel date
     *
     * @param float|int $excelDate The Excel serial date number
     * @return int
     */
    public static function getDatePart($excelDate)
    {
        return (int) floor($excelDate);
    }

    /**
     * Extract just the time portion from an Excel date (as fraction of day)
     *
     * @param float|int $excelDate The Excel serial date number
     * @return float
     */
    public static function getTimePart($excelDate)
    {
        return $excelDate - floor($excelDate);
    }

    /**
     * Check if Excel date has a time component
     *
     * @param float|int $excelDate The Excel serial date number
     * @return bool
     */
    public static function hasTime($excelDate)
    {
        return self::getTimePart($excelDate) > 0;
    }
}
