<?php

namespace Office;

use Exception;

class GoogleSpreadSheetsTaker
{
    private string $strUrl; //Исходная ссылка таблицы
    private string $strValidUrl; //Обработанная ссылка таблицы
    private string $strGid; //ID листа таблицы
    private array $arParams; //ID дополнительных параметров TODO: развивать!
    private const arParamsMap = [ //TODO: Тоже развивать!
        "range",
    ];

    public function __construct( string $strUrl, array $arParams = [] )
    {
        $this->strUrl = $strUrl;

        try {
            $this->strValidUrl = static::validateUrl($strUrl);
            $this->strGid = static::getGidFromUrl($strUrl);
            if ( !empty($arParams) ) {
                $this->arParams = static::validateParams($arParams);
            }
        } catch (Exception $e) {
            die($e->getMessage());
        }
    }

    /**
     * @throws Exception
     */
    protected static function validateUrl (string $strUrl ) : string
    {
        if ( !$strUrl ) {
            throw new Exception("Не передана ссылка " . $strUrl);
        }

        if( strpos($strUrl, '?') ) {
            $strUrl = substr($strUrl, 0, strpos($strUrl, '?'));
        }

        switch ( true ) {
            case (strpos($strUrl, 'edit')) :
                $strUrl = substr($strUrl, 0, strpos($strUrl, 'edit'));
                break;
            case (strpos($strUrl, 'export')) :
                $strUrl = substr($strUrl, 0, strpos($strUrl, 'export'));
                break;
        }
        return $strUrl;
    }

    /**
     * @throws Exception
     */
    protected static function getGidFromUrl ($strUrl ) : string
    {
        if ( strpos($strUrl, '#') === false ) {
            throw new Exception("Не передан id листа (gid)");
        }
        return substr($strUrl, strpos($strUrl, '#')+5);
    }

    /**
     * @throws Exception
     */
    public function getDataFromRange (string $strRange ) : array
    {
        $strCsvTableData = self::getCsvData(self::getRequestUrl($strRange));

        if ( !$strCsvTableData ) {
            throw new Exception("Не удается найти данные в диапазоне " . $strRange);
        }

        return self::parseCsvToArray($strCsvTableData);
    }

    public function getData() : array
    {
        try {
            return self::getDataFromRange($this->arParams["range"]);
        } catch (Exception $e) {
            die($e->getMessage());
        }
    }

    /**
     * @throws Exception
     */
    protected function validateParams ( array $arParams ) : array
    {
        $arResultParams = [];

        foreach ( $arParams as $strParamName => $mixParamValue ) {
            if ( !in_array($strParamName, self::arParamsMap, true) ) {
                throw new Exception("Не найден параметр " . $strParamName);
            }
            $arResultParams[$strParamName] = $mixParamValue;
        }

        return $arResultParams;
    }

    /**
     * @throws Exception
     */
    public function getFirstRowFromRange (string $strRange ) : array
    {
        if ( empty($strRange) ) {
            throw new Exception("Не определен параметр range");
        }

        if ( strpos($strRange, ":") === false ) {
            throw new Exception("Неверный формат диапазона " . $strRange . ". Не удается найти ':'.");
        }

        $arRange = explode(":", $strRange);

        if ( count($arRange) != 2 ) {
            throw new Exception("Неверный формат диапазона " . $strRange . ". Количество координат не равно 2.");
        }

        $strFirstCoordinateLetters = preg_replace('/[^A-Z]/', '', $arRange[0]);
        $strFirstCoordinateNumbers = preg_replace('/[^0-9]/', '', $arRange[0]);
        $strSecondCoordinateNumbers = preg_replace('/[^0-9]/', '', $arRange[1]);

        $strNewRange = $strFirstCoordinateLetters . $strFirstCoordinateNumbers . ":" . $strFirstCoordinateLetters . $strSecondCoordinateNumbers;

        return self::getDataFromRange( $strNewRange );
    }

    /**
     * @throws Exception
     */
    public function getFirstRow()
    {
        if ( !$this->arParams["range"] ) {
            throw new Exception("Не определен параметр range");
        }

        try {
            return self::getFirstRowFromRange($this->arParams["range"]);
        } catch (Exception $e) {
            die($e->getMessage());
        }
    }

    /**
     * @throws Exception
     */
    protected function getRequestUrl (string $strRange = "" ) : string
    {
        if ( empty($strRange) && !$this->arParams["range"] ) {
            throw new Exception("Не определен параметр range");
        }

        $strUrlRange = $strRange ?: $this->arParams["params"];

        return $this->strValidUrl . 'export?format=csv&gid=' . $this->strGid . '&range=' . $strUrlRange;
    }

    protected function parseCsvToArray( string $strCsvTableData ) : array
    {
        $arTableRows = explode("\r\n", $strCsvTableData);

        $arTableData = [];

        foreach ( $arTableRows as $strRow ) {
            $arTableData[] = explode(',', $strRow);
        }

        return $arTableData;
    }

    protected function getCsvData( string $strRequestUrl ) : string
    {
        try {
            return file_get_contents($strRequestUrl);
        } catch (Exception $e) {
            die($e->getMessage());
        }
    }

    public function getValidUrl() : string
    {
        return $this->strValidUrl;
    }

    public function getGid() : string
    {
        return $this->strGid;
    }

    public function getUrl() : string
    {
        return $this->strUrl;
    }

    public function getParams() : array
    {
        return $this->arParams;
    }
}