<?php

use Google\Service\Sheets;
use Google\Service\Sheets\GridRange;
use Google\Service\Sheets\Spreadsheet;
use Google\Client;


class Service
{
    protected Spreadsheet $spreadsheet;
    protected Client $client;
    protected Sheets $sheetService;
    protected array $sheetData = [];

    public function __construct(Client $client)
    {
        $this->client = $client;
        $this->sheetService = new Sheets($this->client);
        $this->spreadsheet = $this->sheetService->spreadsheets->get($this->getSheetDocument());

        if($this->sheet === null) {
            throw new SyncException("Not found sheet {$this->getSheetPage()} in excel document");
        }

        $this->sheetData = $this->sheetService->spreadsheets_values->get(
            "Enter id document",
            "A:I" // enter columns for load for work
        )->getValues();
    }

    public function getRangeData(GridRange $range, bool $withTitle = false): array
    {
        $data = [];
        $titles = [];
        for ($currentRow = $range->getStartRowIndex();
             $currentRow < $range->getEndRowIndex();
             $currentRow++
        ) {
            $rowData = $this->sheetData[$currentRow];
            $rowRangeData = [];
            for ($currentColumn = $range->getStartColumnIndex();
                 $currentColumn < $range->getEndColumnIndex();
                 $currentColumn++
            ) {
                $columnValue = isset($rowData[$currentColumn]) ? $rowData[$currentColumn] : null;

                if($withTitle && $currentRow !== $range->getStartRowIndex()) {
                    $rowRangeData[$titles[$currentColumn]] = $columnValue;
                } else {
                    $rowRangeData[] = $columnValue;
                }
            }

            if($withTitle && $currentRow === $range->getStartRowIndex()) {
                $titles = $rowRangeData;
                continue;
            }

            $data[] = $rowRangeData;
        }

        return $data;
    }

    public function getRangeByName(string $name): ?GridRange
    {
        $ranges = $this->spreadsheet->getNamedRanges();

        foreach ($ranges as $namedRange) {
            if($namedRange->getName() === $name) {
                return $namedRange->getRange();
            }
        }

        return null;
    }
}


// Inti google api client
// $client =  new Client()

$service = new Service($client);

$range = $service->getRangeByName("MY_RANGE");
$data = $service->getRangeData($range, true); // if withTitle true: remove first row and use values from 1 row for keys in array
