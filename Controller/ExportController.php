<?php

/** @noinspection PhpMultipleClassDeclarationsInspection */

namespace Newageerp\SfExport\Controller;

use Doctrine\ORM\EntityManagerInterface;
use Exception;
use Newageerp\SfAuth\Service\AuthService;
use Newageerp\SfUservice\Controller\UControllerBase;
use Newageerp\SfUservice\Service\UService;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use Symfony\Component\EventDispatcher\EventDispatcherInterface;

use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Annotation\Route;
use OpenApi\Annotations as OA;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Newageerp\SfSocket\Service\SocketService;

/**
 * @Route(path="/app/nae-core/export")
 */
class ExportController extends UControllerBase
{
    protected array $letters = [];

    protected array $headerStyle = [
        'font' => [
            'bold' => true,
        ],
    ];

    protected function applyStyleToRow($sheet, int $row, array $style)
    {
        $sheet->getStyle('A' . $row . ':X' . $row)->applyFromArray($style);
    }

    protected function applyStyleToCell($sheet, string $cell, array $style)
    {
        $sheet->getStyle($cell)->applyFromArray($style);
    }

    /**
     * OaListController constructor.
     */
    public function __construct(EntityManagerInterface $em, EventDispatcherInterface $eventDispatcher, SocketService $socketService)
    {
        parent::__construct($em, $eventDispatcher, $socketService);
        $this->letters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'X', 'Y', 'Z'];
    }

    /**
     * @Route(path="/doExport")
     * @OA\Post (operationId="NAEUExport")
     */
    public function doExport(Request $request, UService $uService): JsonResponse
    {
        try {
            $storageDir = $_ENV['NAE_SFS_PUBLIC_DOC_DIR'];
            $exportDir = '/public/export';

            $request = $this->transformJsonBody($request);

            $user = $this->findUser($request);
            if (!$user) {
                throw new Exception('Invalid user');
            }
            AuthService::getInstance()->setUser($user);

            $schema = $request->get('schema');
            $title = $request->get('title');
            $exportOptions = $request->get('exportOptions');
            $fields = $request->get('fields');
            $columns = $request->get('columns');

            $fieldsToReturn = $exportOptions['fieldsToReturn'] ?? ['id'];

            $data = $uService->getListDataForSchema(
                $schema,
                1,
                9999999,
                $fieldsToReturn,
                $exportOptions['filter'] ?? [],
                $exportOptions['extraData'] ?? [],
                $exportOptions['sort'] ?? [],
                $exportOptions['totals'] ?? []
            )['data'];
            $recordsCount = count($data);

            $properties = $this->getPropertiesForSchema($schema);

            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setCellValue('A1', $title);
            $sheet->setCellValue('F1', 'Cols');
            $sheet->setCellValue('G1', count($fieldsToReturn));

            $fileName = $exportDir . '/' . $title . '_' . time() . '.xlsx';

            $this->applyStyleToRow($sheet, 3, $this->headerStyle);

            $parseColumns = $columns ?: array_map(function ($field) use ($schema) {
                $field['path'] = isset($field['relName']) ?
                    $schema . '.' . $field['key'] . '.' . $field['relName'] :
                    $schema . '.' . $field['key'];
                return $field;
            }, $fields);

            $parseColumns = array_map(function ($field) use ($properties) {
                $pathArray = explode(".", $field['path']);
                $relName = null;
                if (count($pathArray) === 3) {
                    [$schema, $fieldKey, $relName] = $pathArray;
                } else {
                    [$schema, $fieldKey] = $pathArray;
                }

                $field['schema'] = $schema;
                $field['fieldKey'] = $fieldKey;
                $field['title'] = isset($field['customTitle']) && $field['customTitle'] ? $field['customTitle'] : $properties[$fieldKey]['title'];
                $field['pivotTitle'] = isset($field['pivotCustomTitle']) ? $field['pivotCustomTitle'] : $field['title'];
                return $field;
            }, $parseColumns);

            $hasPivot = false;
            foreach ($parseColumns as $col) {
                if (isset($col['pivotSetting']) && $col['pivotSetting']) {
                    $hasPivot = true;
                }
            }

            $col = 1;

            foreach ($parseColumns as $field) {
                $pathArray = explode(".", $field['path']);
                $relName = null;
                if (count($pathArray) === 3) {
                    [$schema, $fieldKey, $relName] = $pathArray;
                } else {
                    [$schema, $fieldKey] = $pathArray;
                }

                $title = $field['title'];
                $sheet->setCellValue($this->letters[$col] . '3', $title);

                if (isset($field['allowEdit']) && $field['allowEdit']) {
                    $sheet->setCellValue($this->letters[$col] . '2', $fieldKey);

                    $sheet
                        ->getStyle($this->letters[$col] . '2:' . $this->letters[$col] . ($recordsCount + 3))
                        ->getFill()
                        ->setFillType(Fill::FILL_SOLID)
                        ->getStartColor()
                        ->setARGB('FFCFE2F3');
                }
                $col++;
            }
            $row = 4;
            $pivotRow = 0;
            $pivotData = [];
            foreach ($data as $item) {
                $col = 1;
                $pivotCol = 0;
                foreach ($parseColumns as $field) {
                    $pathArray = explode(".", $field['path']);
                    $relName = null;
                    if (count($pathArray) === 3) {
                        [$schema, $fieldKey, $relName] = $pathArray;
                    } else {
                        [$schema, $fieldKey] = $pathArray;
                    }

                    $val = $relName && isset($item[$fieldKey]) && $item[$fieldKey] ?
                        $item[$fieldKey][$relName] :
                        $item[$fieldKey];

                    if (isset($properties[$fieldKey]['enum']) && $properties[$fieldKey]['enum']) {
                        foreach ($properties[$fieldKey]['enum'] as $p) {
                            if ($p['value'] === $val) {
                                $val = $p['label'];
                                break;
                            }
                        }
                    }

                    $sheet->getCellByColumnAndRow($col, $row)->setValue($val);

                    if (!isset($pivotData[$pivotRow])) {
                        $pivotData[$pivotRow] = [
                            -1 => '-'
                        ];
                    }
                    $pivotData[$pivotRow][$pivotCol] = $val;

                    $col++;
                    $pivotCol++;
                }
                $row++;
                $pivotRow++;
            }

            foreach ($this->letters as $letter) {
                if ($letter) {
                    $sheet->getColumnDimension($letter)->setAutoSize(true);
                }
            }

            if ($hasPivot) {
                $pivotRowTitle = '';
                $pivotColTitle = '';
                $pivotTotalTitles = [];
                $pivotTotalIndexes = [];
                $pivotTotalTypes = [];
                

                $pivotRowIndex = -1;
                $pivotColIndex = -1;

                $pivotSheet = $spreadsheet->createSheet(1);
                $pivotSheet->setTitle('Ataskaita');

                foreach ($parseColumns as $colIndex => $col) {
                    if (isset($col['pivotSetting'])) {
                        if ($col['pivotSetting'] === 'row') {
                            $pivotRowTitle = $col['pivotTitle'];
                            $pivotRowIndex = $colIndex;
                        }
                        if ($col['pivotSetting'] === 'col') {
                            $pivotColTitle = $col['pivotTitle'];
                            $pivotColIndex = $colIndex;
                        }
                        if ($col['pivotSetting'] === 'total' || $col['pivotSetting'] === 'count') {
                            $pivotTotalTitles[] = $col['pivotTitle'];
                            $pivotTotalIndexes[] = $colIndex;
                            $pivotTotalTypes[] = $col['pivotSetting'];
                        }
                    }
                }
                $pivotTotalsCount = count($pivotTotalTitles);

                $pivotSheet->getCellByColumnAndRow(1, 3)->setValue($pivotRowTitle);
                $pivotSheet->getCellByColumnAndRow(2, 1)->setValue($pivotColTitle);

                $this->applyStyleToRow($sheet, 2, $this->headerStyle);
                $this->applyStyleToCell($sheet, 'A3', $this->headerStyle);


                $pivotRowValues = array_values(
                    array_unique(
                        array_map(
                            function ($item) use ($pivotRowIndex) {
                                return $item[$pivotRowIndex];
                            },
                            $pivotData
                        )
                    )
                );

                $pivotColValues = array_values(
                    array_unique(
                        array_map(
                            function ($item) use ($pivotColIndex) {
                                return $item[$pivotColIndex];
                            },
                            $pivotData
                        )
                    )
                );

                foreach ($pivotRowValues as $rowIndex => $rowValue) {
                    $pivotSheet->getCellByColumnAndRow(1, 4 + $rowIndex)->setValue($rowValue);
                }
                foreach ($pivotColValues as $colIndex => $colValue) {
                    $pivotSheet->getCellByColumnAndRow(2 + ($colIndex * $pivotTotalsCount), 2)->setValue($colValue);
                    foreach ($pivotTotalTitles as $totalIndex => $totalTitle) {
                        $pivotSheet->getCellByColumnAndRow(2 + ($colIndex * $pivotTotalsCount) + $totalIndex, 3)->setValue($totalTitle);
                    }
                }

                foreach ($pivotRowValues as $rowIndex => $rowValue) {
                    foreach ($pivotColValues as $colIndex => $colValue) {
                        foreach ($pivotTotalTitles as $totalIndex => $totalTitle) {
                            $colData = array_map(
                                function ($item) use ($totalIndex, $pivotTotalIndexes) {
                                    return $item[$pivotTotalIndexes[$totalIndex]];
                                },
                                array_filter(
                                    $pivotData,
                                    function ($item) use ($pivotRowIndex, $pivotColIndex, $rowValue, $colValue) {
                                        return $item[$pivotRowIndex] === $rowValue && $item[$pivotColIndex] === $colValue;
                                    }
                                )
                            );
                            $val = 0;
                            if ($pivotTotalTypes[$totalIndex] === 'count') {
                                $val = count(array_unique($colData));
                            } else if ($pivotTotalTypes[$totalIndex] === 'total') {
                                $val = array_sum($colData);
                            }
                            $pivotSheet->getCellByColumnAndRow(2 + ($colIndex * $pivotTotalsCount) + $totalIndex, 4 + $rowIndex)->setValue($val);
                        }
                    }
                }

                foreach ($this->letters as $letter) {
                    if ($letter) {
                        $pivotSheet->getColumnDimension($letter)->setAutoSize(true);
                    }
                }
            }

            $writer = new Xlsx($spreadsheet);
            $writer->save($storageDir . $fileName);

            return $this->json([
                'fileName' => $fileName
            ]);
        } catch (Exception $e) {
            $response = $this->json([
                'description' => $e->getMessage(),
                'f' => $e->getFile(),
                'l' => $e->getLine(),
                'fileName' => 'about:blank'

            ]);
            $response->setStatusCode(Response::HTTP_BAD_REQUEST);
            return $response;
        }
    }
}
