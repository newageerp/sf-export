<?php /** @noinspection PhpMultipleClassDeclarationsInspection */

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

    /**
     * OaListController constructor.
     */
    public function __construct(EntityManagerInterface $em, EventDispatcherInterface $eventDispatcher)
    {
        parent::__construct($em, $eventDispatcher);
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

            $col = 1;

            foreach ($parseColumns as $field) {
                $pathArray = explode(".", $field['path']);
                $relName = null;
                if (count($pathArray) === 3) {
                    [$schema, $fieldKey, $relName] = $pathArray;
                } else {
                    [$schema, $fieldKey] = $pathArray;
                }

                $sheet->setCellValue($this->letters[$col] . '3', $properties[$fieldKey]['title']);

                if (isset($field['allowEdit'])) {
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
            foreach ($data as $item) {
                $col = 1;
                foreach ($parseColumns as $field) {
                    $pathArray = explode(".", $field['path']);
                    $relName = null;
                    if (count($pathArray) === 3) {
                        [$schema, $fieldKey, $relName] = $pathArray;
                    } else {
                        [$schema, $fieldKey] = $pathArray;
                    }

                    $val = $relName ?
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
                    $col++;
                }
                $row++;
            }

            foreach ($this->letters as $letter) {
                if ($letter) {
                    $sheet->getColumnDimension($letter)->setAutoSize(true);
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