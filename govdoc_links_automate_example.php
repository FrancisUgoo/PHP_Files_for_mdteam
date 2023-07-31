<?php
function linkxml() {
    require 'vendor/autoload.php'; // Load PhpSpreadsheet library

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\IOFactory;

    echo "Before we begin, You need to do the following:" . PHP_EOL;
    echo "1. Get the full path to the excel (.xlsx) files. The full path must contain the file name and extension at the end" . PHP_EOL;
    echo "2. Get the number of the start row and last row you wish to update now." . PHP_EOL;
    echo "3. Finally, If any of the excel files are open in your system, kindly close them now" . PHP_EOL;
    readline("If you are ready, Click Enter and let's begin" . PHP_EOL);

    $oldfile = readline('Please enter the fullpath and filename of the excel file:' . PHP_EOL);

    $startnum_ = intval(readline('Please enter the start row number:' . PHP_EOL));
    $oldnum_ = intval(readline('Please enter the last row number:' . PHP_EOL));

    echo "Please Wait Processing..." . PHP_EOL;

    $empty_cell = ['pattern' => 'solid', 'fgColor' => ['rgb' => '87ceeb']];
    $ntsure_cell = ['pattern' => 'solid', 'fgColor' => ['rgb' => 'ffc9bb']];
    $ntrec_cell = ['pattern' => 'solid', 'fgColor' => ['rgb' => 'e32428']];
    $manually_cell = ['pattern' => 'solid', 'fgColor' => ['rgb' => '4a521e']];

    $startnum = $startnum_ + 1;
    $oldnum = $oldnum_ + 1;

    if (file_exists($oldfile)) {
        $spreadsheet = IOFactory::load($oldfile);
        $nsheetold = $spreadsheet->getActiveSheet();

        for ($n = $startnum; $n < $oldnum; $n++) {
            $ns = strval($n);
            echo $n . PHP_EOL;
            $titlecolno = "B" . $ns;
            $pubcolno = "C" . $ns;
            $urlcolno = "D" . $ns;
            $statuscolno = "E" . $ns;
            $urlbb = strval($nsheetold->getCell($urlcolno)->getValue());

            try {
                $titl1 = strval($nsheetold->getCell($titlecolno)->getValue());
                $titl1 = strval($titl1);
                if (strlen($titl1) < 11) {
                    $titl1 = "zzzzzzzzzzzzzzzzzzz";
                }
                $titl1 = str_replace("[", "", $titl1);
                $titl1 = str_replace("]", "", $titl1);
                $titl1 = str_replace(":", "", $titl1);
                $titl = substr($titl1, 0, 9);

                $pubb1 = strval($nsheetold->getCell($pubcolno)->getValue());
                $pubb1 = strval($pubb1);
                if (strlen($pubb1) < 13) {
                    $pubb1 = "zzzzzzzzzzzzzzzzzzzzzzzzzz";
                }
                $pubb1 = str_replace("[", "", $pubb1);
                $pubb1 = str_replace("]", "", $pubb1);
                $pubb1 = str_replace(":", "", $pubb1);
                $pubb = substr($pubb1, 0, 12);

                $urla = requests::get($urlbb);
                $content_type = $urla->headers->get('content-type');

                if (strpos($content_type, 'application/pdf') !== false) {
                    $ext = '.pdf';
                    $nsheetold->setCellValue($statuscolno, "Good Link to a Pdf");
                } elseif (strpos($content_type, 'text/html') !== false) {
                    $ext = '.html';
                    $htmltext = $urla->text;
                    try {
                        if (preg_match('/' . $titl . '/', $htmltext) || preg_match('/' . $pubb . '/', $htmltext)) {
                            $nsheetold->setCellValue($statuscolno, "Good Link to a webpage");
                        } else {
                            $nsheetold->setCellValue($statuscolno, "Not Sure about this Link, advised to confirm");
                            $nsheetold->getStyle($statuscolno)->applyFromArray($ntsure_cell);
                        }
                    } catch (requests::exceptions::RequestException $err) {
                        $nsheetold->setCellValue($statuscolno, "Not Sure about this Link, advised to confirm");
                        $nsheetold->getStyle($statuscolno)->applyFromArray($ntsure_cell);
                    }
                } else {
                    $ext = '';
                    $nsheetold->setCellValue($statuscolno, "Link not Recognized");
                    $nsheetold->getStyle($statuscolno)->applyFromArray($ntrec_cell);
                }
            } catch (requests::exceptions::RequestException $err) {
                $nsheetold->setCellValue($statuscolno, "Please confirm this Link manually");
                $nsheetold->getStyle($statuscolno)->applyFromArray($manually_cell);
                if (strlen($urlbb) < 6) {
                    $nsheetold->setCellValue($statuscolno, "Empty");
                    $nsheetold->getStyle($statuscolno)->applyFromArray($empty_cell);
                }
            }
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($oldfile);

        echo "We are done!!!" . PHP_EOL;
    } else {
        echo "Read Me Errorr!!! The Excel file paths entered were wrong, please try again" . PHP_EOL;
    }

    return 0;
}

try {
    linkxml();
} catch (Exception $e) {
    $feer = " ";
}
?>
