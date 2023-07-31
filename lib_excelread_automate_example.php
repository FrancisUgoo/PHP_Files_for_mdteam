<?php
function parsexml() {
    require 'vendor/autoload.php'; // Load PhpSpreadsheet library

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\IOFactory;

    $dirx = readline('Please enter the folder fullpath (remove any \ at the end):' . PHP_EOL);
    $num = intval(readline('Please enter the number of xml files in folder:' . PHP_EOL));
    $exfile = readline('Please enter the fullpath and filename of excel file:' . PHP_EOL);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setCellValue('A1', ' ');
    $sheet->setCellValue('B1', ' ');
    $sheet->setCellValue('C1', ' ');

    $count = 1;

    while ($num >= $count) {
        $ndirx = $dirx . '/' . $count . '.xml';
        // $ndirx = $dirx . '\\' . $count . '.xml';

        if (file_exists($ndirx)) {
            $xml = new DOMDocument();
            $xml->load($ndirx);

            $title = $xml->getElementsByTagName("DISS_title")[0]->nodeValue;
            $fname = $xml->getElementsByTagName("DISS_fname")[0]->nodeValue;
            $mname = $xml->getElementsByTagName("DISS_middle")[0]->nodeValue;
            $sname = $xml->getElementsByTagName("DISS_surname")[0]->nodeValue;

            $adfname = $xml->getElementsByTagName("DISS_fname")[1]->nodeValue;
            $admname = $xml->getElementsByTagName("DISS_middle")[1]->nodeValue;
            $adsname = $xml->getElementsByTagName("DISS_surname")[1]->nodeValue;
            $pubyear = $xml->getElementsByTagName("DISS_comp_date")[0]->nodeValue;
            $degree = $xml->getElementsByTagName("DISS_degree")[0]->nodeValue;
            $department = $xml->getElementsByTagName("DISS_inst_contact")[0]->nodeValue;

            $paras = $xml->getElementsByTagName("DISS_para");
            $abstra = " ";
            $Keywords = " ";

            foreach ($paras as $para) {
                if ($para->firstChild) {
                    $keyw = substr($para->firstChild->nodeValue, 0, 9);
                    if ($keyw == 'Keywords:') {
                        $csss = $count + 1;
                        $C1 = "C" . $csss;
                        $akeyword = $para->firstChild->nodeValue;
                        $sheet->setCellValue($C1, substr($akeyword, 10));
                    }
                    if (strlen($para->firstChild->nodeValue) > strlen($abstra)) {
                        $abstra = $para->firstChild->nodeValue;
                    }
                }
            }

            $cs = $count + 1;

            $A1 = "A" . $cs;
            $sheet->setCellValue($A1, $title);
            $D1 = "D" . $cs;
            $sheet->setCellValue($D1, $abstra);
            $E1 = "E" . $cs;
            $sheet->setCellValue($E1, $fname);
            if ($mname) {
                $F1 = "F" . $cs;
                $sheet->setCellValue($F1, $mname);
            }
            $G1 = "G" . $cs;
            $sheet->setCellValue($G1, $sname);

            $advisor = $adfname . ' ' . $admname . ' ' . $adsname;
            $K1 = "K" . $cs;
            $sheet->setCellValue($K1, $advisor);
            $R1 = "R" . $cs;
            $sheet->setCellValue($R1, $pubyear);

            if ($degree) {
                $O1 = "O" . $cs;
                $Q1 = "Q" . $cs; //doc type
                $degr = strtolower($degree);
                $degr = str_replace(".", "", $degr);

                if (substr($degr, 0, 3) == 'edd' || substr($degr, 0, 3) == 'phd') {
                    $sheet->setCellValue($Q1, 'dissertation');
                } else {
                    $sheet->setCellValue($Q1, 'thesis');
                }

                switch (substr($degr, 0, 3)) {
                    case 'edd':
                        $sheet->setCellValue($O1, 'Doctor of Education (Ed.D)');
                        break;
                    case 'phd':
                        $sheet->setCellValue($O1, 'Doctor of Philosophy (Ph.D)');
                        break;
                    case 'msa':
                        $sheet->setCellValue($O1, 'Master of Accountancy (MSA)');
                        break;
                    case 'mba':
                        $sheet->setCellValue($O1, 'Master of Business Administration (MBA)');
                        break;
                    case 'med':
                        $sheet->setCellValue($O1, 'Master of Education (MED)');
                        break;
                    case 'mfa':
                        $sheet->setCellValue($O1, 'Master of Fine Arts (MFA)');
                        break;
                    case 'mph':
                        $sheet->setCellValue($O1, 'Master of Public Health (MPH)');
                        break;
                    case 'msn':
                        $sheet->setCellValue($O1, 'Master of Science in Nursing (MSN)');
                        break;
                    case 'msw':
                        $sheet->setCellValue($O1, 'Master of Social Work (MSW)');
                        break;
                    case 'ssp':
                        $sheet->setCellValue($O1, 'Specialist in School Psychology (SSP)');
                        break;
                    case 'ma':
                        $sheet->setCellValue($O1, 'Master of Arts (MA)');
                        break;
                    case 'mm':
                        $sheet->setCellValue($O1, 'Master of Music (MM)');
                        break;
                    case 'ms':
                        $sheet->setCellValue($O1, 'Master of Science (MS)');
                        break;
                    default:
                        $sheet->setCellValue($O1, $degree);
                        break;
                }
            }

            if ($department) {
                $P1 = "P" . $cs;
                $dept = $department;
                $f3dept = strtolower(substr($dept, 0, 4));
                $f3dept = str_replace(".", "", $f3dept);
                $f3dept = str_replace("-", " ", $f3dept);
                $f2dept = strtolower(substr($dept, 0, 3));
                $f2dept = str_replace(".", "", $f2dept);
                $f2dept = str_replace("-", " ", $f2dept);

                if ($f3dept == 'edd' || $f3dept == 'phd' || $f3dept == 'msa' || $f3dept == 'mba' || $f3dept == 'med' || $f3dept == 'mfa' || $f3dept == 'mph' || $f3dept == 'msn' || $f3dept == 'msw' || $f3dept == 'ssp') {
                    $sheet->setCellValue($P1, substr($dept, 4));
                } elseif ($f2dept == 'ma' || $f2dept == 'mm' || $f2dept == 'ms') {
                    $sheet->setCellValue($P1, substr($dept, 3));
                } else {
                    $sheet->setCellValue($P1, $dept);
                }
            }
        } else {
            echo 'Incorrect path or Not a text file.' . PHP_EOL;
        }

        $count++;
    }

    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($exfile);

    return 0;
}

parsexml();
