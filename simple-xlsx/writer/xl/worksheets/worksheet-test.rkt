#lang racket

(require "../../../xlsx.rkt")

(require rackunit/text-ui)

(require rackunit "worksheet.rkt")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-write-data-sheet"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4") (201601 100 110 1110)))

      (check-equal? (write-data-sheet "Sheet1" xlsx)
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dimension ref=\"A1:D2\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"13.5\"/><sheetData><row r=\"1\" spans=\"1:4\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"C1\" t=\"s\"><v>2</v></c><c r=\"D1\" t=\"s\"><v>3</v></c></row><row r=\"2\" spans=\"1:4\"><c r=\"A2\"><v>201601</v></c><c r=\"B2\"><v>100</v></c><c r=\"C2\"><v>110</v></c><c r=\"D2\"><v>1110</v></c></row></sheetData><phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"200\" verticalDpi=\"200\" r:id=\"rId1\"/></worksheet>"))
    ))
   ))

(run-tests test-worksheet)
