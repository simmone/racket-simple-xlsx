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
      (send xlsx add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4") (201601 100 110 120)))

      (check-equal? (write-data-sheet "Sheet1" xlsx)
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dimension ref=\"A1:D2\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"13.5\"/><sheetData><row r=\"1\" spans=\"1:4\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"C1\" t=\"s\"><v>2</v></c><c r=\"D1\" t=\"s\"><v>3</v></c></row><row r=\"2\" spans=\"1:4\"><c r=\"A2\"><v>201601</v></c><c r=\"B2\"><v>100</v></c><c r=\"C2\"><v>110</v></c><c r=\"D2\"><v>1110</v></c></row><row r=\"3\" spans=\"1:4\"><c r=\"A3\" s=\"1\"><v>201602</v></c><c r=\"B3\" s=\"1\"><v>200</v></c><c r=\"C3\" s=\"1\"><v>210</v></c><c r=\"D3\"><v>1210</v></c></row><row r=\"4\" spans=\"1:4\"><c r=\"A4\" s=\"1\"><v>201603</v></c><c r=\"B4\" s=\"1\"><v>300</v></c><c r=\"C4\" s=\"1\"><v>310</v></c><c r=\"D4\"><v>1310</v></c></row><row r=\"5\" spans=\"1:4\"><c r=\"A5\" s=\"1\"><v>201604</v></c><c r=\"B5\" s=\"1\"><v>400</v></c><c r=\"C5\" s=\"1\"><v>410</v></c><c r=\"D5\"><v>1410</v></c></row><row r=\"6\" spans=\"1:4\"><c r=\"A6\" s=\"1\"><v>201605</v></c><c r=\"B6\" s=\"1\"><v>500</v></c><c r=\"C6\" s=\"1\"><v>510</v></c><c r=\"D6\"><v>1510</v></c></row><row r=\"7\" spans=\"1:4\"><c r=\"A7\" s=\"1\"><v>201606</v></c><c r=\"B7\" s=\"1\"><v>600</v></c><c r=\"C7\" s=\"1\"><v>610</v></c><c r=\"D7\"><v>1610</v></c></row><row r=\"8\" spans=\"1:4\"><c r=\"A8\" s=\"1\"><v>201607</v></c><c r=\"B8\" s=\"1\"><v>500</v></c><c r=\"C8\" s=\"1\"><v>510</v></c><c r=\"D8\"><v>1510</v></c></row><row r=\"9\" spans=\"1:4\"><c r=\"A9\" s=\"1\"><v>201608</v></c><c r=\"B9\" s=\"1\"><v>400</v></c><c r=\"C9\" s=\"1\"><v>410</v></c><c r=\"D9\"><v>1410</v></c></row><row r=\"10\" spans=\"1:4\"><c r=\"A10\" s=\"1\"><v>201609</v></c><c r=\"B10\" s=\"1\"><v>300</v></c><c r=\"C10\" s=\"1\"><v>310</v></c><c r=\"D10\"><v>1310</v></c></row><row r=\"11\" spans=\"1:4\"><c r=\"A11\" s=\"1\"><v>201610</v></c><c r=\"B11\" s=\"1\"><v>300</v></c><c r=\"C11\" s=\"1\"><v>310</v></c><c r=\"D11\"><v>1310</v></c></row><row r=\"12\" spans=\"1:4\"><c r=\"A12\"><v>201611</v></c><c r=\"B12\"><v>100</v></c><c r=\"C12\"><v>110</v></c><c r=\"D12\"><v>1110</v></c></row><row r=\"13\" spans=\"1:4\"><c r=\"A13\"><v>201612</v></c><c r=\"B13\"><v>100</v></c><c r=\"C13\"><v>110</v></c><c r=\"D13\"><v>1110</v></c></row><row r=\"14\" spans=\"1:4\"><c r=\"A14\"><v>201701</v></c><c r=\"B14\"><v>500</v></c><c r=\"C14\"><v>510</v></c><c r=\"D14\"><v>1510</v></c></row></sheetData><phoneticPr fontId=\"1\" type=\"noConversion\"/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"200\" verticalDpi=\"200\" r:id=\"rId1\"/></worksheet>"))
    ))
   ))

(run-tests test-worksheet)
