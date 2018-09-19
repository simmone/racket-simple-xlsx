#lang racket

(require "../../xlsx/xlsx.rkt")

(require rackunit/text-ui)

(require rackunit "workbook.rkt")

(define test-workbook
  (test-suite
   "test-workbook"

   (test-case
    "test-workbook"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "Sheet1" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet2" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet3" #:sheet_data '((1)))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart2" #:topic "Chart2" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart3" #:topic "Chart3" #:x_topic "")

      (check-equal? (write-workbook (get-field sheets xlsx))
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/><workbookPr filterPrivacy=\"1\" defaultThemeVersion=\"124226\"/><bookViews><workbookView xWindow=\"0\" yWindow=\"90\" windowWidth=\"19200\" windowHeight=\"10590\"/></bookViews><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/><sheet name=\"Sheet3\" sheetId=\"3\" r:id=\"rId3\"/><sheet name=\"Chart1\" sheetId=\"4\" r:id=\"rId4\"/><sheet name=\"Chart2\" sheetId=\"5\" r:id=\"rId5\"/><sheet name=\"Chart3\" sheetId=\"6\" r:id=\"rId6\"/></sheets><calcPr calcId=\"124519\"/></workbook>"))))
    ))

(run-tests test-workbook)
