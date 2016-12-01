#lang racket

(require "../../xlsx.rkt")

(require rackunit/text-ui)

(require rackunit "workbook.rkt")

(define test-workbook
  (test-suite
   "test-workbook"

   (test-case
    "test-workbook"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet "Sheet1" '((1)))
      (send xlsx add-data-sheet "Sheet2" '((1)))
      (send xlsx add-data-sheet "Sheet3" '((1)))
      (send xlsx add-line-chart-sheet "Chart1" "Chart1")
      (send xlsx add-line-chart-sheet "Chart2" "Chart1")
      (send xlsx add-line-chart-sheet "Chart3" "Chart1")

      (check-equal? (write-workbook (get-field sheets xlsx))
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/><workbookPr filterPrivacy=\"1\" defaultThemeVersion=\"124226\"/><bookViews><workbookView xWindow=\"0\" yWindow=\"90\" windowWidth=\"19200\" windowHeight=\"10590\"/></bookViews><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/><sheet name=\"Sheet3\" sheetId=\"3\" r:id=\"rId3\"/><sheet name=\"Chart1\" sheetId=\"4\" r:id=\"rId4\"/><sheet name=\"Chart2\" sheetId=\"5\" r:id=\"rId5\"/><sheet name=\"Chart3\" sheetId=\"6\" r:id=\"rId6\"/></sheets><calcPr calcId=\"124519\"/></workbook>"))))
    ))

(run-tests test-workbook)
