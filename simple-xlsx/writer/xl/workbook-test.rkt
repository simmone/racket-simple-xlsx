#lang racket

(require rackunit/text-ui)

(require rackunit "workbook.rkt")

(define test-workbook
  (test-suite
   "test-workbook"

   (test-case
    "test-workbook"
    (check-equal? (write-workbook '("陈晓" "陈思衡" "陈敏"))
                  (string-append
                   "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><fileVersion appName=\"xl\" lastEdited=\"4\" lowestEdited=\"4\" rupBuild=\"4505\"/><workbookPr filterPrivacy=\"1\" defaultThemeVersion=\"124226\"/><bookViews><workbookView xWindow=\"0\" yWindow=\"90\" windowWidth=\"19200\" windowHeight=\"10590\"/></bookViews><sheets><sheet name=\"陈晓\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"陈思衡\" sheetId=\"2\" r:id=\"rId2\"/><sheet name=\"陈敏\" sheetId=\"3\" r:id=\"rId3\"/></sheets><calcPr calcId=\"124519\"/></workbook>")))
   ))

(run-tests test-workbook)
