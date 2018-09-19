#lang racket

(require rackunit/text-ui)

(require "../../../xlsx/xlsx.rkt")

(require rackunit "workbook-xml-rels.rkt")

(define test-workbook-xml-rels
  (test-suite
   "test-workbook-xml-rels"

   (test-case
    "test-workbook-xml-rels"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "数据页面" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet3" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet2" #:sheet_data '((1)))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart4" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart5" #:topic "Chart1" #:x_topic "")


      (check-equal? (write-workbook-xml-rels xlsx)
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet1.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet2.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet3.xml\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>")))

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet #:sheet_name "数据页面" #:sheet_data '(("1")))
      (send xlsx add-data-sheet #:sheet_name "Sheet3" #:sheet_data '((1)))
      (send xlsx add-data-sheet #:sheet_name "Sheet2" #:sheet_data '((1)))
      (send xlsx add-chart-sheet #:sheet_name "Chart1" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart4" #:topic "Chart1" #:x_topic "")
      (send xlsx add-chart-sheet #:sheet_name "Chart5" #:topic "Chart1" #:x_topic "")

      (check-equal? (write-workbook-xml-rels xlsx)
                    (string-append
                     "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                     "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet1.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet2.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" Target=\"chartsheets/sheet3.xml\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>")))
    )
   ))

(run-tests test-workbook-xml-rels)
