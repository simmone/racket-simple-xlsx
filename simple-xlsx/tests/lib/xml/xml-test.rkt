#lang racket

(require rackunit/text-ui)
(require racket/date)

(require racket/runtime-path)
(define-runtime-path workbook_xml_file "workbook.xml")
(define-runtime-path sharedStrings_xml_file "sharedStrings.xml")
(define-runtime-path sheet_xml_file "sheet.xml")

(require rackunit "../../../lib/xml.rkt")

(define test-xml
  (test-suite
   "test-workbook"

   (test-case
    "test-workbook"

    (let ([xml_hash (load-xml-hash workbook_xml_file '(sheets))])
      (printf "~a\n" xml_hash)
      (check-equal? (hash-count xml_hash) 58)
      (check-equal? (hash-ref xml_hash "workbook.xmlns") "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      (check-equal? (hash-ref xml_hash "fileVersion.appName") "xl")
      (check-equal? (hash-ref xml_hash "workbookView.xWindow") "0")
      (check-equal? (hash-ref xml_hash "calcPr.calcId") "124519")

      (check-equal? (hash-ref xml_hash "sheet.count") 10)

      (check-equal? (hash-ref xml_hash "sheet1.name") "DataSheet")
      (check-equal? (hash-ref xml_hash "sheet1.sheetId") "1")
      (check-equal? (hash-ref xml_hash "sheet1.r:id") "rId1")

      (check-equal? (hash-ref xml_hash "sheet10.name") "PieChart3D")
      (check-equal? (hash-ref xml_hash "sheet10.sheetId") "10")
      (check-equal? (hash-ref xml_hash "sheet10.r:id") "rId10")
      )

    )

;   (test-case
;    "test-shared-string"
;
;    (let ([xml_hash (load-xml-hash sharedStrings_xml_file '(t phoneticPr))])
;      (check-equal? (hash-count xml_hash) 73)
;      (check-equal? (hash-ref xml_hash "sst.count") "17")
;      (check-equal? (hash-ref xml_hash "sst.uniqueCount") "17")
;
;      (check-equal? (hash-ref xml_hash "t.count") 17)
;      (check-equal? (hash-ref xml_hash "phoneticPr.count") 17)
;
;      (check-equal? (hash-ref xml_hash "t1") "")
;      (check-equal? (hash-ref xml_hash "t2") "201601")
;      (check-equal? (hash-ref xml_hash "t10") "Center")
;      (check-equal? (hash-ref xml_hash "t17") "month/brand")
;      )
;    )
;   
;   (test-case
;    "test-load-sheet"
;    
;    (let (
;          [xml_hash (load-xml-hash sheet_xml_file '(row))]
;          )
;      
;      (printf "~a\n" xml_hash)
;
;      (check-equal? (hash-count xml_hash) 60)
;      (check-equal? (hash-ref xml_hash "col.count") 4)
;      (check-equal? (hash-ref xml_hash "row.count") 4)
;
;      ))
  ))

(run-tests test-xml)
