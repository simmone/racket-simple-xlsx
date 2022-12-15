#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/_rels/workbook-xml-rels.rkt")

(require racket/runtime-path)
(define-runtime-path workbook_xml_rels_file "workbook.xml.rels")


(define test-workbook-xml-rels
  (test-suite
   "test-workbook-xml-rels"

   (test-case
    "test-write-workbook-rels"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (dynamic-wind
           (lambda ()
             (write-workbook-rels (apply build-path (drop-right (explode-path workbook_xml_rels_file) 1))))
           (lambda ()
             (call-with-input-file workbook_xml_rels_file
               (lambda (expected)
                 (call-with-input-string
                  (lists->xml (xl-rels))
                  (lambda (actual)
                    (check-lines? expected actual))))))
           (lambda ()
             (when (file-exists? workbook_xml_rels_file) (delete-file workbook_xml_rels_file))))
       )))
    ))

  (run-tests test-workbook-xml-rels)
