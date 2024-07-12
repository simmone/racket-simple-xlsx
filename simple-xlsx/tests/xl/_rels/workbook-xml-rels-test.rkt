#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/_rels/workbook-xml-rels.rkt"
         racket/runtime-path)

(define-runtime-path rels1_file "workbook.xml.rels1")
(define-runtime-path rels2_file "workbook.xml.rels2")

(define test-workbook-xml-rels
  (test-suite
   "test-workbook-xml-rels"

   (test-case
    "test-workbook-xml-rels1"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '((1)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (call-with-input-file rels1_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml (xl-rels))
            (lambda (actual)
              (check-lines? expected actual)))))))

    (test-case
     "test-workbook-xml-rels2"

     (with-xlsx
      (lambda ()
        (add-data-sheet "Sheet1" '((1)))
        (add-data-sheet "Sheet2" '((1)))
        (add-data-sheet "Sheet3" '((1)))
        (add-chart-sheet "Chart1" 'LINE "Chart1" '())
        (add-chart-sheet "Chart2" 'LINE "Chart2" '())
        (add-chart-sheet "Chart3" 'LINE "Chart3" '())

        (call-with-input-file rels2_file
          (lambda (expected)
            (call-with-input-string
             (lists-to-xml (xl-rels))
             (lambda (actual)
               (check-lines? expected actual)))))))))
     ))

   (run-tests test-workbook-xml-rels)
