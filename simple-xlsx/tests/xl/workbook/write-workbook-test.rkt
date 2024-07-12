#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../xlsx/xlsx.rkt"
         "../../../lib/lib.rkt"
         "../../../xl/workbook.rkt"
         racket/runtime-path)

(define-runtime-path workbook_file "workbook.xml")

(define test-workbook
  (test-suite
   "test-workbook"

   (test-case
    "test-workbook"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("1")))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (dynamic-wind
           (lambda ()
             (write-workbook (apply build-path (drop-right (explode-path workbook_file) 1))))
           (lambda ()
             (call-with-input-file workbook_file
               (lambda (expected)
                 (call-with-input-string
                  (lists-to-xml (to-workbook))
                  (lambda (actual)
                    (check-lines? expected actual))))))
           (lambda ()
             (when (file-exists? workbook_file) (delete-file workbook_file)))))))
   ))

(run-tests test-workbook)
