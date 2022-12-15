#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/chartsheets/_rels/chartsheets-rels.rkt")

(require racket/runtime-path)
(define-runtime-path sheet1_rels_file "sheet1.xml.rels")
(define-runtime-path sheet2_rels_file "sheet2.xml.rels")
(define-runtime-path sheet3_rels_file "sheet3.xml.rels")

(define test-chartsheets-rels
  (test-suite
   "test-chartsheets-rels"

   (test-case
    "test-write-chartsheets-rels"

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
             (write-chartsheets-rels (apply build-path (drop-right (explode-path sheet1_rels_file) 1))))
           (lambda ()
             (call-with-input-file sheet1_rels_file
               (lambda (expected1)
                 (call-with-input-file sheet2_rels_file
                   (lambda (expected2)
                     (call-with-input-file sheet3_rels_file
                       (lambda (expected3)
                         (call-with-input-string
                          (lists->xml (chartsheets-rels 1))
                          (lambda (actual)
                            (check-lines? expected1 actual)))
                         (call-with-input-string
                          (lists->xml (chartsheets-rels 2))
                          (lambda (actual)
                            (check-lines? expected2 actual)))
                         (call-with-input-string
                          (lists->xml (chartsheets-rels 3))
                          (lambda (actual)
                            (check-lines? expected3 actual))))))))))
           (lambda ()
             (when (file-exists? sheet1_rels_file) (delete-file sheet1_rels_file))
             (when (file-exists? sheet2_rels_file) (delete-file sheet2_rels_file))
             (when (file-exists? sheet3_rels_file) (delete-file sheet3_rels_file))
             )))))
    ))

  (run-tests test-chartsheets-rels)
