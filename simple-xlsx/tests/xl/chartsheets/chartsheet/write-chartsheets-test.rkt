#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/chartsheets/chartsheet.rkt")

(require racket/runtime-path)
(define-runtime-path sheet1_file "sheet1.xml")
(define-runtime-path sheet2_file "sheet2.xml")
(define-runtime-path sheet3_file "sheet3.xml")

(define test-chart-sheet
  (test-suite
   "test-chart-sheet"

   (test-case
    "test-write-chartsheets"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1" '(("1")))
       (add-data-sheet "Sheet2" '((1)))
       (add-chart-sheet "Chart1" 'LINE "Chart1" '())
       (add-chart-sheet "Chart2" 'LINE "Chart2" '())
       (add-data-sheet "Sheet3" '((1)))
       (add-chart-sheet "Chart3" 'LINE "Chart3" '())

       (dynamic-wind
           (lambda ()
             (write-chartsheets (apply build-path (drop-right (explode-path sheet1_file) 1))))
           (lambda ()
             (call-with-input-file sheet1_file
               (lambda (expected1)
                 (call-with-input-file sheet2_file
                   (lambda (expected2)
                     (call-with-input-file sheet3_file
                       (lambda (expected3)
                         (call-with-input-string
                          (lists->xml (chart-sheet 1))
                          (lambda (actual)
                            (check-lines? expected1 actual)))
                         (call-with-input-string
                          (lists->xml (chart-sheet 2))
                          (lambda (actual)
                            (check-lines? expected2 actual)))
                         (call-with-input-string
                          (lists->xml (chart-sheet 3))
                          (lambda (actual)
                            (check-lines? expected3 actual))))))))))
           (lambda ()
             (when (file-exists? sheet1_file) (delete-file sheet1_file))
             (when (file-exists? sheet2_file) (delete-file sheet2_file))
             (when (file-exists? sheet3_file) (delete-file sheet3_file))
             )))))

   ))


(run-tests test-chart-sheet)
