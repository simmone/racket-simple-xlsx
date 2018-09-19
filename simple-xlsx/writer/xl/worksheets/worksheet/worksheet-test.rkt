#lang racket

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../lib/lib.rkt")

(require rackunit/text-ui)

(require rackunit "worksheet.rkt")

(require racket/runtime-path)
(define-runtime-path test_file "worksheet-test.dat")

(define test-worksheet
  (test-suite
   "test-worksheet"

   (test-case
    "test-write-data-sheet"

    (let ([xlsx (new xlsx%)])
      (send xlsx add-data-sheet 
            #:sheet_name "Sheet1" 
            #:sheet_data '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))

      (call-with-input-file test_file
        (lambda (expected)
          (call-with-input-string
           (write-data-sheet "Sheet1" xlsx)
           (lambda (actual)
             (check-lines? expected actual)))))))
   
   (test-case
    "test-get-col-width-map"
    
    (let* ([rows '(
                  ("123" "3456" "陈晓" "出色的咖34的肯定")
                  ("12" "345678" "陈晓快速的口" "出色34的肯定")
                  )]
           [col_width_map (get-col-width-map rows)])
      (check-equal? (hash-ref col_width_map 1) 5)
      (check-equal? (hash-ref col_width_map 2) 8)
      (check-equal? (hash-ref col_width_map 3) 14)
      (check-equal? (hash-ref col_width_map 4) 18)
      ))
   ))

(run-tests test-worksheet)
