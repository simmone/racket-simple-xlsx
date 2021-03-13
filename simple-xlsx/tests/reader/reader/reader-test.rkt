#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader.rkt")

(require "../../../xlsx/xlsx.rkt")

(require racket/runtime-path)
(define-runtime-path test1_file "test1.xlsx")

(define test-reader
  (test-suite
   "test-reader"
   
   (test-case
    "test-reader"

    (with-input-from-xlsx-file
     test1_file
     (lambda (xlsx)
      (check-equal? (XLSX-sheet_count xlsx) 10)
      ))
    
    )))

(run-tests test-reader)
