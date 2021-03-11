#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader/load-shared-strings.rkt")

(require "../../../xlsx/xlsx.rkt")

(require racket/runtime-path)
(define-runtime-path sharedStrings_file "sharedStrings.xml")

(define test-load-shared-strings
  (test-suite
   "test-load-shared-strings"
   
   (test-case
    "test-load-shared-strings"

    (let ([_xlsx (new-xlsx)])
      (load-shared-strings sharedStrings_file _xlsx)

      (check-equal? (hash-count (XLSX-shared_strings_map _xlsx)) 17)
      (check-equal? (hash-ref (XLSX-shared_strings_map _xlsx) "1") "")
      (check-equal? (hash-ref (XLSX-shared_strings_map _xlsx) "2") "201601")
      (check-equal? (hash-ref (XLSX-shared_strings_map _xlsx) "17") "month/brand")
      ))))

(run-tests test-load-shared-strings)
