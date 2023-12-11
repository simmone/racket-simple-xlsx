#lang racket

(require rackunit/text-ui rackunit)

(require "../../main.rkt")

(require racket/runtime-path)
(define-runtime-path rows_file "_rows.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-invalid-rows"

    (check-exn
     exn:fail?
     (lambda ()
       (dynamic-wind
           (lambda () (void))
           (lambda ()
             (write-xlsx
              "rows.xlsx"
              (lambda ()
                (add-data-sheet
                 "Sheet1"
                 '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110))))))
           (lambda ()
             (when (file-exists? rows_file) (delete-file rows_file)))))))
   ))

(run-tests test-writer)
