#lang racket

(require rackunit/text-ui
         rackunit
         "../../main.rkt"
         racket/runtime-path)

(define-runtime-path fill_rows1_file "_fill_rows1.xlsx")
(define-runtime-path fill_rows2_file "_fill_rows2.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-fill-rows1"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           fill_rows1_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("month1" "month2" "month3" "month4" "real")
                (201601 100 110 1110)))))

          (read-xlsx
           fill_rows1_file
           (lambda ()
             (with-sheet
              (lambda ()
                (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))
                (check-equal? (get-row 2) '(201601 100 110 1110 ""))))))
          )
        (lambda ()
          (delete-file fill_rows1_file))))

   (test-case
    "test-fill-rows2"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           fill_rows2_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(
                ("month1" "month2" "month3" "month4" "real")
                (201601 100 110)
                ("a")
                )
              #:fill? 0
              )))

          (read-xlsx
           fill_rows2_file
           (lambda ()
             (with-sheet
              (lambda ()
                (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))
                (check-equal? (get-row 2) '(201601 100 110 0 0))
                (check-equal? (get-row 3) '("a" 0 0 0 0))
                ))))
          )
        (lambda ()
          (delete-file fill_rows2_file)
          )))
   ))

(run-tests test-writer)
