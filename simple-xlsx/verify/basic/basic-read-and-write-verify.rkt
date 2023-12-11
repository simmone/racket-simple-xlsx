#lang racket

(require rackunit/text-ui rackunit)

(require "../../main.rkt")

(require racket/runtime-path)
(define-runtime-path basic_write_file "basic_write_app.xlsx")
(define-runtime-path basic_read_and_write_file "basic_read_and_write_app.xlsx")

(define test-basic
  (test-suite
   "test-basic"

   (test-case
    "test-basic"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           basic_write_file
           (lambda ()
             (add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4" "real")))

             (add-data-sheet "Sheet2" '((201601 100 110 1110 6.9)))))

          (read-xlsx
           basic_write_file
           (lambda ()
             (check-equal? (get-sheet-name-list) '("Sheet1" "Sheet2"))

             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))))

             (with-sheet-name
              "Sheet2"
              (lambda ()
                (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))))

          (read-and-write-xlsx
           basic_write_file
           basic_read_and_write_file
           (lambda ()
             (check-equal? (get-sheet-name-list) '("Sheet1" "Sheet2"))

             (with-sheet-*name*
              "Sheet"
              (lambda ()
                (set-cell! "B1" "John")

                (check-equal? (get-row 1) '("month1" "John" "month3" "month4" "real"))))

             (with-sheet-ref
              1
              (lambda ()
                (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))
             ))

          ;; issue: read multiple sheet count
          (read-xlsx
           basic_write_file
           (lambda ()
             (check-equal?
               (let loop ([sheet_names (get-sheet-name-list)]
                          [rows_sum 0])
                 (if (not (null? sheet_names))
                     (loop (cdr sheet_names)
                           (+ rows_sum (get-sheet-name-rows-count (car sheet_names))))
                     rows_sum))
               2)))
          )
        (lambda ()
          ;(void)
          (delete-file basic_write_file)
          (delete-file basic_read_and_write_file)
          )))
   ))

(run-tests test-basic)
