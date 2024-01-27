#lang racket

(require rackunit/text-ui rackunit)

(require "../../main.rkt")

(require racket/runtime-path)
(define-runtime-path mixed_file "_mixed.xlsx")
(define-runtime-path mixed_read_and_write_file "_mixed_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-mixed"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           mixed_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              '(("month1" 3.4 "month3" "month4" 8000.0) ("201601" 100 110 "1110" 6.9)))
             ))

          (read-xlsx
           mixed_file
           (lambda ()
             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (get-row 1) '("month1" 3.4 "month3" "month4" 8000.0))
                (check-equal? (get-row 2) '("201601" 100 110 "1110" 6.9))))
             ))

          (read-and-write-xlsx
           mixed_file
           mixed_read_and_write_file
           (lambda ()
             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (get-row 1) '("month1" 3.4 "month3" "month4" 8000.0))
                (check-equal? (get-row 2) '("201601" 100 110 "1110" 6.9))

                (set-row! 2 '("201602" 102 111 "1111" 6.8))
                (check-equal? (get-row 2) '("201602" 102 111 "1111" 6.8))
                ))
             ))
          )
        (lambda ()
          ;(void)
          (delete-file mixed_file)
          (delete-file mixed_read_and_write_file)
          )))
   ))

(run-tests test-writer)
