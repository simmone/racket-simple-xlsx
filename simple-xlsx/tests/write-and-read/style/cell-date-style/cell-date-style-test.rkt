#lang racket

(require rackunit/text-ui rackunit)

(require racket/date)

(require "../../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_date_file "cell_date.xlsx")
(define-runtime-path cell_date_read_and_write_file "cell_date_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-cell-range-date"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           cell_date_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              (list
               (list
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
               )))

             (with-sheet
              (lambda ()
                (set-cell-range-date-style "A1" "yyyy-mm-dd")
                (set-cell-range-date-style "B1" "yyyy/mm/dd")
                (set-cell-range-date-style "C1" "yyyymmdd")

                (set-col-range-width "A-C" 20)
                (set-row-range-height "1-3" 20)
                ))))

          (read-and-write-xlsx
           cell_date_file
           cell_date_read_and_write_file
           (lambda ()
             (void)))
          )
        (lambda ()
;;          (void)
          (delete-file cell_date_file)
          (delete-file cell_date_read_and_write_file)
          )))

   ))

(run-tests test-writer)
