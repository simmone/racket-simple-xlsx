#lang racket

(provide (contract-out
          [add-cell-date-sheet (-> void?)]
          [set-more-cell-dates (-> void?)]
          ))

(require rackunit/text-ui rackunit)

(require racket/date)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path cell_date_file "_cell_date.xlsx")
(define-runtime-path cell_date_read_and_write_file "_cell_date_read_and_write.xlsx")

(define (add-cell-date-sheet)
  (add-data-sheet
   "cell date"
   (list
    (list
     (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
     (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
     (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
     (seconds->date (find-seconds 0 0 0 9 11 2023 #f))
     )))

  (with-sheet-name
   "cell date"
   (lambda ()
     (set-cell-range-date-style "A1" "yyyy-mm-dd")
     (set-cell-range-date-style "B1" "yyyy/mm/dd")
     (set-cell-range-date-style "C1" "yyyymmdd")

     (set-col-range-width "A-C" 20)
     (set-row-range-height "1" 20)
     )))

(define (set-more-cell-dates)
  (with-sheet-name
   "cell date"
   (lambda ()
     (set-cell-range-date-style "C1" "yyyy!+mm!+dd")
     (set-cell-range-date-style "D1" "yyyy年mm月dd日")
     (set-col-range-width "D" 40)
     )))

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
             (add-cell-date-sheet)))

          (read-and-write-xlsx
           cell_date_file
           cell_date_read_and_write_file
           (lambda ()
             (set-more-cell-dates)))
          )
        (lambda ()
          ;(void)
          (delete-file cell_date_file)
          (delete-file cell_date_read_and_write_file)
          )))
   ))

(run-tests test-writer)
