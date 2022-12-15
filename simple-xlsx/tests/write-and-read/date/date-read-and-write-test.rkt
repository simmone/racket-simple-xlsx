#lang racket

(require rackunit/text-ui rackunit)

(require racket/date)

(require "../../../main.rkt")

(require racket/runtime-path)
(define-runtime-path date_write_file "date_write.xlsx")
(define-runtime-path date_read_and_write_file "date_read_and_write.xlsx")

(define test-writer
  (test-suite
   "test-writer"

   (test-case
    "test-date"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (write-xlsx
           date_write_file
           (lambda ()
             (add-data-sheet
              "Sheet1"
              (list
               (list
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
                (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))
               ))

             (with-sheet
              (lambda ()
                (set-row-range-date-style "1" "yyyy-mm-dd")
                (set-col-range-width "1-5" 30)
                ))
             ))

          (read-xlsx
           date_write_file
           (lambda ()
             (check-equal? (get-sheet-name-list) '("Sheet1"))

             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (get-row 1) '(43360 43360 43360 43360 43360))))))

          (read-and-write-xlsx
           date_write_file
           date_read_and_write_file
           (lambda ()
             (check-equal? (get-sheet-name-list) '("Sheet1"))

             (with-sheet-ref
              0
              (lambda ()
                (check-equal? (get-row 1) '(43360 43360 43360 43360 43360))

                (set-row-range-date-style "1" "yyyy/mm/dd")
                (set-row-range-alignment-style "1" "center" "center")
                ))
             ))
          )
        (lambda ()
;;          (void)
          (delete-file date_write_file)
          (delete-file date_read_and_write_file)
          )
      ))
   ))

(run-tests test-writer)
