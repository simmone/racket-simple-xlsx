#lang racket

(require rackunit)
(require rackunit/text-ui)

(require racket/date)
(require racket/runtime-path)
(define-runtime-path write_back_file "write_back.xlsx")
(define-runtime-path line_feed_file "linefeed.xlsx")

(require "../../main.rkt")
(require rackunit "../../xlsx/sheet.rkt")

(define test-line-feed
  (test-suite
   "test-line-feed"

   (dynamic-wind
       (lambda () (void))
       (lambda ()
         (with-input-from-xlsx-file
          line_feed_file
          (lambda (xlsx)
            (let ([write_xlsx (from-read-to-write-xlsx xlsx)])
              (send write_xlsx set-data-sheet-col-width!
                    #:sheet_name "Sheet1"
                    #:col_range "B-C" #:width 50)
              (send write_xlsx set-data-sheet-col-width!
                    #:sheet_name "Sheet1"
                    #:col_range "A" #:width 100)
              (write-xlsx-file write_xlsx write_back_file))))

         (let ([rows (sheet-ref-rows write_back_file 0)])
           (check-equal? (car (list-ref rows 2)) "Line1\r\nLine2\r\nLine3")))
       (lambda ()
         (delete-file write_back_file)))))

(run-tests test-line-feed)

