#lang racket

(require rackunit/text-ui rackunit)

(require "../../sheet/sheet.rkt")

(require "../../xlsx/xlsx.rkt")

(define test-xlsx
  (test-suite
   "test-xlsx"

   (test-case
    "test-xlsx"

    (with-xlsx
     (lambda ()
     (check-equal? (length (XLSX-sheet_list (*XLSX*))) 0)

     (add-data-sheet "测试1" '((1)))

     (check-equal? (length (XLSX-sheet_list (*XLSX*))) 1)

     (check-exn exn:fail? (lambda () (add-data-sheet "测试1" '())))
     )))

   (test-case
    "test-get-sheet-name-list"

    (with-xlsx
     (lambda ()
       (add-data-sheet "sheet1" '((1)))

       (add-data-sheet "sheet2" '((1 2)))

       (add-data-sheet "sheet3" '(
                                  (1 2 3)
                                  (1 2 3)
                                  ))

       (add-data-sheet "sheet4" '(
                                  (1 2)
                                  (1 2)
                                  (1 2)
                                  ))

       (check-equal? (get-sheet-name-list) '("sheet1" "sheet2" "sheet3" "sheet4"))
       )))

   ))

(run-tests test-xlsx)
