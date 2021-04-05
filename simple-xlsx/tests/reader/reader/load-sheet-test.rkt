#lang racket

(require rackunit/text-ui)

(require rackunit "../../../reader.rkt")

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")

(require racket/runtime-path)
(define-runtime-path test1_file "test1.xlsx")

(define test-load-sheet
  (test-suite
   "test-load-sheet"
   
   (test-case
    "test-load-sheet"

    (with-input-from-xlsx-file
     test1_file
     (lambda ()
       (load-sheet "DataSheet")
       
       (check-equal? (length (XLSX-sheet_list (*CURRENT_XLSX*))) 1)
      )))

   (test-case
    "test-load-sheet"

    (with-input-from-xlsx-file
     test1_file
     (lambda ()
       (load-sheet "DataSheet")
       
       (check-equal? (length (XLSX-sheet_list (*CURRENT_XLSX*))) 1)
      )))

   (test-case
    "test-load-sheets"

    (with-input-from-xlsx-file
     test1_file
     (lambda ()
       (load-sheets)
       
       (check-equal? (length (XLSX-sheet_list (*CURRENT_XLSX*))) 10)
      )))

   (test-case
    "test-load-sheet-user-proc"

    (with-input-from-xlsx-file
     test1_file
     (lambda ()
       (load-sheet 
        "DataSheet"
        (lambda ()
          (check-equal? (sheet-dimension) '(4 . 6))

          (check-equal? (get-cell-value "A1") "month/brand")
          (check-equal? (get-cell-value "a1") "month/brand")
          (check-equal? (get-cell-value "a") "")
          
          (check-equal? (get-rows)
                        (list
                         (list "month/brand" "201601" "201602" "201603" "201604" "201605")
                         (list "CAT"   100 300 200 0.6934 43360)
                         (list "Puma"  200 400 300 139999.89223 43361)
                         (list "Asics" 300 500 400 23.34 43362)
                         ))

          (check-equal? (get-row 0) (list "month/brand" "201601" "201602" "201603" "201604" "201605"))
          (check-equal? (get-row 3) (list "Asics" 300 500 400 23.34 43362))
          )))))


    ))

(run-tests test-load-sheet)
