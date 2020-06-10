#lang racket

(require rackunit)
(require rackunit/text-ui)

(require racket/date)
(require racket/runtime-path)
(define-runtime-path test_file "test.xlsx")
(define-runtime-path write_back_file "write_back.xlsx")

(require "../../main.rkt")

(define test-from-read-to-write-xlsx
  (test-suite
   "test-test1"

   (dynamic-wind
       (lambda () (void))
       (lambda ()
         (let ([xlsx (new xlsx%)]
               [sheet1_data (list
                             (list "month/brand" "201601" "201602" "201603" "201604" "201605")
                             (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)))
                             (list "Puma" 200 400 300 139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)))
                             (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018))))]
               [sheet2_data (list
                             (list "month/brand" "201606" "201607" "201608" "201609" "201610")
                             (list "CAT" 100 300 200 0.6934 (seconds->date (find-seconds 0 0 0 17 9 2018)))
                             (list "Puma" 200 400 300 139999.89223 (seconds->date (find-seconds 0 0 0 18 9 2018)))
                             (list "Asics" 300 500 400 23.34 (seconds->date (find-seconds 0 0 0 19 9 2018))))]
               )

           (send xlsx add-data-sheet #:sheet_name "DataSheet1" #:sheet_data sheet1_data)
           (send xlsx add-data-sheet #:sheet_name "DataSheet2" #:sheet_data sheet2_data)

           (write-xlsx-file xlsx test_file)

           (with-input-from-xlsx-file
            test_file
            (lambda (xlsx)
              (let ([write_xlsx (from-read-to-write-xlsx xlsx)])
                (send write_xlsx set-data-sheet-col-width!
                      #:sheet_name "DataSheet1"
                      #:col_range "A-F" #:width 20)
                (write-xlsx-file write_xlsx write_back_file)))))

         (with-input-from-xlsx-file
          test_file
          (lambda (xlsx)
            (test-case 
             "test-get-sheets"
             
             (check-equal? (get-sheet-names xlsx) '("DataSheet1" "DataSheet2")))
            
            (test-case
             "test-get-sheet-data"
             
             (load-sheet "DataSheet1" xlsx)
             (check-equal? (get-cell-value "B1" xlsx) "201601")

             (load-sheet "DataSheet2" xlsx)
             (check-equal? (get-cell-value "B1" xlsx) "201606")
             ))))
       (lambda ()
         (delete-file test_file)
         (delete-file write_back_file)))))

(run-tests test-from-read-to-write-xlsx)

