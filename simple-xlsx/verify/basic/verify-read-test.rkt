#lang racket

(require rackunit/text-ui
         rackunit
         "../../main.rkt"
         "../../style/styles.rkt"
         racket/runtime-path)

(define-runtime-path basic_write_google_file "basic_write_google.xlsx")
(define-runtime-path basic_write_ms_file "basic_write_ms.xlsx")
(define-runtime-path basic_write_wps_file "basic_write_wps.xlsx")
(define-runtime-path basic_write_libre_file "basic_write_libre.xlsx")

(define test-basic
  (test-suite
   "test-basic"

   (test-case
    "test-google"

    (read-xlsx
     basic_write_google_file
     (lambda ()
       (check-equal? (get-sheet-name-list) '("Shoot1" "Shoot2"))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))))

       (with-sheet-name
        "Shoot2"
        (lambda ()
          (check-equal? (get-row 1) '(201601.0 100.0 110.0 1110.0 6.9))))))
    )

   (test-case
    "test-ms"

    (read-xlsx
     basic_write_ms_file
     (lambda ()
       (check-equal? (get-sheet-name-list) '("Shoot1" "Shoot2"))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))))

       (with-sheet-name
        "Shoot2"
        (lambda ()
          (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))))
    )

   (test-case
    "test-wps"

    (read-xlsx
     basic_write_wps_file
     (lambda ()
       (check-equal? (get-sheet-name-list) '("Shoot1" "Shoot2"))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))))

       (with-sheet-name
        "Shoot2"
        (lambda ()
          (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))))
    )

   (test-case
    "test-libre"

    (read-xlsx
     basic_write_libre_file
     (lambda ()
       (check-equal? (get-sheet-name-list) '("Shoot1" "Shoot2"))

       (with-sheet-ref
        0
        (lambda ()
          (check-equal? (get-row 1) '("month1" "month2" "month3" "month4" "real"))))

       (with-sheet-name
        "Shoot2"
        (lambda ()

          (check-equal? (get-row 1) '(201601 100 110 1110 6.9))))))
    )


   ))

(run-tests test-basic)
