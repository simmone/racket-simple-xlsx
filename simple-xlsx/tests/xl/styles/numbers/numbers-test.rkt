#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../style/style.rkt"
         "../../../../style/styles.rkt"
         "../../../../style/number-style.rkt"
         "../../../../style/assemble-styles.rkt"
         "../../../../style/set-styles.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/styles/numbers.rkt")

(provide (contract-out
          [set-number-styles (-> void?)]
          [check-number-styles (-> void?)]
          [numbers_file path-string?]
          [check-chaos-number-styles (-> void?)]
          [chaos_numbers_file path-string?]
          ))

(require racket/runtime-path)
(define-runtime-path numbers_file "numbers.xml")
(define-runtime-path chaos_numbers_file "chaos_numbers.xml")

(define (set-number-styles)
  (set-cell-range-number-style "A1" "0.00")
  (set-cell-range-number-style "B1" "#,###.00")
  (set-cell-range-date-style "C1" "yyyymmdd")
  (set-cell-range-date-style "D1" "yyyy-mm-dd"))

(define (check-number-styles)
  (check-equal? (length (STYLES-number_list (*STYLES*))) 4)
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 0) (NUMBER-STYLE "1" "0.00"))
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 1) (NUMBER-STYLE "2" "#,###.00"))
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 2) (NUMBER-STYLE "3" "yyyymmdd"))
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 3) (NUMBER-STYLE "4" "yyyy-mm-dd"))
  )

(define (check-chaos-number-styles)
  (check-equal? (length (STYLES-number_list (*STYLES*))) 4)
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 0) (NUMBER-STYLE "1" "0.00"))
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 1) (NUMBER-STYLE "3" "#,###.00"))
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 2) (NUMBER-STYLE "5" "yyyymmdd"))
  (check-equal? (list-ref (STYLES-number_list (*STYLES*)) 3) (NUMBER-STYLE "7" "yyyy-mm-dd"))
  )

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-numbers"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-number-styles)))

       (strip-styles)
       (assemble-styles)

       (call-with-input-file numbers_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content
             (to-numbers (STYLES-number_list (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-numbers"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-numbers
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string numbers_file)))
         '(
           "styleSheet.numFmts.numFmt.formatCode"
           "styleSheet.numFmts.numFmt.numFmtId"
           )
         ))

       (check-number-styles)
       )))

   (test-case
    "test-from-chaos-numbers"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-numbers
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string chaos_numbers_file)))
         '(
           "styleSheet.numFmts.numFmt.formatCode"
           "styleSheet.numFmts.numFmt.numFmtId"
           )
         ))

       (check-chaos-number-styles)
       )))

   ))

(run-tests test-styles)
