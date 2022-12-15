#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../style/number-style.rkt")
(require "../../../style/sort-styles.rkt")
(require "../../../style/set-styles.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/styles/numbers.rkt")

(provide (contract-out
          [set-number-styles (-> void?)]
          [check-number-styles (-> void?)]
          [numbers_file path-string?]
          ))

(require racket/runtime-path)
(define-runtime-path numbers_file "numbers.xml")

(define (set-number-styles)
  (set-cell-range-number-style "A1" "0.00")
  (set-cell-range-number-style "B1" "#,###.00")
  (set-cell-range-date-style "C1" "yyyymmdd")
  (set-cell-range-date-style "D1" "yyyy-mm-dd")
  (set-cell-range-number-style "E1" "0.00%"))

(define (check-number-styles)
  (check-equal? (hash-count (*NUMBER_STYLE->INDEX_MAP*)) 5)
  (check-equal? (hash-count (*NUMBER_INDEX->STYLE_MAP*)) 5)

  (check-equal? (hash-ref (*NUMBER_STYLE->INDEX_MAP*) "#,###.00") 0)
  (check-equal? (hash-ref (*NUMBER_INDEX->STYLE_MAP*) 0) "#,###.00")

  (check-equal? (hash-ref (*NUMBER_STYLE->INDEX_MAP*) "0.00") 1)
  (check-equal? (hash-ref (*NUMBER_INDEX->STYLE_MAP*) 1) "0.00")

  (check-equal? (hash-ref (*NUMBER_STYLE->INDEX_MAP*) "0.00%") 2)
  (check-equal? (hash-ref (*NUMBER_INDEX->STYLE_MAP*) 2) "0.00%")

  (check-equal? (hash-ref (*NUMBER_STYLE->INDEX_MAP*) "yyyy-mm-dd") 3)
  (check-equal? (hash-ref (*NUMBER_INDEX->STYLE_MAP*) 3) "yyyy-mm-dd")

  (check-equal? (hash-ref (*NUMBER_STYLE->INDEX_MAP*) "yyyymmdd") 4)
  (check-equal? (hash-ref (*NUMBER_INDEX->STYLE_MAP*) 4) "yyyymmdd")
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

       (sort-styles)

       (call-with-input-file numbers_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-numbers
              (map
               (lambda (p)
                 (number-style-from-hash-code (car p)))
               (sort (hash->list (STYLES-number_style->index_map (*STYLES*))) < #:key cdr))))
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
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string numbers_file)))))

       (check-number-styles)
       ))
    )
   ))

(run-tests test-styles)
