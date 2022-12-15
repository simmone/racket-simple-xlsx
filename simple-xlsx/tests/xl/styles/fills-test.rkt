#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../style/fill-style.rkt")
(require "../../../style/sort-styles.rkt")
(require "../../../style/set-styles.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/styles/fills.rkt")

(require racket/runtime-path)
(define-runtime-path fills_file "fills.xml")

(provide (contract-out
          [set-fill-styles (-> void?)]
          [check-fill-styles (-> void?)]
          [fills_file path-string?]
          ))

(define (set-fill-styles)
  (set-cell-range-fill-style "B1" "000000" "solid")
  (set-cell-range-fill-style "A1" "ff0000" "gray125")
  (set-cell-range-fill-style "C1" "ffff00" "solid")
  (set-cell-range-fill-style "D1" "ffffff" "solid"))

(define (check-fill-styles)
  (check-equal? (hash-count (*FILL_STYLE->INDEX_MAP*)) 4)
  (check-equal? (hash-count (*FILL_INDEX->STYLE_MAP*)) 4)

  (check-equal? (hash-ref (*FILL_INDEX->STYLE_MAP*) 0) "000000<p>solid")
  (check-equal? (hash-ref (*FILL_STYLE->INDEX_MAP*) "000000<p>solid") 0)

  (check-equal? (hash-ref (*FILL_INDEX->STYLE_MAP*) 1) "FF0000<p>gray125")
  (check-equal? (hash-ref (*FILL_STYLE->INDEX_MAP*) "FF0000<p>gray125") 1)

  (check-equal? (hash-ref (*FILL_INDEX->STYLE_MAP*) 2) "FFFF00<p>solid")
  (check-equal? (hash-ref (*FILL_STYLE->INDEX_MAP*) "FFFF00<p>solid") 2)

  (check-equal? (hash-ref (*FILL_INDEX->STYLE_MAP*) 3) "FFFFFF<p>solid")
  (check-equal? (hash-ref (*FILL_STYLE->INDEX_MAP*) "FFFFFF<p>solid") 3)
  )

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-fills"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-fill-styles)))

       (sort-styles)

       (call-with-input-file fills_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-fills
              (map
               (lambda (p)
                 (fill-style-from-hash-code (car p)))
               (sort (hash->list (STYLES-fill_style->index_map (*STYLES*))) < #:key cdr))))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-fills"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-fills
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fills_file)))))

       (check-fill-styles)
       )))
   ))

(run-tests test-styles)
