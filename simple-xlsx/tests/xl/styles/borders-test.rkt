#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../style/border-style.rkt")
(require "../../../style/sort-styles.rkt")
(require "../../../style/set-styles.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/styles/borders.rkt")

(require racket/runtime-path)
(define-runtime-path borders_file "borders.xml")
(define-runtime-path side_borders_file "side_borders.xml")

(provide (contract-out
          [set-border-styles (-> void?)]
          [check-border-styles (-> void?)]
          [borders_file path-string?]
          ))

(define (set-border-styles)
  (set-cell-range-border-style "B1" "all" "000000" "dashed")
  (set-cell-range-border-style "A1" "all" "ff0000" "dashed")
  (set-cell-range-border-style "C1" "bottom" "0000ff" "thick")
  (set-cell-range-border-style "D1" "top" "000000" "thin"))

(define (check-border-styles)
  (check-equal? (hash-count (*BORDER_STYLE->INDEX_MAP*)) 4)
  (check-equal? (hash-count (*BORDER_INDEX->STYLE_MAP*)) 4)

  (check-equal? (hash-ref (*BORDER_STYLE->INDEX_MAP*) "000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed") 0)
  (check-equal? (hash-ref (*BORDER_INDEX->STYLE_MAP*) 0) "000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed")

  (check-equal? (hash-ref (*BORDER_STYLE->INDEX_MAP*) "000000<p>thin<p><p><p><p><p><p>") 1)
  (check-equal? (hash-ref (*BORDER_INDEX->STYLE_MAP*) 1) "000000<p>thin<p><p><p><p><p><p>")

  (check-equal? (hash-ref (*BORDER_STYLE->INDEX_MAP*) "<p><p>0000FF<p>thick<p><p><p><p>") 2)
  (check-equal? (hash-ref (*BORDER_INDEX->STYLE_MAP*) 2) "<p><p>0000FF<p>thick<p><p><p><p>")

  (check-equal? (hash-ref (*BORDER_STYLE->INDEX_MAP*) "FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed") 3)
  (check-equal? (hash-ref (*BORDER_INDEX->STYLE_MAP*) 3) "FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed")
  )

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-borders"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-border-styles)))

       (sort-styles)

       (call-with-input-file borders_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-borders
              (map
               (lambda (p)
                 (border-style-from-hash-code (car p)))
               (sort (hash->list (STYLES-border_style->index_map (*STYLES*))) < #:key cdr))))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-borders"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-borders
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string borders_file)))))

       (check-border-styles)
       )))

   (test-case
    "test-side-borders"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(
                         ("month1" "month2" "month3")
                         ("month1" "month2" "month3")
                         ("month1" "month2" "month3")))

       (with-sheet
        (lambda ()
          (set-cell-range-border-style "A1-C3" "side" "000000" "thick")))

       (sort-styles)

       (check-equal? (hash-count (STYLES-border_style->index_map (*STYLES*))) 8)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "000000<p>thick<p><p><p>000000<p>thick<p><p>") 0)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "000000<p>thick<p><p><p><p><p>000000<p>thick") 1)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "000000<p>thick<p><p><p><p><p><p>") 2)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "<p><p>000000<p>thick<p>000000<p>thick<p><p>") 3)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "<p><p>000000<p>thick<p><p><p>000000<p>thick") 4)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "<p><p>000000<p>thick<p><p><p><p>") 5)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "<p><p><p><p>000000<p>thick<p><p>") 6)
       (check-equal? (hash-ref (STYLES-border_style->index_map (*STYLES*)) "<p><p><p><p><p><p>000000<p>thick") 7)

       (call-with-input-file side_borders_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-borders
              (map
               (lambda (p)
                 (border-style-from-hash-code (car p)))
               (sort (hash->list (STYLES-border_style->index_map (*STYLES*))) < #:key cdr))))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-styles)
