#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../style/font-style.rkt")
(require "../../../style/sort-styles.rkt")
(require "../../../style/set-styles.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/styles/fonts.rkt")

(require racket/runtime-path)
(define-runtime-path fonts_file "fonts.xml")

(provide (contract-out
          [set-font-styles (-> void?)]
          [check-font-styles (-> void?)]
          [fonts_file path-string?]
          ))

(define (set-font-styles)
  (set-cell-range-font-style "A1" 10 "Arial" "0000ff")
  (set-cell-range-font-style "B1" 15 "Arial" "0000ff")
  (set-cell-range-font-style "C1" 15 "宋体" "0000ff")
  (set-cell-range-font-style "D1" 15 "宋体" "ff0000"))

(define (check-font-styles)
  (check-equal? (hash-count (*FONT_STYLE->INDEX_MAP*)) 4)
  (check-equal? (hash-count (*FONT_INDEX->STYLE_MAP*)) 4)

  (check-equal? (hash-ref (*FONT_STYLE->INDEX_MAP*) "10<p>Arial<p>0000FF") 0)
  (check-equal? (hash-ref (*FONT_INDEX->STYLE_MAP*) 0) "10<p>Arial<p>0000FF")

  (check-equal? (hash-ref (*FONT_STYLE->INDEX_MAP*) "15<p>Arial<p>0000FF") 1)
  (check-equal? (hash-ref (*FONT_INDEX->STYLE_MAP*) 1) "15<p>Arial<p>0000FF")

  (check-equal? (hash-ref (*FONT_STYLE->INDEX_MAP*) "15<p>宋体<p>0000FF") 2)
  (check-equal? (hash-ref (*FONT_INDEX->STYLE_MAP*) 2) "15<p>宋体<p>0000FF")

  (check-equal? (hash-ref (*FONT_STYLE->INDEX_MAP*) "15<p>宋体<p>FF0000") 3)
  (check-equal? (hash-ref (*FONT_INDEX->STYLE_MAP*) 3) "15<p>宋体<p>FF0000")
  )

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-fonts"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-font-styles)))

       (sort-styles)

       (call-with-input-file fonts_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-fonts
              (map
               (lambda (p)
                 (font-style-from-hash-code (car p)))
               (sort (hash->list (STYLES-font_style->index_map (*STYLES*))) < #:key cdr))))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-fonts"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-fonts
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fonts_file)))))

       (check-font-styles)
       )))
   ))

(run-tests test-styles)
