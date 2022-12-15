#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")
(require "../../../style/style.rkt")
(require "../../../style/sort-styles.rkt")
(require "../../../style/set-styles.rkt")
(require "../../../lib/lib.rkt")

(require"../../../xl/styles/cellXfs.rkt")

(require racket/runtime-path)
(define-runtime-path cellXfs_file "cellXfs.xml")

(require "borders-test.rkt")
(require "fills-test.rkt")
(require "numbers-test.rkt")
(require "fonts-test.rkt")

(require "../../../xl/styles/borders.rkt")
(require "../../../xl/styles/fills.rkt")
(require "../../../xl/styles/numbers.rkt")
(require "../../../xl/styles/fonts.rkt")

(provide (contract-out
          [set-cellXfses (-> void?)]
          [check-cellXfses (-> void?)]
          [cellXfs_file path-string?]
          ))

(define (set-cellXfses)
  (set-border-styles)
  (set-fill-styles)
  (set-number-styles)
  (set-font-styles)

  (set-cell-range-border-style "A1" "all" "000000" "dashed")
  (set-cell-range-border-style "A2" "all" "FF0000" "dashed")
  (set-cell-range-alignment-style "A1" "left" "center")
  (set-cell-range-number-style "B1" "#,###.00")
  (set-cell-range-alignment-style "B1" "right" "top")
  (set-cell-range-fill-style "C1" "FFFF00" "solid")
  (set-cell-range-alignment-style "C1" "center" "bottom")
  (set-cell-range-font-style "D1" 15 "宋体" "FF0000")

  (set-cell-range-border-style "E1" "all" "FF0000" "dashed")
  (set-cell-range-number-style "E1" "0.00%")
  (set-cell-range-fill-style "E1" "FFFF00" "solid")
  (set-cell-range-font-style "E1" 15 "Arial" "0000FF")
  (set-cell-range-alignment-style "E1" "center" "top"))

(define (check-cellXfses)
  (check-equal? (hash-count (*STYLE->INDEX_MAP*)) 6)
  (check-equal? (hash-count (*INDEX->STYLE_MAP*)) 6)

  (check-equal? (hash-ref (*STYLE->INDEX_MAP*)
                "<p><p>0000FF<p>thick<p><p><p><p><s>15<p>宋体<p>0000FF<s>center<p>bottom<s>yyyymmdd<s>FFFF00<p>solid")
                1)
  (check-equal? (hash-ref (*INDEX->STYLE_MAP*) 1)
                "<p><p>0000FF<p>thick<p><p><p><p><s>15<p>宋体<p>0000FF<s>center<p>bottom<s>yyyymmdd<s>FFFF00<p>solid")

  (check-equal? (hash-ref (*STYLE->INDEX_MAP*)
                "000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>center<s>0.00<s>FF0000<p>gray125")
                2)
  (check-equal? (hash-ref (*INDEX->STYLE_MAP*) 2)
                "000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>center<s>0.00<s>FF0000<p>gray125")

  (check-equal? (hash-ref (*STYLE->INDEX_MAP*)
                "000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<s>15<p>Arial<p>0000FF<s>right<p>top<s>#,###.00<s>000000<p>solid")
                3)
  (check-equal? (hash-ref (*INDEX->STYLE_MAP*) 3)
                "000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<p>000000<p>dashed<s>15<p>Arial<p>0000FF<s>right<p>top<s>#,###.00<s>000000<p>solid")

  (check-equal? (hash-ref (*STYLE->INDEX_MAP*)
                "000000<p>thin<p><p><p><p><p><p><s>15<p>宋体<p>FF0000<s>left<p>center<s>yyyy-mm-dd<s>FFFFFF<p>solid")
                4)
  (check-equal? (hash-ref (*INDEX->STYLE_MAP*) 4)
                "000000<p>thin<p><p><p><p><p><p><s>15<p>宋体<p>FF0000<s>left<p>center<s>yyyy-mm-dd<s>FFFFFF<p>solid")

  (check-equal? (hash-ref (*STYLE->INDEX_MAP*)
                "FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<s><s>left<p>center<s><s>")
                5)
  (check-equal? (hash-ref (*INDEX->STYLE_MAP*) 5)
                "FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<s><s>left<p>center<s><s>")

  (check-equal? (hash-ref (*STYLE->INDEX_MAP*)
                "FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<s>15<p>Arial<p>0000FF<s>center<p>top<s>0.00%<s>FFFF00<p>solid")
                6)
  (check-equal? (hash-ref (*INDEX->STYLE_MAP*) 6)
                "FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<p>FF0000<p>dashed<s>15<p>Arial<p>0000FF<s>center<p>top<s>0.00%<s>FFFF00<p>solid"))

(define test-styles
(test-suite
 "test-styles"

 (test-case
  "test-to-cellXfs"

  (with-xlsx
   (lambda ()
     (add-data-sheet "Sheet1"
                     '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
     (add-data-sheet "Sheet2" '((1)))
     (add-data-sheet "Sheet3" '((1)))

     (with-sheet (lambda () (set-cellXfses)))

     (sort-styles)

     (call-with-input-file cellXfs_file
       (lambda (expected)
         (call-with-input-string
          (lists->xml_content
           (to-cellXfs
            (map
             (lambda (p)
               (style-from-hash-code (car p)))
             (sort (hash->list (STYLES-style->index_map (*STYLES*))) < #:key cdr))))
          (lambda (actual)
            (check-lines? expected actual))))))))

 (test-case
  "test-from-cellXfs"

  (with-xlsx
   (lambda ()
     (add-data-sheet "Sheet1"
                     '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
     (add-data-sheet "Sheet2" '((1)))
     (add-data-sheet "Sheet3" '((1)))

     (check-equal? (hash-count (*STYLE->INDEX_MAP*)) 0)
     (check-equal? (hash-count (*INDEX->STYLE_MAP*)) 0)

     (from-borders
      (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string borders_file)))))

     (from-fonts
      (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fonts_file)))))

     (from-numbers
      (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string numbers_file)))))

     (from-fills
      (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fills_file)))))

     (from-cellXfs
      (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string cellXfs_file)))))

     (check-cellXfses)
     )))

 ))

(run-tests test-styles)
