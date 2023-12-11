#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../style/style.rkt")
(require "../../../../style/styles.rkt")
(require "../../../../style/font-style.rkt")
(require "../../../../style/assemble-styles.rkt")
(require "../../../../style/set-styles.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/styles/fonts.rkt")

(require racket/runtime-path)
(define-runtime-path fonts_file "fonts.xml")

(provide (contract-out
          [set-font-styles (-> void?)]
          [check-font-styles (-> void?)]
          [fonts_file path-string?]
          ))

(define (set-font-styles)
  (set-cell-range-font-style "A1" 10 "Arial" "0000FF")
  (set-cell-range-font-style "B1" 15 "Arial" "0000FF")
  (set-cell-range-font-style "C1" 15 "宋体" "0000FF")
  (set-cell-range-font-style "D1" 15 "宋体" "FF0000"))

(define (check-font-styles)
  (check-equal? (length (STYLES-font_list (*STYLES*)))  5)
  (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 0) (FONT-STYLE 10 "Arial" "000000"))
  (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 1) (FONT-STYLE 10 "Arial" "0000FF"))
  (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 2) (FONT-STYLE 15 "Arial" "0000FF"))
  (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 3) (FONT-STYLE 15 "宋体" "0000FF"))
  (check-equal? (list-ref (STYLES-font_list (*STYLES*)) 4) (FONT-STYLE 15 "宋体" "FF0000"))
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

       (strip-styles)
       (assemble-styles)

       (call-with-input-file fonts_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-fonts (STYLES-font_list (*STYLES*))))
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
