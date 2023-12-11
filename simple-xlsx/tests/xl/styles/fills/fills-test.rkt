#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../style/style.rkt")
(require "../../../../style/styles.rkt")
(require "../../../../style/fill-style.rkt")
(require "../../../../style/assemble-styles.rkt")
(require "../../../../style/set-styles.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/styles/fills.rkt")

(require racket/runtime-path)
(define-runtime-path fills_file "fills.xml")
(define-runtime-path empty_fills_file "empty_fills.xml")
(define-runtime-path theme_fills_file "theme_fills.xml")

(provide (contract-out
          [set-fill-styles (-> void?)]
          [check-fill-styles (-> void?)]
          [fills_file path-string?]
          ))

(define (set-fill-styles)
  (set-cell-range-fill-style "A1" "FF0000" "gray125")
  (set-cell-range-fill-style "B1" "000000" "solid")
  (set-cell-range-fill-style "C1" "FFFF00" "solid")
  (set-cell-range-fill-style "D1" "FFFFFF" "solid"))

(define (check-fill-styles)
  (check-equal? (length (STYLES-fill_list (*STYLES*))) 6)
  (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 0) (FILL-STYLE "FFFFFF" "none"))
  (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 1) (FILL-STYLE "FFFFFF" "gray125"))
  (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 2) (FILL-STYLE "FF0000" "gray125"))
  (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 3) (FILL-STYLE "000000" "solid"))
  (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 4) (FILL-STYLE "FFFF00" "solid"))
  (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 5) (FILL-STYLE "FFFFFF" "solid"))
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

       (strip-styles)
       (assemble-styles)

       (call-with-input-file fills_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-fills (STYLES-fill_list (*STYLES*))))
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

   (test-case
    "test-to-empty-fills"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (strip-styles)
       (assemble-styles)

       (call-with-input-file empty_fills_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-fills (STYLES-fill_list (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-theme-fills"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-fills
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string theme_fills_file)))))

       (check-equal? (length (STYLES-fill_list (*STYLES*))) 1)
       (check-equal? (list-ref (STYLES-fill_list (*STYLES*)) 0) (FILL-STYLE "FFFFFF" "none"))
       )))

   ))

(run-tests test-styles)
