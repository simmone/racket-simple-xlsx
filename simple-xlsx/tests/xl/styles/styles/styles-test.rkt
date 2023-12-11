#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../lib/lib.rkt")
(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../style/style.rkt")
(require "../../../../style/styles.rkt")
(require "../../../../style/number-style.rkt")
(require "../../../../style/font-style.rkt")
(require "../../../../style/fill-style.rkt")
(require "../../../../style/border-style.rkt")
(require "../../../../style/assemble-styles.rkt")
(require "../../../../style/set-styles.rkt")
(require "../../../../xl/styles/styles.rkt")

(require racket/runtime-path)
(define-runtime-path styles_file "styles.xml")
(define-runtime-path chaos_number_styles_file "chaos_number_styles.xml")
(define-runtime-path complex_styles_file "complex_styles_test.xml")

(require "../borders/borders-test.rkt")
(require "../fills/fills-test.rkt")
(require "../fonts/fonts-test.rkt")
(require "../numbers/numbers-test.rkt")
(require "../cellXfs/cellXfs-alignment-test.rkt")

(require "../../../../xl/styles/borders.rkt")
(require "../../../../xl/styles/fills.rkt")
(require "../../../../xl/styles/numbers.rkt")
(require "../../../../xl/styles/fonts.rkt")
(require "../../../../xl/styles/cellXfs.rkt")

(provide (contract-out
          [check-styles (-> void?)]
          ))

(define (check-styles)
  (check-border-styles)

  (check-fill-styles)

  (check-font-styles)

  (check-number-styles)

  (check-cellXfses))

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-cellXfses)))

       (strip-styles)
       (assemble-styles)

       (call-with-input-file styles_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content (to-styles))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (check-equal? (length (STYLES-styles (*STYLES*))) 0)
       (from-borders
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string borders_file)))))
       (check-border-styles)

       (from-fonts
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fonts_file)))))
       (check-font-styles)

       (from-numbers
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string numbers_file)))))
       (check-number-styles)

       (from-fills
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fills_file)))))
       (check-fill-styles)

       (from-cellXfs
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string cellXfs_file)))))
       (check-cellXfses)

       (from-styles styles_file)

       (check-styles)
       )))

   (test-case
    "test-from-complex-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-styles complex_styles_file)
       )))

   (test-case
    "test-from-chaos-number-styles"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-styles chaos_number_styles_file)
       )))
   ))

(run-tests test-styles)
