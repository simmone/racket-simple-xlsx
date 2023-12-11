#lang racket

(require simple-xml)

(require rackunit/text-ui rackunit)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../sheet/sheet.rkt")
(require "../../../../style/style.rkt")
(require "../../../../style/alignment-style.rkt")
(require "../../../../style/border-style.rkt")
(require "../../../../style/font-style.rkt")
(require "../../../../style/fill-style.rkt")
(require "../../../../style/number-style.rkt")
(require "../../../../style/styles.rkt")
(require "../../../../style/assemble-styles.rkt")
(require "../../../../style/set-styles.rkt")
(require "../../../../lib/lib.rkt")

(require"../../../../xl/styles/cellXfs.rkt")

(require racket/runtime-path)
(define-runtime-path cellXfs_file "cellXfs.xml")
(define-runtime-path cellXfs_apply_true_file "cellXfs_apply_true.xml")
(define-runtime-path cellXfs_alignment_empty_file "cellXfs_alignment_empty.xml")
(define-runtime-path empty_cellXfs_file "empty_cellXfs.xml")
(define-runtime-path chaos_number_cellXfs_file "chaos_number_cellXfs.xml")
(define-runtime-path number_not_numfmts_cellXfs_file "number_not_numfmts_cellXfs.xml")

(require "../borders/borders-test.rkt")
(require "../fills/fills-test.rkt")
(require "../numbers/numbers-test.rkt")
(require "../fonts/fonts-test.rkt")

(require "../../../../xl/styles/borders.rkt")
(require "../../../../xl/styles/fills.rkt")
(require "../../../../xl/styles/numbers.rkt")
(require "../../../../xl/styles/fonts.rkt")

(provide (contract-out
          [set-cellXfses (-> void?)]
          [check-cellXfses (-> void?)]
          [check-chaos-number-cellXfses (-> void?)]
          [cellXfs_file path-string?]
          ))

(define (set-alignment-styles)
  (set-cell-range-alignment-style "A1" "left" "center")
  (set-cell-range-alignment-style "B1" "" "top")
  (set-cell-range-alignment-style "C1" "right" "")
  (set-cell-range-alignment-style "D1" "" ""))

(define (set-cellXfses)
  (set-border-styles)
  (set-font-styles)
  (set-number-styles)
  (set-fill-styles)
  (set-alignment-styles)
  )

(define (check-cellXfses)
  (check-equal? (length (STYLES-styles (*STYLES*))) 6)
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                (STYLE #f #f (ALIGNMENT-STYLE "center" "bottom") #f (FILL-STYLE "FFFFFF" "none")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                (STYLE #f #f (ALIGNMENT-STYLE "center" "bottom") #f (FILL-STYLE "FFFFFF" "gray125")))

  (check-equal? (list-ref
                 (STYLES-styles (*STYLES*)) 2)
                (STYLE
                 (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed")
                 (FONT-STYLE 10 "Arial" "0000FF")
                 (ALIGNMENT-STYLE "left" "center")
                 (NUMBER-STYLE "1" "0.00")
                 (FILL-STYLE "FF0000" "gray125")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                (STYLE
                 (BORDER-STYLE "000000" "dashed" "000000" "dashed" "000000" "dashed" "000000" "dashed")
                 (FONT-STYLE 15 "Arial" "0000FF")
                 (ALIGNMENT-STYLE "center" "top")
                 (NUMBER-STYLE "2" "#,###.00")
                 (FILL-STYLE "000000" "solid")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                (STYLE
                 (BORDER-STYLE #f #f #f #f #f #f "0000FF" "thick")
                 (FONT-STYLE 15 "宋体" "0000FF")
                 (ALIGNMENT-STYLE "right" "bottom")
                 (NUMBER-STYLE "3" "yyyymmdd")
                 (FILL-STYLE "FFFF00" "solid")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 5)
                (STYLE
                 (BORDER-STYLE #f #f #f #f "000000" "thin" #f #f)
                 (FONT-STYLE 15 "宋体" "FF0000")
                 (ALIGNMENT-STYLE "center" "bottom")
                 (NUMBER-STYLE "4" "yyyy-mm-dd")
                 (FILL-STYLE "FFFFFF" "solid")))
  )

(define (check-chaos-number-cellXfses)
  (check-equal? (length (STYLES-styles (*STYLES*))) 6)
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                (STYLE #f #f (ALIGNMENT-STYLE "center" "bottom") #f (FILL-STYLE "FFFFFF" "none")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                (STYLE #f #f (ALIGNMENT-STYLE "center" "bottom") #f (FILL-STYLE "FFFFFF" "gray125")))
  (check-equal? (list-ref
                 (STYLES-styles (*STYLES*)) 2)
                (STYLE
                 (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed")
                 (FONT-STYLE 10 "Arial" "0000FF")
                 (ALIGNMENT-STYLE "left" "center")
                 (NUMBER-STYLE "1" "0.00")
                 (FILL-STYLE "FF0000" "gray125")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                (STYLE
                 (BORDER-STYLE "000000" "dashed" "000000" "dashed" "000000" "dashed" "000000" "dashed")
                 (FONT-STYLE 15 "Arial" "0000FF")
                 (ALIGNMENT-STYLE "center" "top")
                 (NUMBER-STYLE "3" "#,###.00")
                 (FILL-STYLE "000000" "solid")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                (STYLE
                 (BORDER-STYLE #f #f #f #f #f #f "0000FF" "thick")
                 (FONT-STYLE 15 "宋体" "0000FF")
                 (ALIGNMENT-STYLE "right" "bottom")
                 (NUMBER-STYLE "5" "yyyymmdd")
                 (FILL-STYLE "FFFF00" "solid")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 5)
                (STYLE
                 (BORDER-STYLE #f #f #f #f "000000" "thin" #f #f)
                 (FONT-STYLE 15 "宋体" "FF0000")
                 (ALIGNMENT-STYLE "center" "bottom")
                 (NUMBER-STYLE "7" "yyyy-mm-dd")
                 (FILL-STYLE "FFFFFF" "solid")))
  )


(define (check-number-not-in-numfmts-cellXfses)
  (check-equal? (length (STYLES-styles (*STYLES*))) 6)
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 0)
                (STYLE #f #f (ALIGNMENT-STYLE "center" "bottom") #f (FILL-STYLE "FFFFFF" "none")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 1)
                (STYLE #f #f (ALIGNMENT-STYLE "center" "bottom") #f (FILL-STYLE "FFFFFF" "gray125")))
  (check-equal? (list-ref
                 (STYLES-styles (*STYLES*)) 2)
                (STYLE
                 (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed")
                 (FONT-STYLE 10 "Arial" "0000FF")
                 (ALIGNMENT-STYLE "left" "center")
                 (NUMBER-STYLE "1" 'APP)
                 (FILL-STYLE "FF0000" "gray125")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 3)
                (STYLE
                 (BORDER-STYLE "000000" "dashed" "000000" "dashed" "000000" "dashed" "000000" "dashed")
                 (FONT-STYLE 15 "Arial" "0000FF")
                 (ALIGNMENT-STYLE "center" "top")
                 (NUMBER-STYLE "3" 'APP)
                 (FILL-STYLE "000000" "solid")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 4)
                (STYLE
                 (BORDER-STYLE #f #f #f #f #f #f "0000FF" "thick")
                 (FONT-STYLE 15 "宋体" "0000FF")
                 (ALIGNMENT-STYLE "right" "bottom")
                 (NUMBER-STYLE "5" 'APP)
                 (FILL-STYLE "FFFF00" "solid")))
  (check-equal? (list-ref (STYLES-styles (*STYLES*)) 5)
                (STYLE
                 (BORDER-STYLE #f #f #f #f "000000" "thin" #f #f)
                 (FONT-STYLE 15 "宋体" "FF0000")
                 (ALIGNMENT-STYLE "center" "bottom")
                 (NUMBER-STYLE "7" 'APP)
                 (FILL-STYLE "FFFFFF" "solid")))
  )

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

       (strip-styles)
       (assemble-styles)

       (call-with-input-file cellXfs_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-cellXfs (STYLES-styles (*STYLES*))))
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
       )))

   (test-case
    "test-from-cellXfs-apply-true"

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
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>"
                                              (file->string cellXfs_apply_true_file)))))
       (check-cellXfses)
       )))

   (test-case
    "test-from-cellXfs-alignment-empty"

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
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>"
                                              (file->string cellXfs_alignment_empty_file)))))
       (check-cellXfses)
       )))

   (test-case
    "test-from-chaos-number-cellXfs"

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
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string chaos_numbers_file)))))
       (check-chaos-number-styles)

       (from-fills
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fills_file)))))
       (check-fill-styles)

       (from-cellXfs
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>"
                                              (file->string chaos_number_cellXfs_file)))))
       (check-chaos-number-cellXfses)

       (call-with-input-file chaos_number_cellXfs_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-cellXfs (STYLES-styles (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual)))))
       )))

   (test-case
    "test-from-number-not-in-numFmts-cellXfs"

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

       (from-fills
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fills_file)))))
       (check-fill-styles)

       (from-cellXfs
        (xml->hash (open-input-string (format "<styleSheet>~a</styleSheet>"
                                              (file->string chaos_number_cellXfs_file)))))
       (check-number-not-in-numfmts-cellXfses)

       (call-with-input-file chaos_number_cellXfs_file
         (lambda (expected)
           (call-with-input-string
            (lists->xml_content
             (to-cellXfs (STYLES-styles (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual)))))
       )))

   ))

(run-tests test-styles)
