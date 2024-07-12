#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../lib/lib.rkt"
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../style/style.rkt"
         "../../../../style/styles.rkt"
         "../../../../style/number-style.rkt"
         "../../../../style/font-style.rkt"
         "../../../../style/fill-style.rkt"
         "../../../../style/border-style.rkt"
         "../../../../style/assemble-styles.rkt"
         "../../../../style/set-styles.rkt"
         "../../../../xl/styles/styles.rkt"
         racket/runtime-path
         "../borders/borders-test.rkt"
         "../fills/fills-test.rkt"
         "../fonts/fonts-test.rkt"
         "../numbers/numbers-test.rkt"
         "../cellXfs/cellXfs-alignment-test.rkt"
         "../../../../xl/styles/borders.rkt"
         "../../../../xl/styles/fills.rkt"
         "../../../../xl/styles/numbers.rkt"
         "../../../../xl/styles/fonts.rkt"
         "../../../../xl/styles/cellXfs.rkt")

(define-runtime-path styles_file "styles.xml")
(define-runtime-path chaos_number_styles_file "chaos_number_styles.xml")
(define-runtime-path complex_styles_file "complex_styles_test.xml")

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
            (lists-to-xml_content (to-styles))
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
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string borders_file)))
         '(
           "styleSheet.borders.border"
           "styleSheet.borders.border.left.color.rgb"
           "styleSheet.borders.border.right.color.rgb"
           "styleSheet.borders.border.top.color.rgb"
           "styleSheet.borders.border.bottom.color.rgb"
           "styleSheet.borders.border.left.style"
           "styleSheet.borders.border.right.style"
           "styleSheet.borders.border.top.style"
           "styleSheet.borders.border.bottom.style"
           )))
       (check-border-styles)

       (from-fonts
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fonts_file)))
         '(
           "styleSheet.fonts.font.sz.val"
           "styleSheet.fonts.font.name.val"
           "styleSheet.fonts.font.color.rgb"
           )))
         
       (check-font-styles)

       (from-numbers
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string numbers_file)))
         '(
           "styleSheet.numFmts.numFmt.formatCode"
           "styleSheet.numFmts.numFmt.numFmtId"
           )))
         
       (check-number-styles)

       (from-fills
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string fills_file)))
         '(
           "styleSheet.fills.fill.patternFill.patternType"
           "styleSheet.fills.fill.patternFill.fgColor.rgb"
           )))
         
       (check-fill-styles)

       (from-cellXfs
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string cellXfs_file)))
         '(
           "styleSheet.cellXfs.xf.fillId"
           "styleSheet.cellXfs.xf.applyFont"
           "styleSheet.cellXfs.xf.fontId"
           "styleSheet.cellXfs.xf.applyBorder"
           "styleSheet.cellXfs.xf.borderId"
           "styleSheet.cellXfs.xf.numFmtId"
           "styleSheet.cellXfs.xf.alignment.horizontal"
           "styleSheet.cellXfs.xf.alignment.vertical"
           )
         ))
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
