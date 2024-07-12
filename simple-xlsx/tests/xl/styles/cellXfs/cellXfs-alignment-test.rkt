#lang racket

(require fast-xml
         rackunit/text-ui
         rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../style/style.rkt"
         "../../../../style/alignment-style.rkt"
         "../../../../style/border-style.rkt"
         "../../../../style/font-style.rkt"
         "../../../../style/fill-style.rkt"
         "../../../../style/number-style.rkt"
         "../../../../style/styles.rkt"
         "../../../../style/assemble-styles.rkt"
         "../../../../style/set-styles.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/styles/cellXfs.rkt"
         racket/runtime-path
         "../borders/borders-test.rkt"
         "../fills/fills-test.rkt"
         "../numbers/numbers-test.rkt"
         "../fonts/fonts-test.rkt"
         "../../../../xl/styles/borders.rkt"
         "../../../../xl/styles/fills.rkt"
         "../../../../xl/styles/numbers.rkt"
         "../../../../xl/styles/fonts.rkt")

(define-runtime-path cellXfs_file "cellXfs.xml")
(define-runtime-path cellXfs_apply_true_file "cellXfs_apply_true.xml")
(define-runtime-path cellXfs_alignment_empty_file "cellXfs_alignment_empty.xml")
(define-runtime-path empty_cellXfs_file "empty_cellXfs.xml")
(define-runtime-path chaos_number_cellXfs_file "chaos_number_cellXfs.xml")
(define-runtime-path number_not_numfmts_cellXfs_file "number_not_numfmts_cellXfs.xml")

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
                 (NUMBER-STYLE "3" "yyyymmdd")
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
            (lists-to-xml_content
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
           )
         ))
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
           )))
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
           )
         ))
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
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string cellXfs_apply_true_file)))
         '(
           "styleSheet.cellXfs.xf.fillId"
           "styleSheet.cellXfs.xf.applyFont"
           "styleSheet.cellXfs.xf.fontId"
           "styleSheet.cellXfs.xf.applyBorder"
           "styleSheet.cellXfs.xf.borderId"
           "styleSheet.cellXfs.xf.numFmtId"
           "styleSheet.cellXfs.xf.alignment.horizontal"
           "styleSheet.cellXfs.xf.alignment.vertical"
           )))
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
           )
         ))
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
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string cellXfs_alignment_empty_file)))
         '(
           "styleSheet.cellXfs.xf.fillId"
           "styleSheet.cellXfs.xf.applyFont"
           "styleSheet.cellXfs.xf.fontId"
           "styleSheet.cellXfs.xf.applyBorder"
           "styleSheet.cellXfs.xf.borderId"
           "styleSheet.cellXfs.xf.numFmtId"
           "styleSheet.cellXfs.xf.alignment.horizontal"
           "styleSheet.cellXfs.xf.alignment.vertical"
           )))
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
           )
         ))
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
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string chaos_number_cellXfs_file)))
         '(
           "styleSheet.cellXfs.xf.fillId"
           "styleSheet.cellXfs.xf.applyFont"
           "styleSheet.cellXfs.xf.fontId"
           "styleSheet.cellXfs.xf.applyBorder"
           "styleSheet.cellXfs.xf.borderId"
           "styleSheet.cellXfs.xf.numFmtId"
           "styleSheet.cellXfs.xf.alignment.horizontal"
           "styleSheet.cellXfs.xf.alignment.vertical"
           )))
       (check-chaos-number-cellXfses)

       (call-with-input-file chaos_number_cellXfs_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content
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
         (open-input-string (format "<styleSheet>~a</styleSheet>"
                                    (file->string chaos_number_cellXfs_file)))
         '(
           "styleSheet.cellXfs.xf.fillId"
           "styleSheet.cellXfs.xf.applyFont"
           "styleSheet.cellXfs.xf.fontId"
           "styleSheet.cellXfs.xf.applyBorder"
           "styleSheet.cellXfs.xf.borderId"
           "styleSheet.cellXfs.xf.numFmtId"
           "styleSheet.cellXfs.xf.alignment.horizontal"
           "styleSheet.cellXfs.xf.alignment.vertical"
           )))

       (check-number-not-in-numfmts-cellXfses)

       (call-with-input-file chaos_number_cellXfs_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content
             (to-cellXfs (STYLES-styles (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual)))))
       )))

   ))

(run-tests test-styles)
