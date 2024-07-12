#lang racket

(require fast-xml
         "../../xlsx/xlsx.rkt"
         "../../sheet/sheet.rkt"
         "../../style/style.rkt"
         "../../style/styles.rkt"
         "../../style/border-style.rkt"
         "../../style/fill-style.rkt"
         "../../style/alignment-style.rkt"
         "../../style/number-style.rkt"
         "../../style/font-style.rkt"
         "../../style/fill-style.rkt"
         "../../lib/lib.rkt"
         "numbers.rkt"
         "fills.rkt"
         "borders.rkt"
         "fonts.rkt"
         "cellXfs.rkt")

(provide (contract-out
          [dxfs (-> list?)]
          [to-styles (-> list?)]
          [from-styles (-> path-string? void?)]
          [write-styles (->* () (path-string?) void?)]
          [read-styles (->* () (path-string?)  void?)]
          ))
(define (dxfs)
  '("dxfs" ("count" . "0")))

(define (tableStyles)
  '("tableStyles" ("count" . "0") ("defaultTableStyle" . "TableStyleMedium9") ("defaultPivotStyle" . "PivotStyleLight16")))

(define (to-styles)
  (let ([style_list (STYLES-styles (*STYLES*))]
        [border_list (STYLES-border_list (*STYLES*))]
        [fill_list (STYLES-fill_list (*STYLES*))]
        [font_list (STYLES-font_list (*STYLES*))]
        [number_list (STYLES-number_list (*STYLES*))])
  (list
   "styleSheet"
   '("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
   (to-numbers number_list)
   (to-fonts font_list)
   (to-fills fill_list)
   (to-borders border_list)
   (to-cellXfs style_list)
   (tableStyles))))

(define (from-styles styles_file)
  (let ([xml_hash (xml-file-to-hash
                   styles_file
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
                     "styleSheet.fills.fill.patternFill.patternType"
                     "styleSheet.fills.fill.patternFill.fgColor.rgb"
                     "styleSheet.fonts.font.sz.val"
                     "styleSheet.fonts.font.name.val"
                     "styleSheet.fonts.font.color.rgb"
                     "styleSheet.numFmts.numFmt.formatCode"
                     "styleSheet.numFmts.numFmt.numFmtId"
                     "styleSheet.cellXfs.xf.fillId"
                     "styleSheet.cellXfs.xf.applyFont"
                     "styleSheet.cellXfs.xf.fontId"
                     "styleSheet.cellXfs.xf.applyBorder"
                     "styleSheet.cellXfs.xf.borderId"
                     "styleSheet.cellXfs.xf.numFmtId"
                     "styleSheet.cellXfs.xf.alignment.horizontal"
                     "styleSheet.cellXfs.xf.alignment.vertical"
                     )
                   )])

    (from-numbers xml_hash)
    (from-fonts xml_hash)
    (from-fills xml_hash)
    (from-borders xml_hash)
    (from-cellXfs xml_hash)))

(define (write-styles [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (make-directory* dir)

    (with-output-to-file (build-path dir "styles.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists-to-xml (to-styles)))))))

(define (read-styles [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (from-styles (build-path dir "styles.xml"))))
