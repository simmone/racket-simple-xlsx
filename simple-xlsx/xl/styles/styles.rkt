#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")
(require "../../style/style.rkt")
(require "../../style/styles.rkt")
(require "../../style/border-style.rkt")
(require "../../style/fill-style.rkt")
(require "../../style/alignment-style.rkt")
(require "../../style/number-style.rkt")
(require "../../style/font-style.rkt")
(require "../../style/fill-style.rkt")
(require "../../lib/lib.rkt")

(require "numbers.rkt")
(require "fills.rkt")
(require "borders.rkt")
(require "fonts.rkt")
(require "cellXfs.rkt")

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
  (let ([xml_hash (xml->hash styles_file)])
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
        (printf "~a" (lists->xml (to-styles)))))))

(define (read-styles [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (from-styles (build-path dir "styles.xml"))))
