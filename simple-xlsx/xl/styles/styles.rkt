#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")
(require "../../style/style.rkt")
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
          [cellStyleXfs (-> list?)]
          [cellStyles (-> list?)]
          [dxfs (-> list?)]
          [to-styles (-> list?)]
          [from-styles (-> path-string? void?)]
          [write-styles (->* () (path-string?) void?)]
          [read-styles (->* () (path-string?)  void?)]
          ))

(define (cellStyleXfs)
  '("cellStyleXfs"
    ("count" . "1")
    ("xf" ("numFmtId" . "0") ("fontId" . "0") ("fillId" . "0") ("borderId" . "0")
     ("alignment" ("vertical" . "center")))))

(define (cellStyles)
  '("cellStyles" ("count" . "1")
    ("cellStyle" ("name" . "Normal") ("xfId" . "0") ("builtinId" . "0"))))

(define (dxfs)
  '("dxfs" ("count" . "0")))

(define (tableStyles)
  '("tableStyles" ("count" . "0") ("defaultTableStyle" . "TableStyleMedium9") ("defaultPivotStyle" . "PivotStyleLight16")))

(define (to-styles)
  (let ([style_list
         (map (lambda (item) (style-from-hash-code (cdr item)))  (sort (hash->list (*INDEX->STYLE_MAP*)) < #:key car))]
        [border_list
         (map (lambda (item) (border-style-from-hash-code (cdr item))) (sort (hash->list (*BORDER_INDEX->STYLE_MAP*)) < #:key car))]
        [fill_list
         (map (lambda (item) (fill-style-from-hash-code (cdr item)))   (sort (hash->list (*FILL_INDEX->STYLE_MAP*)) < #:key car))]
        [font_list
         (map (lambda (item) (font-style-from-hash-code (cdr item)))   (sort (hash->list (*FONT_INDEX->STYLE_MAP*)) < #:key car))]
        [number_list
         (map (lambda (item) (number-style-from-hash-code (cdr item))) (sort (hash->list (*NUMBER_INDEX->STYLE_MAP*)) < #:key car))])

  (list
   "styleSheet"
   '("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
   (to-numbers number_list)
   (to-fonts font_list)
   (to-fills fill_list)
   (to-borders border_list)
   (cellStyleXfs)
   (to-cellXfs style_list)
   (cellStyles)
   (dxfs)
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

