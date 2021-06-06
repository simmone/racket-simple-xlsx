#lang racket

(require simple-xml)

(require "../../../xlsx/xlsx.rkt")
(require "../../../lib/lib.rkt")

(provide (contract-out
          [fonts (-> list? list?)]
          [get-numFormatCode (-> hash? string?)]
          [numFmts (-> list? list?)]
          [fills (-> list? list?)]
          [borders (-> list? list?)]
          [cellStyleXfs (-> list?)]
          [cellXfs (-> list? list?)]
          [cellStyles (-> list?)]
          [dxfs (-> list?)]
          [styles (-> list? list? list? list? list? list?)]
          [write-styles (-> list? list? list? list? list? void?)]
          [read-styles (-> void?)]
          ))

(define (fonts font_list)
  (append
   (list "fonts" (cons "count" (number->string (add1 (length font_list)))))
   '(("font"
      ("sz" ("val" . "11"))
      ("color" ("theme" . "1"))
      ("name" ("val" . "宋体"))
      ("family" ("val" . "2"))
      ("charset" ("val" . "134"))
      ("scheme" ("val" . "minor"))))
   (let loop ([loop_list font_list]
              [result_list '()])
     (if (not (null? loop_list))
         (let ([fontSize (hash-ref (car loop_list) 'fontSize 11)]
               [fontColor (hash-ref (car loop_list) 'fontColor #f)]
               [fontName (hash-ref (car loop_list) 'fontName "宋体")])
           (loop
            (cdr loop_list)
            (cons
             (list
              "font"
              (list "sz" (cons "val" (number->string fontSize)))
              (if fontColor
                  (list "color" (cons "rgb" fontColor))
                  (list "color" (cons "theme" "1")))
              (list "name" (cons "val" fontName))
              '("family" ("val" . "2"))
              (if (not (regexp-match #rx"^([a-zA-Z]| |-|_|[0-9])+$" fontName))
                  (list "charset" (cons "val" "134"))
                  '())
              (list "scheme" (cons "val" "minor")))
             result_list)))
           (reverse result_list)))))

(define (get-numFormatCode format_hash)
  (cond
   [(hash-has-key? format_hash 'dateFormat)
    (let ([format_str (hash-ref format_hash 'dateFormat)])
      (with-output-to-string
        (lambda ()
          (let loop ([loop_list (string->list format_str)]
                     [critical_tag? #f]
                     [connect_tag? #f]
                     )
            (if (not (null? loop_list))
                (if (or
                     (char=? (car loop_list) #\y)
                     (char=? (car loop_list) #\m)
                     (char=? (car loop_list) #\d))
                    (if connect_tag?
                        (begin
                          (printf "&quot;~a" (car loop_list))
                          (loop (cdr loop_list) #t #f))
                        (begin
                          (printf "~a" (car loop_list))
                          (loop (cdr loop_list) #t #f)))
                    (if critical_tag?
                        (begin
                          (printf "&quot;~a" (car loop_list))
                          (loop (cdr loop_list) #f #t))
                        (begin
                          (printf "~a" (car loop_list))
                          (loop (cdr loop_list) #f #t))))
                (when connect_tag?
                      (printf "&quot;"))))
          (printf ";@"))))]
   [(or
     (hash-has-key? format_hash 'numberPrecision)
     (hash-has-key? format_hash 'numberThousands)
     (hash-has-key? format_hash 'numberPercent))
    (let* ([raw_number_precision (hash-ref format_hash 'numberPrecision 2)]
           [number_precision (if (natural? raw_number_precision) raw_number_precision 2)]
           [number_precision_str
            (format "0~a~a"
                    (if (not (= number_precision 0)) "." "")
                    (~a "" #:min-width number_precision #:pad-string "0"))]
           [number_thousands (hash-ref format_hash 'numberThousands #f)]
           [number_percent (hash-ref format_hash 'numberPercent #f)])
      (cond
       [number_thousands
        (string-append "#,###" number_precision_str)]
       [number_percent
        (string-append number_precision_str "%")]
       [else
        number_precision_str]
       ))]
   [(hash-has-key? format_hash 'formatCode)
    (hash-ref format_hash 'formatCode)]
   [else
    "0.00"]))

(define (numFmts numFmt_list)
  (append
   (list
    "numFmts" (cons "count" (format "~a" (add1 (length numFmt_list)))))
   '(("numFmt" ("numFmtId" . "164") ("formatCode" . "General")))
   (let loop ([loop_list numFmt_list]
              [loop_numId 164]
              [result_list '()])
     (if (not (null? loop_list))
         (loop
          (cdr loop_list)
          (add1 loop_numId)
          (cons
           (list "numFmt" (cons "numFmtId" (format "~a" (add1 loop_numId))) (cons "formatCode" (format "~a" (get-numFormatCode (car loop_list)))))
           result_list))
         (reverse result_list)))))

(define (fills fill_list)
  (append
   (list "fills" (cons "count" (number->string (+ 2 (length fill_list)))))
   '(("fill" ("patternFill" ("patternType" . "none"))))
   '(("fill" ("patternFill" ("patternType" . "gray125"))))
   (let loop ([loop_list fill_list]
              [result_list '()])
     (if (not (null? loop_list))
         (let ([backgroundColor (hash-ref (car loop_list) 'fgColor "FFFFFF")])
           (loop 
            (cdr loop_list)
            (cons
             (list "fill"
                   (list
                    "patternFill"
                    '("patternType" . "solid")
                    (list "fgColor" (cons "rgb" backgroundColor))
                    '("bgColor" ("indexed" . "64"))))
             result_list)))
         (reverse result_list)))))

(define (borders border_list)
  (append
   '("borders")
   (list (cons "count" (number->string (add1 (length border_list)))))
   '(("border" ("left") ("right") ("top") ("bottom") ("diagonal")))
   (let loop ([loop_list border_list]
              [result_list '()])
     (if (not (null? loop_list))
         (loop
          (cdr loop_list)
          (cons
           (append
            '("border")
            (let ([borderDirection (hash-ref (car loop_list) 'borderDirection 'all)]
                  [borderStyle (hash-ref (car loop_list) 'borderStyle 'thin)]
                  [borderColor (hash-ref (car loop_list) 'borderColor "000000")])
              (let direction-loop ([directions '(left right top bottom)]
                                   [direction_result '()])
                (if (not (null? directions))
                    (if (or 
                         (eq? borderDirection 'all)
                         (eq? (car directions) borderDirection))
                        (direction-loop (cdr directions)
                                        (cons
                                         (list
                                          (car directions)
                                          (cons "style" borderStyle)
                                          (list "color" (cons "rgb" borderColor)))
                                         direction_result))
                        (direction-loop (cdr directions) direction_result))
                    (reverse direction_result))))
            '(("diagonal")))
           result_list))
         (reverse result_list)))))

(define (cellStyleXfs)
  '("cellStyleXfs"
    ("count" . "1")
    ("xf" ("numFmtId" . "0") ("fontId" . "0") ("fillId" . "0") ("borderId" . "0")
     ("alignment" ("vertical" . "center")))))

(define (cellXfs style_list)
  (append
   (list "cellXfs" (cons "count" (number->string (add1 (length style_list)))))
   '(("xf"
     ("numFmtId" . "0") ("fontId" . "0") ("fillId" . "0") ("borderId" . "0") ("xfId" . "0")
     ("alignment" ("vertical" . "center"))))
   (let loop ([loop_list style_list]
              [result_list '()])
     (if (not (null? loop_list))
         (let* (
                [fill (hash-ref (car loop_list) 'fill 0)]
                [font (hash-ref (car loop_list) 'font 0)]
                [numFmt (hash-ref (car loop_list) 'numFmt 0)]
                [border (hash-ref (car loop_list) 'border 0)]
                [alignment_hash (hash-ref (car loop_list) 'alignment #f)]
                [alignment_list
                 (list
                  (list
                   "alignment"
                   (if alignment_hash
                       (append
                        (if (hash-has-key? alignment_hash 'horizontalAlign) (cons "horizontal" (hash-ref alignment_hash 'horizontalAlign)) '())
                        (if (hash-has-key? alignment_hash 'verticalAlign)
                            (cons "vertical" 
                                  (if (eq? (hash-ref alignment_hash 'verticalAlign) 'middle)
                                      'center
                                      (hash-ref alignment_hash 'verticalAlign)))
                            '()))
                       '("vertical" . "center"))))]
                )
           (loop
            (cdr loop_list)
            (cons
             (append
              (list "xf"
                    (cons "numFmtId" (format "~a" numFmt))
                    (cons "fontId" (format "~a" font))
                    (cons "fillId" (format "~a" fill))
                    (cons "borderId" (format "~a" border))
                    (cons "xfId" "0"))
              (if (not (= font 0)) '(("applyFont" . "1")) '())
              (if (not (= fill 0)) '(("applyFill" . "1")) '())
              (if (not (= border 0)) '(("applyBorder" . "1")) '())
              (if alignment_hash '(("applyAlignment" . "1")) '())
              alignment_list)
             result_list)))
         (reverse result_list)))))

(define (cellStyles)
  '("cellStyles" ("count" . "1")
    ("cellStyle" ("name" . "Normal") ("xfId" . "0") ("builtinId" . "0"))))

(define (dxfs)
  '("dxfs" ("count" . "0")))

(define (tableStyles)
  '("tableStyles" ("count" . "0") ("defaultTableStyle" . "TableStyleMedium9") ("defaultPivotStyle" . "PivotStyleLight16")))

(define (styles style_list fill_list font_list numFmt_list border_list)
  (list
   "styleSheet"
   '("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
   (numFmts numFmt_list)
   (fonts font_list)
   (fills fill_list)
   (borders border_list)
   (cellStyleXfs)
   (cellXfs style_list)
   (cellStyles)
   (dxfs)
   (tableStyles)))

(define (write-styles style_list fill_list font_list numFmt_list border_list)
  (let ([dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "xl")])
    (make-directory* dir)

    (with-output-to-file (build-path dir "styles.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->compact_xml
                      (styles 
                       style_list
                       fill_list
                       font_list
                       numFmt_list
                       border_list)))))))

(define (read-styles)
  (void))
