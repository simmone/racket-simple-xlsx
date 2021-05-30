#lang racket

(require simple-xml)

(require "../../../lib/lib.rkt")

(provide (contract-out
          [write-header (-> string?)]
          [fonts (-> list? list?)]
          [get-numFormatCode (-> hash? string?)]
          [numFmts (-> list? list?)]
          [fills (-> list? list?)]
          [borders (-> list? list?)]
          [cellStyleXfs (-> list?)]
          [cellXfs (-> list? list?)]
          [write-cellStyles (-> string?)]
          [write-dxfs (-> string?)]
          [write-footer (-> string?)]
          [write-styles (-> list? list? list? list? list? string?)]
          [write-styles-file (-> path-string? list? list? list? list? list? void?)]
          ))

(define (fonts font_list)
  (append
   (list "fonts" (cons "count" (number->string (add1 (length font_list)))))
   '("font"
     ("sz" ("val" . "11"))
     ("color" ("theme" . "1"))
     ("name" ("val" . "宋体"))
     ("family" ("val" . "2"))
     ("charset" ("val" . "134"))
     ("scheme" ("val" . "minor")))
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
              (list "family" (cons val "2"))
              (if (not (regexp-match #rx"^([a-zA-Z]| |-|_|[0-9])+$" fontName))
                  (list "charset" (cons "val" "134"))
                  '())
              (list "scheme" (cons "val" "minor")))
             result_list))
           (reverse result_list))))))

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
    (hash-ref format_hash 'formatCode)]))

(define (numFmts numFmt_list)
  (append
   (list
    "numFmts" (cons "count" (format "~a" (add1 (length numFmt_list)))))
   '("numFmt" ("numFmtId" . "164") ("formatCode" . "General"))
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
   '("fill" ("patternFill" ("patternType" . "none")))
   '("fill" ("patternFill" ("patternType" . "gray125")))
   (let loop ([loop_list fill_list]
              [result_list '()])
     (if (not (null? loop_list))
         (let ([backgroundColor (hash-ref (car loop_list) 'fgColor "FFFFFF")])
           (loop 
            (cdr loop_list)
            (list "fill"
                  '("patternFill" ("patternType" . "solid"))
                  (list "fgColor" (cons "rgb" backgroundColor))
                  '("bgColor" ("indexed" . "64")))))
         (reverse result_list)))))

(define (borders border_list)
  (append
   (list "borders" (cons "count" (number->string (add1 (length border_list)))))
   '("border" ("left") ("right") ("top") ("bottom") ("diagonal"))
   (let loop ([loop_list border_list]
              [result_list '()])
     (if (not (null? loop_list))
         (loop
          (cdr loop_list)
          (cons
           (list
            "border"
            (let ([borderDirection (hash-ref (car loop_list) 'borderDirection 'all)]
                  [borderStyle (hash-ref (car loop_list) 'borderStyle 'thin)]
                  [borderColor (hash-ref (car loop_list) 'borderColor "000000")])
              (let direction-loop ([directions '(left right top bottom)]
                                   [direction_result '()])
                (if (not (null? directions))
                    (if (or 
                         (eq? borderDirection 'all)
                         (eq? (car directions) borderDirection))
                        (direction-loop (cdr loop_list)
                                        (cons
                                         (list
                                          (car directions)
                                          (cons "style" borderStyle)
                                          (list "color" (cons "rgb" borderColor)))
                                         direction_result))
                        (reverse direction_result))))))
           result_list))
         (reverse result_list)))
   '("diagonal")))

(define (cellStyleXfs)
  '("cellStyleXfs"
    ("count" . "1")
    ("xf" ("numFmtId" . "0") ("fontId" . "0") ("fillId" . "0") ("borderId" . "0")
     ("alignment" ("vertical" . "center")))))

(define (cellXfs style_list)
<cellXfs count="@|(number->string (add1 (length style_list)))|">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>
@|(with-output-to-string
    (lambda ()
      (let loop ([loop_list style_list])
        (when (not (null? loop_list))
          (let* ([fill (hash-ref (car loop_list) 'fill 0)]
                 [font (hash-ref (car loop_list) 'font 0)]
                 [numFmt (hash-ref (car loop_list) 'numFmt 0)]
                 [border (hash-ref (car loop_list) 'border 0)]
                 [alignment_hash (hash-ref (car loop_list) 'alignment #f)]
                 [alignment_str
                    (format "<alignment~a/>"
                      (if alignment_hash
                         (format
                           "~a~a"
                           (if (hash-has-key? alignment_hash 'horizontalAlign) (format " horizontal=\"~a\"" (hash-ref alignment_hash 'horizontalAlign)) "")
                           (if (hash-has-key? alignment_hash 'verticalAlign) (format " vertical=\"~a\""
                                                                               (if (eq? (hash-ref alignment_hash 'verticalAlign) 'middle)
                                                                                  'center
                                                                                  (hash-ref alignment_hash 'verticalAlign)))
                                                                             ""))
                       " vertical=\"center\""))]
                )
            (printf "  <xf numFmtId=\"~a\" fontId=\"~a\" fillId=\"~a\" borderId=\"~a\" xfId=\"0\"" numFmt font fill border)
            (when (not (= font 0)) (printf " applyFont=\"1\""))
            (when (not (= fill 0)) (printf " applyFill=\"1\""))
            (when (not (= border 0)) (printf " applyBorder=\"1\""))
            (when alignment_hash (printf " applyAlignment=\"1\""))
            (printf ">~a</xf>\n" alignment_str))
          (loop (cdr loop_list))))))|</cellXfs>
})

(define (write-cellStyles) @S{
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
})

(define (write-dxfs) @S{
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
})

(define (write-footer) @S{
</styleSheet>
})

(define (styles style_list fill_list font_list numFmt_list border_list)
  (append
   (list "styleSheet" (cons "xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
   (numFmts numFmt_list)
   (fonts font_list)
   (fills fill_list)
   (borders border_list)
   (cellStyleXfs)
@|(prefix-each-line (write-cellXfs style_list) "  ")|

@|(prefix-each-line (write-cellStyles) "  ")|

@|(prefix-each-line (write-dxfs) "  ")|

@|(write-footer)|
})

(define (write-styles-file dir style_list fill_list font_list numFmt_list border_list)
  (make-directory* dir)

  (with-output-to-file (build-path dir "styles.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-styles 
                    style_list
                    fill_list
                    font_list
                    numFmt_list
                    border_list
                    )))))
