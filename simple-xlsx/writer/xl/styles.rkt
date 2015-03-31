#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(require "../../lib/lib.rkt")

(provide (contract-out
          [write-styles (-> list? string?)]
          [write-styles-file (-> path-string? hash? hash?)]
          ))

(define S string-append)

(define (write-styles style_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="@|(number->string (+ 2 (length style_list)))|"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>@|(with-output-to-string
  (lambda ()
    (for-each
     (lambda (style_rec)
       (printf "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"~a\"/><bgColor indexed=\"64\"/></patternFill></fill>" (first style_rec)))
     style_list)))|</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>
})

(define (write-styles-file dir sheet_attr_map)
  (let ([color_style_map (make-hash)])
    (with-output-to-file (build-path dir "styles.xml")
      #:exists 'replace
      (lambda ()
        (let ([style_list '()]
              [style_index 2])
          (hash-for-each
           sheet_attr_map
           (lambda (sheet_index col_attr_map)
             (hash-for-each
              col_attr_map
              (lambda (col_name col_attr)
                (let ([color (col-attr-color col_attr)])
                  (when (not (string=? color ""))
                        (when (not (hash-has-key? color_style_map color))
                              (hash-set! color_style_map color style_index)
                              (set! style_list `(,@style_list (,color)))
                              (set! style_index (add1 style_index)))))))))

          (printf "~a" (write-styles style_list)))))
    color_style_map))

