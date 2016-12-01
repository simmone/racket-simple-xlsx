#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../lib/lib.rkt")
(require "../../xlsx.rkt")

(provide (contract-out
          [write-styles (-> list? string?)]
          [write-styles-file (-> path-string? (is-a?/c xlsx%) hash?)]
          ))

(define S string-append)

(define (write-styles style_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="@|(number->string (+ 2 (length style_list)))|"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>@|(with-output-to-string
  (lambda ()
    (for-each
     (lambda (style_rec)
       (printf "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"~a\"/><bgColor indexed=\"64\"/></patternFill></fill>" (first style_rec)))
     style_list)))|</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="@|(number->string (add1 (length style_list)))|"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>@|(with-output-to-string
(lambda ()
  (let loop ([loop_list style_list]
             [index 2])
    (when (not (null? loop_list))
          (printf "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"~a\" borderId=\"0\" xfId=\"0\" applyFill=\"1\"><alignment vertical=\"center\"/></xf>" index)
          (loop (cdr loop_list) (add1 index))))))|</cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>
})

(define (write-styles-file dir xlsx)
  (make-directory* dir)

  (with-output-to-file (build-path dir "styles.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-styles (send xlsx get-styles-list))))))
