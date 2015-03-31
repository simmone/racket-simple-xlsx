#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>

(provide (contract-out
          [write-styles (-> list? string?)]
          [write-styles-file (-> path-string? hash? hash?)]
          ))

(define S string-append)

(define (write-styles style_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>
})

(define (write-styles-file dir sheet_attr_hash)
  (with-output-to-file (build-path dir "styles.xml")
    #:exists 'replace
    (lambda ()
      (let ([styles_list '()]
            [color_style_map (make-hash)]
            [index 1])
        (hash-for-each
         sheet_attr_hash
         (lambda (sheet_index col_attr_hash)
      (printf "~a" (write-styles styles_list)))))

