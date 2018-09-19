#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/file)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")
(require "../../xlsx/sheet.rkt")

(provide (contract-out
          [write-docprops-app (-> list? string?)]
          [write-docprops-app-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)

(define (print-sheet-variant sheet_list)
  (let ([sheet_type_count_map (make-hash)])
    (let loop ([loop_list sheet_list])
      (when (not (null? loop_list))
           (hash-set! sheet_type_count_map (sheet-type (car loop_list)) (add1 (hash-ref sheet_type_count_map (sheet-type (car loop_list)) 0)))
           (loop (cdr loop_list))))

    (with-output-to-string
      (lambda ()
        (printf "<HeadingPairs><vt:vector size=\"~a\" baseType=\"variant\">" (* (hash-count sheet_type_count_map) 2))
        (for-each
         (lambda (type_count)
           (cond
            [(eq? (car type_count) 'data)
             (printf "<vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>~a</vt:i4></vt:variant>" (cdr type_count))]
            [(eq? (car type_count) 'chart)
             (printf "<vt:variant><vt:lpstr>图表</vt:lpstr></vt:variant><vt:variant><vt:i4>~a</vt:i4></vt:variant>" (cdr type_count))]
            ))
         (sort (hash->list sheet_type_count_map) (lambda (c d) (string>? (symbol->string c) (symbol->string d)))  #:key car))))))
 
(define (write-docprops-app sheet_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>@|(print-sheet-variant sheet_list)|</vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="@|(number->string (length sheet_list))|" baseType="lpstr">@|(with-output-to-string
  (lambda ()
    (for-each
      (lambda (sheet)
        (printf "<vt:lpstr>~a</vt:lpstr>" (sheet-name sheet)))
     sheet_list)))|</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>
})

(define (write-docprops-app-file dir xlsx)
  (make-directory* dir)

  (with-output-to-file (build-path dir "app.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-docprops-app (get-field sheets xlsx))))))

