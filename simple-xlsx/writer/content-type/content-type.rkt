#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")
(require "../../xlsx/sheet.rkt")

(require "../../lib/lib.rkt")

(provide (contract-out
          [write-content-type (-> (is-a?/c xlsx%) string?)]
          [xlsx->content-type (-> (is-a?/c xlsx%) string?)]
          [write-content-type-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)

(define (write-header) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>

  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
})

(define (xlsx->content-type xlsx)
  (let ([sheet_list (get-field sheets xlsx)])
    (with-output-to-string 
      (lambda ()
        (let loop ([loop_list sheet_list]
                   [count 1])
          (when (not (null? loop_list))
                (let ([sheet (car loop_list)])
                  (if (eq? (sheet-type sheet) 'data)
                      (begin
                        (printf "<Override PartName=\"/xl/worksheets/sheet~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n" count)
                        (loop (cdr loop_list) (add1 count)))
                      (loop (cdr loop_list) count)))))

        (printf "\n")
        
        (let loop ([loop_list sheet_list]
                   [count 1])
          (when (not (null? loop_list))
                (let ([sheet (car loop_list)])
                  (if (eq? (sheet-type sheet) 'chart)
                      (begin
                        (printf "<Override PartName=\"/xl/charts/chart~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>\n" count)
                        (printf "<Override PartName=\"/xl/drawings/drawing~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>\n" count)
                        (printf "<Override PartName=\"/xl/chartsheets/sheet~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/>\n" count)
                        (when (not (null? (cdr loop_list)))
                              (printf "\n"))

                        (loop (cdr loop_list) (add1 count)))
                      (loop (cdr loop_list) count)))))
        ))))

(define (xlsx->shared-string xlsx)
  (with-output-to-string
    (lambda ()
      (when (> (hash-count (get-field string_item_map xlsx)) 0)
            (printf "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>")))))

(define (write-footer) @S{
</Types>
})

(define (write-content-type xlsx) @S{
@|(write-header)|

@|(prefix-each-line (xlsx->content-type xlsx) "  ")|
@|(prefix-each-line (xlsx->shared-string xlsx) "  ")|

@|(write-footer)|
})

(define (write-content-type-file dir xlsx)
  (with-output-to-file (build-path dir "[Content_Types].xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-content-type xlsx)))))
