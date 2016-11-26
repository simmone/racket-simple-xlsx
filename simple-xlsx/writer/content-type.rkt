#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(require "../xlsx.rkt")

(provide (contract-out
          [write-content-type (-> list? string?)]
          [write-content-type-file (-> path-string? list? void?)]
          ))

(define S string-append)

(define (sheets->content-type sheet_list)
  (with-output-to-string 
    (lambda ()
      (let loop ([loop_list sheet_list]
                 [count 1])
        (when (not (null? loop_list))
              (let ([sheet (car loop_list)])
                (if (eq? (sheet-type sheet) 'data)
                    (begin
                      (printf 
                       "<Override PartName=\"/xl/worksheets/sheet~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" count)
                      (loop (cdr loop_list) (add1 count)))
                    (loop (cdr loop_list) count)))))
      
      (let loop ([loop_list sheet_list]
                 [count 1])
        (when (not (null? loop_list))
              (let ([sheet (car loop_list)])
                (if (eq? (sheet-type sheet) 'chart)
                    (begin
                      (printf 
                       "<Override PartName=\"/xl/charts/chart~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocume;nt.drawingml.chart+xml\"/><Override PartName=\"/xl/drawings/drawing~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/chartsheets/sheet~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/>"
                       count count count)
                      (loop (cdr loop_list) (add1 count)))
                    (loop (cdr loop_list) count))))))))
 
(define (write-content-type sheet_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>@|(sheets->content-type sheet_list)|<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>
})

(define (write-content-type-file dir sheet_list)
  (with-output-to-file (build-path dir "[Content_Types].xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-content-type sheet_list)))))
