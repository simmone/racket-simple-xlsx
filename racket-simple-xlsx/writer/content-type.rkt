#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-content-type (-> exact-nonnegative-integer? string?)]
          [write-content-type-file (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define S string-append)
 
(define (write-content-type sheet_count) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>@|(with-output-to-string 
    (lambda ()
      (let loop ([count 1])
        (when (<= count sheet_count)
          (printf 
            "<Override PartName=\"/xl/worksheets/sheet~a.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
            count)
          (loop (add1 count))))))|<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>
})

(define (write-content-type-file dir sheet_count)
  (with-output-to-file (build-path dir "[ContentType_Types].xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-content-type sheet_count)))))
    
    
