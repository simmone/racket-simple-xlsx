#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")
(require "../../xlsx/sheet.rkt")

(provide (contract-out
          [write-workbook (-> list? string?)]
          [write-workbook-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)

(define (write-workbook sheet_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr filterPrivacy="1" defaultThemeVersion="124226"/><bookViews><workbookView xWindow="0" yWindow="90" windowWidth="19200" windowHeight="10590"/></bookViews><sheets>@|(with-output-to-string
(lambda ()
  (let loop ([loop_list sheet_list])
    (when (not (null? loop_list))
          (let ([sheet (car loop_list)])
            (printf "<sheet name=\"~a\" sheetId=\"~a\" r:id=\"rId~a\"/>" (sheet-name sheet) (sheet-seq sheet) (sheet-seq sheet)))
          (loop (cdr loop_list))))))|</sheets><calcPr calcId="124519"/></workbook>
})

(define (write-workbook-file dir xlsx)
  (make-directory* dir)

  (with-output-to-file (build-path dir "workbook.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-workbook (get-field sheets xlsx))))))

