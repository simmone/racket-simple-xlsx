#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(provide (contract-out
          [write-workbook (-> list? string?)]
          [write-workbook-file (-> path-string? list? void?)]
          ))

(define S string-append)

(define (write-workbook sheet_name_list) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr filterPrivacy="1" defaultThemeVersion="124226"/><bookViews><workbookView xWindow="0" yWindow="90" windowWidth="19200" windowHeight="10590"/></bookViews><sheets>@|(with-output-to-string
(lambda ()
  (let loop ([name_list sheet_name_list]
             [num 1])
    (when (not (null? name_list))
          (printf "<sheet name=\"~a\" sheetId=\"~a\" r:id=\"rId~a\"/>" (car name_list) num num)
          (loop (cdr name_list) (add1 num))))))|</sheets><calcPr calcId="124519"/></workbook>
})

(define (write-workbook-file dir sheet_name_list)
  (with-output-to-file (build-path dir "workbook.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-workbook sheet_name_list)))))

