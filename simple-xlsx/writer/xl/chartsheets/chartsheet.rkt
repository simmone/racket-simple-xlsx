#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/file)
(require racket/class)
(require racket/contract)

(require "../../../xlsx/xlsx.rkt")
(require "../../../xlsx/sheet.rkt")

(provide (contract-out
          [write-chart-sheet (-> exact-nonnegative-integer? string?)]
          [write-chart-sheet-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)

(define (write-chart-sheet typeSeq) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr/><sheetViews><sheetView zoomScale="115" workbookViewId="0" zoomToFit="1"/></sheetViews><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></chartsheet>
})

(define (write-chart-sheet-file dir xlsx)
  (when (ormap (lambda (rec) (eq? (sheet-type rec) 'chart)) (get-field sheets xlsx))
        (make-directory* dir)

        (let loop ([loop_list (get-field sheets xlsx)])
          (when (not (null? loop_list))
                (when (eq? (sheet-type (car loop_list)) 'chart)
                      (with-output-to-file (build-path dir (format "sheet~a.xml" (sheet-typeSeq (car loop_list))))
                        #:exists 'replace
                        (lambda ()
                          (printf "~a" (write-chart-sheet (sheet-typeSeq (car loop_list)))))))
                (loop (cdr loop_list))))))

