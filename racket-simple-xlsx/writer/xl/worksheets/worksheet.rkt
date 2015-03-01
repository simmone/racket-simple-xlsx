#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(require "../../../lib/lib.rkt")

;; write-sheet '('()...)
(provide (contract-out
          [write-sheet (-> list? hash? string?)]
          ))

(define S string-append)

(define (write-sheet data_list string_index_map) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:@|(get-dimension data_list)|"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.5"/><sheetData>@|(with-output-to-string
(lambda ()
  (let loop-row ([loop_rows data_list]
                 [row_seq 1])
    (when (not (null? loop_rows))
          (printf "<row r=\"~a\">" row_seq)
          (let ([item_list (car loop_rows)])
            (when (not (null? item_list))
                  (let loop-col ([loop_cols item_list]
                                 [col_seq 1])
                    (when (not (null? loop_cols))
                          (let ([cell (car loop_cols)]
                                [dimension (string-append (number->abc col_seq) row_seq)])
                            (cond
                             [(string? cell)
                              (printf "<c r=\"~a\" t=\"s\"><v>~a</v></c>" dimension (hash-ref string_index_map cell))]
                             [(number? cell)
                              (printf "<c r=\"~a\"><v>~a</v></c>" dimension (number->string cell))]
                             [else
                              (printf "<c r=\"~a\"><v>0</v></c>" dimension)]))
                          (loop-col (cdr loop_cols) (add1 col_seq))))))
          (printf "</row>")
          (loop-row (cdr loop_rows) (add1 row_seq))))))|</sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" horizontalDpi="200" verticalDpi="200" r:id="rId1"/></worksheet>
})
