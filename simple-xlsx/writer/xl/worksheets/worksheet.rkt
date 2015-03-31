#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)

(require "../../../lib/lib.rkt")

;; write-sheet '('()...)
(provide (contract-out
          [write-sheet (-> list? hash? hash? exact-nonnegative-integer? boolean? string?)]
          [write-sheet-file (-> path-string? exact-nonnegative-integer? list? hash? hash? void?)]
          ))

(define S string-append)

(define (write-sheet sheet_data_list string_index_map sheet_attr_map sheet_index is_active?) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="@|(if (null? sheet_data_list) "A1" (string-append "A1:" (get-dimension sheet_data_list)))|"/><sheetViews><sheetView @|(if is_active? "tabSelected=\"1\"" "")| workbookViewId="0">@|(if (null? sheet_data_list) "" "<selection activeCell=\"A1\" sqref=\"A1\"/>")|</sheetView></sheetViews><sheetFormatPr defaultRowHeight="13.5"/>@|(with-output-to-string
(lambda ()
  (when (hash-has-key? sheet_attr_map sheet_index)
        (printf "<cols>")
        (let loop-col ([loop_cols (sort (hash->list (hash-ref sheet_attr_map sheet_index)) string<? #:key car)])
          (when (not (null? loop_cols))
                (let* ([col (car loop_cols)]
                       [col_index (abc->number (car col))]
                       [col_attr (cdr col)]
                       [col_width (col-attr-width col_attr)])
                  (printf "<col min=\"~a\" max=\"~a\" width=\"~a\" customWidth=\"1\"/>"
                          col_index
                          col_index
                          (exact->inexact (cx-round (/ col_width 8) 2))))
                (loop-col (cdr loop_cols))))
        (printf "</cols>"))

  (printf "<sheetData>")
                        
  (let loop-row ([loop_rows sheet_data_list]
                 [row_seq 1])
    (when (not (null? loop_rows))
          (printf "<row r=\"~a\">" row_seq)
          (let ([item_list (car loop_rows)])
            (when (not (null? item_list))
                  (let loop-col ([loop_cols item_list]
                                 [col_seq 1])
                    (when (not (null? loop_cols))
                          (let ([cell (car loop_cols)]
                                [dimension (string-append (number->abc col_seq) (number->string row_seq))])
                            (cond
                             [(string? cell)
                              (printf "<c r=\"~a\" t=\"s\"><v>~a</v></c>" dimension (hash-ref string_index_map cell))]
                             [(number? cell)
                              (printf "<c r=\"~a\"><v>~a</v></c>" dimension (number->string (exact->inexact cell)))]
                             [else
                              (printf "<c r=\"~a\"><v>0</v></c>" dimension)]))
                          (loop-col (cdr loop_cols) (add1 col_seq))))))
          (printf "</row>")
          (loop-row (cdr loop_rows) (add1 row_seq))))))|</sheetData><phoneticPr fontId="1" type="noConversion"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" horizontalDpi="200" verticalDpi="200" r:id="rId1"/></worksheet>
})

(define (write-sheet-file dir sheet_index sheet_data_list string_index_map sheet_attr_map)
  (with-output-to-file (build-path dir (string-append "sheet" (number->string sheet_index) ".xml"))
    #:exists 'replace
    (lambda ()
      (if (= sheet_index 1)
          (printf "~a" (write-sheet sheet_data_list string_index_map sheet_attr_map sheet_index #t))
          (printf "~a" (write-sheet sheet_data_list string_index_map sheet_attr_map sheet_index #f))))))
