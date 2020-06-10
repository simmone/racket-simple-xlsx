#lang racket

(provide (contract-out 
          [sheet-name-rows (-> path-string? string? list?)]
          [sheet-ref-rows (-> path-string? exact-nonnegative-integer? list?)]
          [read-xlsx% class?]
          [with-input-from-xlsx-file (-> path-string? (-> (is-a?/c read-xlsx%) any) any)]
          [load-sheet (-> string? (is-a?/c read-xlsx%) void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? (is-a?/c read-xlsx%) void?)]                    
          [get-sheet-names (-> (is-a?/c read-xlsx%) list?)]
          [get-cell-value (-> string? (is-a?/c read-xlsx%) any)]
          [get-cell-formula (-> string? (is-a?/c read-xlsx%) string?)]
          [get-sheet-dimension (-> (is-a?/c read-xlsx%) pair?)]
          [get-sheet-rows (-> (is-a?/c read-xlsx%) list?)]
          [xlsx% class?]
          [write-xlsx-file (-> (is-a?/c xlsx%) path-string? void?)]
          [oa_date_number->date (-> number? date?)]
          [from-read-to-write-xlsx (-> (is-a?/c read-xlsx%) (is-a?/c xlsx%))]
          ))

(require "xlsx/xlsx.rkt")

(require "lib/lib.rkt")

(require "reader.rkt")

(require "writer.rkt")

(define (from-read-to-write-xlsx read_xlsx)
  (let ([xlsx (new xlsx%)])
    (let loop ([sheet_names (get-sheet-names read_xlsx)])
      (when (not (null? sheet_names))
            (load-sheet (car sheet_names) read_xlsx)
            (let ([rows (get-sheet-rows read_xlsx)])
              (when (> (length rows) 0)
                    (send xlsx add-data-sheet #:sheet_name (car sheet_names) #:sheet_data (get-sheet-rows read_xlsx))))
            (loop (cdr sheet_names))))
    xlsx))
