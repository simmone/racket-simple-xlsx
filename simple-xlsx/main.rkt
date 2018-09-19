#lang racket

(provide (contract-out 
          [sheet-name-rows (-> path-string? string? list?)]
          [sheet-ref-rows (-> path-string? exact-nonnegative-integer? list?)]
          [read-xlsx% class?]
          [with-input-from-xlsx-file (-> path-string? (-> (is-a?/c read-xlsx%) void?) void?)]
          [load-sheet (-> string? (is-a?/c read-xlsx%) void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? (is-a?/c read-xlsx%) void?)]                    
          [get-sheet-names (-> (is-a?/c read-xlsx%) list?)]
          [get-cell-value (-> string? (is-a?/c read-xlsx%) any)]
          [get-sheet-dimension (-> (is-a?/c read-xlsx%) pair?)]
          [get-sheet-rows (-> (is-a?/c read-xlsx%) list?)]
          [xlsx% class?]
          [write-xlsx-file (-> (is-a?/c xlsx%) path-string? void?)]
          [oa_date_number->date (-> number? date?)]
          ))

(require "xlsx/xlsx.rkt")

(require "lib/lib.rkt")

(require "reader.rkt")

(require "writer.rkt")
