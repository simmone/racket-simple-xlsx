#lang racket

(provide (contract-out 
          [with-input-from-xlsx-file (-> path-string? (-> any/c void?) void?)]
          [load-sheet (-> string? any/c void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? any/c void?)]                    
          [get-sheet-names (-> any/c list?)]
          [get-cell-value (-> string? any/c any)]
          [get-sheet-dimension (-> any/c pair?)]
          [with-row (-> any/c (-> list? any) any)]
          [xlsx% class?]
          [write-xlsx-file (-> (is-a?/c xlsx%) path-string? void?)]
          ))

(require "xlsx.rkt")

(require "reader.rkt")

(require "writer.rkt")
