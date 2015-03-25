#lang racket

(provide (contract-out 
          [with-input-from-xlsx-file (-> path-string? (-> any/c void?) void?)]          
          [get-sheet-names (-> any/c list?)]
          [get-cell-value (-> string? any/c any)]
          [load-sheet (-> string? any/c void?)]
          [get-sheet-dimension (-> any/c pair?)]
          [with-row (-> any/c (-> list? any) any)]
          [xlsx-data% class?]
          [xlsx-data? any/c]
          [write-xlsx-file (-> xlsx-data? path-string? void?)]
          ))

(require "reader.rkt")

(require "writer.rkt")
