#lang racket

(provide (contract-out 
          [with-input-from-xlsx-file (-> path-string? (-> any/c void?) void?)]          
          [get-sheet-names (-> any/c list?)]
          [get-cell-value (-> string? any/c any)]
          [load-sheet (-> string? any/c void?)]
          [get-sheet-dimension (-> any/c pair?)]
          [with-row (-> any/c (-> list? any) any)]
          [write-xlsx-file (-> list? (or/c list? #f) path-string? void?)]          
          ))

(require "reader.rkt")

(require "writer.rkt")
