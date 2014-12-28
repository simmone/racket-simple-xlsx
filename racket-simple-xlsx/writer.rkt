#lang racket

(provide (contract-out 
          [with-output-to-xlsx-file (-> path-string? void?)]
          ))

(require "writer/content-type.rkt")

(define (with-output-to-xlsx-file file_name)
  (write-content-type 3)
  )
