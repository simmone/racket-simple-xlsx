#lang racket

(provide (contract-out
          [write-theme-file (-> path-string? void?)]
          ))

(define (write-theme-file dir)
  (copy-file
   (build-path "writer" "xl" "theme" "theme.template")
   (build-path dir "theme1.xml")))
