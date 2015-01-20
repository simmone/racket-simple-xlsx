#lang racket

(provide (contract-out
          [create-theme (-> void?)]
          ))

(define (create-theme)
  (make-directory* (build-path "xl" "theme"))
  (copy-file
   "theme.template"
   (build-path "xl" "theme" "theme1.xml")))
