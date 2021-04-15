#lang at-exp racket/base

(require racket/port)
(require racket/list)
(require racket/contract)
(require racket/file)
(require racket/class)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../xlsx/sheet.rkt")

(provide (contract-out
          [write-drawing-rels (-> exact-nonnegative-integer? string?)]
          [write-drawing-rels-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)

(define (write-drawing-rels typeSeq) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart@|(number->string typeSeq)|.xml"/></Relationships>
})

(define (write-drawing-rels-file dir xlsx)
  (when (ormap (lambda (rec) (eq? (sheet-type rec) 'chart)) (get-field sheets xlsx))
        (make-directory* dir)

        (let loop ([loop_list (get-field sheets xlsx)])
          (when (not (null? loop_list))
                (when (eq? (sheet-type (car loop_list)) 'chart)
                      (with-output-to-file (build-path dir (format "drawing~a.xml.rels" (sheet-typeSeq (car loop_list))))
                        #:exists 'replace
                        (lambda ()
                          (printf "~a" (write-drawing-rels (sheet-typeSeq (car loop_list)))))))
                (loop (cdr loop_list))))))

