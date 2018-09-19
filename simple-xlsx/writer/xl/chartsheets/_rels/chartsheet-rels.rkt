#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/file)
(require racket/list)
(require racket/contract)

(require "../../../../xlsx/xlsx.rkt")
(require "../../../../xlsx/sheet.rkt")

(provide (contract-out
          [write-chart-sheet-rels (-> exact-nonnegative-integer? string?)]
          [write-chart-sheet-rels-file (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define S string-append)

(define (write-chart-sheet-rels typeSeq) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing@|(number->string typeSeq)|.xml"/></Relationships>
                                                  })

(define (write-chart-sheet-rels-file dir xlsx)
  (when (ormap (lambda (rec) (eq? (sheet-type rec) 'chart)) (get-field sheets xlsx))
        (make-directory* dir)

        (let loop ([loop_list (get-field sheets xlsx)])
          (when (not (null? loop_list))
                (when (eq? (sheet-type (car loop_list)) 'chart)
                      (with-output-to-file (build-path dir (format "sheet~a.xml.rels" (sheet-typeSeq (car loop_list))))
                        #:exists 'replace
                        (lambda ()
                          (printf "~a" (write-chart-sheet-rels (sheet-typeSeq (car loop_list)))))))
                (loop (cdr loop_list))))))

