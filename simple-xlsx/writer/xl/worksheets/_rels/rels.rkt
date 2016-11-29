#lang racket

(require "../../../xlsx.rkt")

(provide (contract-out
          [write-worksheets-rels-file (-> path-string? list? void?)]
          ))

(define (write-worksheets-rels-file dir sheet_list)
  (let loop ([loop_list sheet_list])
    (when (not (null? loop_list))
          (when (eq? (sheet-type (car loop_list)) 'data)
                (let ([type_seq (number->string (sheet-typeSeq (car loop_list)))])
                  (with-output-to-file (build-path dir (string-append "sheet" type_seq ".xml.rels"))
                    (lambda ()
                      (printf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
                      (printf "<Relationshnips xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId~a\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" Target=\"../printerSettings/printerSettings~a.bin\"/></Relationships>"
                              type_seq type_seq)))))
          (loop (add1 num)))))
