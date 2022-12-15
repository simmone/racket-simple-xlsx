#lang racket

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")

(provide (contract-out
          [printer-settings (-> bytes?)]
          [read-printer-settings (-> void?)]
          [write-printer-settings (->* () (path-string?) void?)]
          ))

(define (printer-settings)
  (bytes 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 4 0 4 220 0 144 0 3 47 0 0 1 0 9 0 0 0 0 0 100 0 1 0 1 0 200 0 1 0 1 0 200 0 1 0 0 0 76 0 101 0 116 0 116 0 101 0 114 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 119 105 100 109 16 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 254 0 0 0 1 0 0 0 0 0 0 0 200 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 10))

(define (write-printer-settings [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl" "printerSettings"))])
    (make-directory* dir)

  (let loop ([data_sheet_list (filter DATA-SHEET? (XLSX-sheet_list (*XLSX*)))]
             [index 1])
    (when (not (null? data_sheet_list))
      (with-output-to-file (build-path dir (format "printerSettings~a.bin" index))
        #:mode 'binary #:exists 'replace
        (lambda ()
          (write-bytes (printer-settings))))
      (loop (cdr data_sheet_list) (add1 index))))))

(define (read-printer-settings)
  (void))
