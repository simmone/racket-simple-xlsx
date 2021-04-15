#lang racket

(require "../../../xlsx/xlsx.rkt")
(require "../../../xlsx/sheet.rkt")

(provide (contract-out
          [create-printer-settings (-> path-string? (is-a?/c xlsx%) void?)]
          ))

(define (create-printer-settings dir xlsx)
  (make-directory* dir)

  (let loop ([loop_list (get-field sheets xlsx)])
    (when (not (null? loop_list))
          (when (eq? (sheet-type (car loop_list)) 'data)
                (with-output-to-file (build-path dir (string-append "printerSettings" (number->string (sheet-typeSeq (car loop_list))) ".bin"))
                  #:mode 'binary #:exists 'replace
                  (lambda ()
                    (for-each
                     (lambda (byte)
                       (write-byte byte))
                     '(0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 4 0 4 220 0 144 0 3 47 0 0 1 0 9 0 0 0 0 0 100 0 1 0 1 0 200 0 1 0 1 0 200 0 1 0 0 0 76 0 101 0 116 0 116 0 101 0 114 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 119 105 100 109 16 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 254 0 0 0 1 0 0 0 0 0 0 0 200 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0))))
                (loop (cdr loop_list))))))
