#lang racket

(provide (contract-out
          [create-printer-settings (-> path-string? exact-nonnegative-integer? void?)]
          ))

(define (create-printer-settings dir sheet_count)
  (let loop ([nums sheet_count])
    (when (>= nums 1)
          (copy-file
           (build-path "writer" "xl" "printerSettings" "printerSettings.template")
           (build-path dir (string-append "printerSettings" (number->string nums) ".bin")))
          (loop (sub1 nums)))))

