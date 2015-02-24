#lang racket

(provide (contract-out 
          [write-xlsx-file (-> list? path-string? void?)]
          ))

;; xlsx data use this struct:
;; '(((1 2 3 4) (1 2 3) (1 3 4)) ((3 4 5 6) (3 6 7 8)))
;; first level children is sheets
;; each sheet contains rows, row's length is not same

(require "writer/content-type.rkt")

(define (write-xlsx-file data_list file_name)
  (let ([tmp_dir #f])
    (dynamic-wind
        (lambda () (set! tmp_dir (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
        (lambda ()
          (let ([sheet_count (length data_list)])
            (write-content-type-file tmp_dir sheet_count)
          ))
        (lambda () (void)))))
