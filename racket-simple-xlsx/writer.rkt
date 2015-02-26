#lang racket

(provide (contract-out 
          [write-xlsx-file (-> list? (or/c list? #f) path-string? void?)]
          ))

;; data list:
;; '(((1 2 3 4) (1 2 3) (1 3 4)) ((3 4 5 6) (3 6 7 8)))
;; first level children is sheets
;; each sheet contains rows, row's length is not same
;; sheet name list:
;; '("Sheet1" "Sheet2" ...)
;; #f use default Sheet1, Sheet2... as sheet name

(require racket/date)

(require "lib/lib.rkt")
(require "writer/content-type.rkt")
(require "writer/_rels/rels.rkt")
(require "writer/docProps/docprops-app.rkt")
(require "writer/docProps/docprops-core.rkt")


(define (write-xlsx-file data_list sheet_name_list file_name)
  (let ([tmp_dir #f])
    (dynamic-wind
        (lambda () (set! tmp_dir (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
        (lambda ()
          (let* ([sheet_count (length data_list)]
                 [real_sheet_name_list (if sheet_name_list sheet_name_list (create-sheet-name-list sheet_count))])
            ;; [Content_Types].xml
            (write-content-type-file tmp_dir sheet_count)

            ;; _rels
            (let ([rels_dir (build-path tmp_dir "_rels")])
              (make-directory* rels_dir)
              (write-rels-file rels_dir))

            ;; docProps
            (let ([doc_props_dir (build-path tmp_dir "docProps")])
              (make-directory* doc_props_dir)
              (write-docprops-app-file doc_props_dir real_sheet_name_list)
              (write-docprops-core-file doc_props_dir (current-date)))
          ))
        (lambda () (void)))))
