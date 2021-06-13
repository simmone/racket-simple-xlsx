#lang racket

(require "../lib/lib.rkt")
(require "../lib/dimension.rkt")
(require "../sheet/sheet.rkt")

(provide (contract-out
          [struct XLSX
                  (
                   (xlsx_dir path-string?)
                   (sheet_count natural?)
                   (sheet_index_id_map (hash/c natural? string?))
                   (sheet_index_name_map (hash/c natural? string?))
                   (sheet_name_index_map (hash/c string? natural?))
                   (sheet_index_rid_map (hash/c natural? string?))
                   (sheet_rid_rel_map (hash/c string? string?))
                   (sheet_index_rel_map (hash/c natural? string?))
                   (shared_strings_map (hash/c string? natural?))
                   (sheet_list (listof (or/c DATA-SHEET? CHART-SHEET?)))
                   (style_hash->index_map (hash/c hash? natural?))
                   (font_hash->index_map (hash/c hash? natural?))
                   (num_hash->index_map (hash/c hash? natural?))
                   (fill_hash->index_map (hash/c hash? natural?))
                   (border_hash->index_map (hash/c hash? natural?))
                   )
                  ]
          [new-xlsx (-> XLSX?)]
          [*CURRENT_XLSX* (parameter/c (or/c XLSX? #f))]
          ))

(define *CURRENT_XLSX* (make-parameter #f))

(struct XLSX
        (
         [xlsx_dir #:mutable]
         [sheet_count #:mutable]
         [sheet_index_id_map #:mutable]
         [sheet_index_name_map #:mutable]
         [sheet_name_index_map #:mutable]
         [sheet_index_rid_map #:mutable]
         [sheet_rid_rel_map #:mutable]
         [sheet_index_rel_map #:mutable]
         [shared_strings_map #:mutable]
         [sheet_list #:mutable]
         ))

(define (new-xlsx)
  (XLSX "" 0
        (make-hash) (make-hash) (make-hash) (make-hash) (make-hash)
        (make-hash) (make-hash)
        '()))
