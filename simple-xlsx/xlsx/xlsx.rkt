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
                   (style_index->hash_map (hash/c hash? natural?))
                   (font_style_hash->index_map (hash/c hash? natural?))
                   (num_style_hash->index_map (hash/c hash? natural?))
                   (fill_style_hash->index_map (hash/c hash? natural?))
                   (border_style_hash->index_map (hash/c hash? natural?))
                   (alignment_style_hash->index_map (hash/c hash? natural?))
                   )
                  ]
          [new-xlsx (-> XLSX?)]
          [*CURRENT_XLSX* (parameter/c (or/c XLSX? #f))]
          [with-sheet (-> string? procedure?)]
          [with-sheet-ref (-> natural? procedure?)]
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
         [style_hash->index_map #:mutable]
         [style_index->hash_map #:mutable]
         [font_style_hash->index_map #:mutable]
         [num_style_hash->index_map #:mutable]
         [fill_style_hash->index_map #:mutable]
         [border_style_hash->index_map #:mutable]
         [alignment_style_hash->index_map #:mutable]
         ))

(define (new-xlsx)
  (XLSX "" 0
        (make-hash) (make-hash) (make-hash) (make-hash) (make-hash)
        (make-hash) (make-hash)
        '()
        (make-hash) (make-hash) (make-hash) (make-hash) (make-hash)
        (make-hash) (make-hash)
        ))

(define (with-sheet sheet_name user_proc)
  (parameterize ([*CURRENT_SHEET*
                  (list-ref (XLSX-sheet_list (*CURRENT_XLSX*))
                            (hash-ref (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name 0))])
    (user_proc)))

(define (with-sheet-ref sheet_index user_proc)
  (parameterize ([*CURRENT_SHEET*
                  (list-ref (XLSX-sheet_list (*CURRENT_XLSX*)) sheet_index 0)])
    (user_proc)))
