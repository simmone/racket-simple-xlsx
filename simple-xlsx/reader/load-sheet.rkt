#lang racket

(provide (contract-out
          [load-data-sheet-file (-> path-string? DATA-SHEET?)]
          [load-chart-sheet-file (-> path-string? CHART-SHEET?)]
          ))

(require simple-xml)

(require "../lib/lib.rkt")
(require "../sheet/sheet.rkt")
(require "../xlsx/range-lib.rkt")

(define (load-chart-sheet-file sheet_file)
  (void))

(define (load-data-sheet-file sheet_file)
  (let ([sheet 
         (DATA-SHEET
          '(0 . 0) (make-hash) (make-hash) (make-hash) '(0 . 0)
          (make-hash) (make-hash) (make-hash) (make-hash) (make-hash)
          (make-hash))])

    (let ([xml_hash (xml->hash sheet_file)])

      (set-DATA-SHEET-dimension!
       sheet
       (dimension->pair (hash-ref xml_hash "worksheet.dimension.ref")))

      (let loop-row ([row_count 1])
        (when (<= row_count (hash-ref xml_hash "worksheet.sheetData.row's count"))
          (let loop-col ([col_count 1])
            (when (<= col_count (hash-ref xml_hash "worksheet.cols.col's count"))
              (let (
                    [para_r (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.r" row_count col_count) #f)]
                    [para_v (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.v" row_count col_count) #f)]
                    [para_t (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.t" row_count col_count) #f)]
                    [para_s (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.s" row_count col_count) #f)]
                    [para_f (hash-ref xml_hash (format "worksheet.sheetData.row~a.c~a.f" row_count col_count) #f)]
                    )

                (when para_r
                      (hash-set! (DATA-SHEET-rvtsf_map sheet) para_r (list para_v para_t para_s para_f))))
              (loop-col (add1 col_count))))
          (loop-row (add1 row_count)))))
    sheet))
