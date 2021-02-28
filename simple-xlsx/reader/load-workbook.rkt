#lang racket

(provide (contract-out
          [load-workbook (-> path-string? xlsx? void?)]
          ))

(require simple-xml)

(require "../xlsx/xlsx.rkt")

(define (load-workbook workbook_file _xlsx)
    (let* ([xml_hash (xml->hash workbook_file)])

      (set-xlsx-sheet_count! _xlsx (hash-ref xml_hash "workbook.sheets.sheet's count" 0))

      (let loop ([index 1])
        (when (<= index (xlsx-sheet_count _xlsx))
              (let ([sheet_name (hash-ref xml_hash (format "workbook.sheets.sheet~a.name" index))]
                    [sheet_id (hash-ref xml_hash (format "workbook.sheets.sheet~a.sheetId" index))]
                    [rid (hash-ref xml_hash (format "workbook.sheets.sheet~a.r:id" index))])
                (hash-set! (xlsx-sheet_index_id_map _xlsx) (sub1 index) sheet_id)
                (hash-set! (xlsx-sheet_name_index_map _xlsx) sheet_name (sub1 index))
                (hash-set! (xlsx-sheet_index_name_map _xlsx) (sub1 index) sheet_name)
                (hash-set! (xlsx-sheet_index_rid_map _xlsx) (sub1 index) rid)
                )
              (loop (add1 index))))))
