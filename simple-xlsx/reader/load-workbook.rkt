#lang racket

(provide (contract-out
          [load-workbook (-> path-string? (values list? hash? hash? hash?))]
          ))

(require simple-xml)

(define (load-workbook workbook_file)
  (let ([sheet_id_list '()]
        [sheet_id_name_map (make-hash)]
        [sheet_name_id_map (make-hash)]
        [sheet_id_rid_map (make-hash)])

    (let ([xml_hash (xml->hash workbook_file)])
      (let loop ([loop_count 1])
        (when (<= loop_count (hash-ref xml_hash "workbook.sheets.sheet's count" 0))
              (let ([sheet_name (hash-ref xml_hash (format "workbook.sheets.sheet~a.name" loop_count))]
                    [sheet_id (hash-ref xml_hash (format "workbook.sheets.sheet~a.sheetId" loop_count))]
                    [rid (hash-ref xml_hash (format "workbook.sheets.sheet~a.r:id" loop_count))])
                (set! sheet_id_list `(,@sheet_id_list ,sheet_id))
                (hash-set! sheet_id_rid_map sheet_id rid)
                (hash-set! sheet_name_id_map sheet_name sheet_id)
                (hash-set! sheet_id_name_map sheet_id sheet_name))
              (loop (add1 loop_count)))))
    (values
     sheet_id_list
     sheet_id_name_map
     sheet_name_id_map
     sheet_id_rid_map)))
