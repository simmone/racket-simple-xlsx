#lang racket

(provide (contract-out
          [load-workbook-rels (-> path-string? XLSX? void?)]
          ))

(require simple-xml)

(require "../xlsx/xlsx.rkt")

(define (load-workbook-rels workbook_relation_file _xlsx)
  (let ([xml_hash (xml->hash workbook_relation_file)])

    (let ([relation_ship_count (hash-ref xml_hash "Relationships.Relationship's count" 0)])
      (let loop ([loop_count 1])
        (when (<= loop_count relation_ship_count)
              (let* (
                     [relation_ship_id (hash-ref xml_hash (format "Relationships.Relationship~a.Id" loop_count))]
                     [relation_ship_target (hash-ref xml_hash (format "Relationships.Relationship~a.Target" loop_count))]
                     )
                (hash-set! (XLSX-sheet_rid_rel_map _xlsx) relation_ship_id relation_ship_target)

                (loop (add1 loop_count))))))))
