#lang racket

(require simple-xml)

(require "../../../xlsx/xlsx.rkt")
(require "../../../sheet/sheet.rkt")

(provide (contract-out
          [xl-rels (-> list?)]
          [write-workbook-rels (-> void?)]
          [read-workbook-rels-file (-> path-string? void?)]
          [read-workbook-rels (-> void?)]
          ))

(define (header)
  '("Relationships"
    ("xmlns" . "http://schemas.openxmlformats.org/package/2006/relationships")))

(define (rels)
  (let loop ([loop_list (XLSX-sheet_list (*CURRENT_XLSX*))]
             [count 1]
             [data_sheet_count 1]
             [chart_sheet_count 1]
             [xml_list '()])
    (if (not (null? loop_list))
        (cond
         [(DATA-SHEET? (car loop_list))
          (loop
           (cdr loop_list)
           (add1 count)
           (add1 data_sheet_count)
           chart_sheet_count
           (cons 
            (list "Relationship"
                  (cons "Id" (format "rId~a" count))
                  (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                  (cons "Target" (format "worksheets/sheet~a.xml" data_sheet_count)))
            xml_list))]
         [(CHART-SHEET? (car loop_list))
          (loop
           (cdr loop_list)
           (add1 count)
           data_sheet_count
           (add1 chart_sheet_count)
           (cons 
            (list "Relationship"
                  (cons "Id" (format "rId~a" count))
                  (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
                  (cons "Target" (format "chartsheets/sheet~a.xml" chart_sheet_count)))
            xml_list))]
         [else
          (loop (cdr loop_list) count data_sheet_count chart_sheet_count)])
        (reverse xml_list))))

(define (footer)
  (let ([seq (XLSX-sheet_count (*CURRENT_XLSX*))])
    (list
     (list "Relationship"
           (cons "Id" (format "rId~a" (add1 seq)))
           (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
           (cons "Target" "theme/theme1.xml"))
     (list "Relationship"
           (cons "Id" (format "rId~a" (+ seq 2)))
           (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
           (cons "Target" "styles.xml"))
     (if (> (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*))) 0)
         (list "Relationship"
               (cons "Id" (format "rId~a" (+ seq 3)))
               (cons "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
               (cons "Target" "sharedStrings.xml"))
         '()))))

(define (xl-rels)
  `(
    ,@(header)
    ,@(rels)
    ,@(footer)))

(define (write-workbook-rels)
  (let ([dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "xl" "_rels")])
    (make-directory* dir)

    (with-output-to-file (build-path dir "workbook.xml.rels")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->compact_xml (xl-rels)))))))

(define (read-workbook-rels-file rels_file)
  (let ([xml_hash (xml->hash rels_file)])

    (let ([relation_ship_count (hash-ref xml_hash "Relationships.Relationship's count" 0)])
      (let loop ([loop_count 1])
        (when (<= loop_count relation_ship_count)
              (let* (
                     [relation_ship_id (hash-ref xml_hash (format "Relationships.Relationship~a.Id" loop_count))]
                     [relation_ship_target (hash-ref xml_hash (format "Relationships.Relationship~a.Target" loop_count))]
                     )
                (hash-set! (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) relation_ship_id relation_ship_target)
                (hash-set! (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) (sub1 loop_count) relation_ship_target)

                (loop (add1 loop_count))))))))

(define (read-workbook-rels)
  (read-workbook-rels-file (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "workbook.xml.rels")))

