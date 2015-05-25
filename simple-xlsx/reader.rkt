#lang racket

(provide (contract-out 
          [with-input-from-xlsx-file (-> path-string? (-> any/c void?) void?)]
          [get-sheet-names (-> any/c list?)]
          [get-cell-value (-> string? any/c any)]
          [load-sheet (-> string? any/c void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? any/c void?)]          
          [get-sheet-dimension (-> any/c pair?)]
          [with-row (-> any/c (-> list? any) any)]
          ))

(require xml)

(require "lib/lib.rkt")

(define xlsx%
  (class object%
         (init-field [xlsx_dir ""] [shared_map #f] [sheet_map #f] [sheet_name_map #f] [relation_name_map #f] [data_type_map #f] [dimension #f])
         (super-new)))

(define (with-input-from-xlsx-file xlsx_file user_proc)
  (with-unzip 
   xlsx_file
   (lambda (tmp_dir)
     (let ([xlsx_obj
            (new xlsx% (xlsx_dir tmp_dir) (shared_map (get-shared-string tmp_dir)) (sheet_name_map (get-sheet-name-map tmp_dir)) (relation_name_map (get-relation-name-map tmp_dir)))])
       (user_proc xlsx_obj)))))

(define (get-sheet-name-map xlsx_dir)
  (let ([data_map (make-hash)])
    (with-input-from-file 
        (build-path xlsx_dir "xl" "workbook.xml")
      (lambda ()
        (let ([xml (xml->xexpr (document-element (read-xml (current-input-port))))]
              [sheet_list '()])
          (for-each
           (lambda (sheet)
             (let ([attr_list (cadr sheet)]
                   [sheet_name ""]
                   [sheet_id ""])
               (for-each
                (lambda (attr_pair)
                  (let ([name (car attr_pair)]
                        [value (cadr attr_pair)])
                    (cond
                     [(equal? name 'name)
                      (set! sheet_name value)]
                     [(equal? name 'r:id)
                      (set! sheet_id value)])))
                attr_list)
               (hash-set! data_map sheet_name sheet_id)))
           (xml-get-list 'sheets xml)))))
    data_map))

(define (get-relation-name-map xlsx_dir)
  (let ([data_map (make-hash)])
    (with-input-from-file 
        (build-path xlsx_dir "xl" "_rels" "workbook.xml.rels")
      (lambda ()
        (let ([xml (xml->xexpr (document-element (read-xml (current-input-port))))]
              [sheet_list '()])
          (for-each
           (lambda (sheet)
             (let ([attr_list (cadr sheet)]
                   [sheet_name ""]
                   [sheet_id ""])
               (for-each
                (lambda (attr_pair)
                  (let ([name (car attr_pair)]
                        [value (cadr attr_pair)])
                    (cond
                     [(equal? name 'Id)
                      (set! sheet_name value)]
                     [(equal? name 'Target)
                      (set! sheet_id value)])))
                attr_list)
               (hash-set! data_map sheet_name sheet_id)))
           (xml-get-list 'Relationships xml)))))
    data_map))

(define (get-sheet-names xlsx)
  (map
   (lambda (item)
     (car item))
   (sort #:key cdr (hash->list (get-field sheet_name_map xlsx)) string<?)))

(define (get-shared-string xlsx_dir)
  (let ([data_map (make-hash)])
    (with-input-from-file 
        (build-path xlsx_dir "xl" "sharedStrings.xml")
      (lambda ()
        (let ([xml_str (port->string (current-input-port))])
          (let loop2 ([split_items1 (regexp-split #rx"</si>" xml_str)]
                      [index 0])
            (when (not (null? split_items1))
                  (let ([split_items2 (regexp-split #rx"<si>" (car split_items1))])
                    (when (> (length split_items2) 1)
                          (let* ([xml (xml->xexpr (document-element (read-xml (open-input-string (string-append "<si>" (second split_items2) "</si>")))))]
                                 [v_list (xml-get-list 't xml)])
                            (hash-set! data_map 
                                       (number->string index) 
                                       (regexp-replace* #rx" " (foldr (lambda (a b) (string-append a b)) "" v_list) " ")))))
                  (loop2 (cdr split_items1) (add1 index)))))))
  data_map))

(define (load-sheet sheet_name xlsx)
  (let ([data_map (make-hash)]
        [type_map (make-hash)])
    (with-input-from-file 
        (build-path (get-field xlsx_dir xlsx) "xl" (hash-ref (get-field relation_name_map xlsx) (hash-ref (get-field sheet_name_map xlsx) sheet_name)))
      (lambda ()
        (let* ([xml (xml->xexpr (document-element (read-xml (current-input-port))))]
               [v_list (xml-get-list 'sheetData xml)]
               [dimension_list (xml-get-list 'dimension xml)])
          (for-each
           (lambda (row_items)
             (for-each
              (lambda (cell_item)
                (when (list? cell_item)
                      (let ([first_item (car cell_item)])
                        (when (and (symbol? first_item) (equal? first_item 'c))
                              (let ([para_part (second cell_item)]
                                    [para_r ""]
                                    [para_s ""]
                                    [para_t ""])
                                (let loop ([para_list para_part])
                                  (when (not (null? para_list))
                                        (let* ([para (car para_list)]
                                               [key (car para)]
                                               [value (cadr para)])
                                          (cond
                                           [(equal? key 'r)
                                            (set! para_r value)]
                                           [(equal? key 's)
                                            (set! para_s value)]
                                           [(equal? key 't)
                                            (set! para_t value)]))
                                        (loop (cdr para_list))))
                                (hash-set! type_map para_r (cons para_t para_s))

                                (let loop-cell ([cell_list (cdr cell_item)])
                                  (when (not (null? cell_list))
                                        (when (equal? (caar cell_list) 'v)
                                              (hash-set! data_map para_r (caddar cell_list)))
                                        (loop-cell (cdr cell_list))))
                                )))))
              row_items))
           v_list)
          
          (let* ([dimension_str (xml-get-attr 'dimension "ref" xml)]
                 [dimension_items (regexp-split #rx":" dimension_str)]
                 [dest_item (list-ref dimension_items (sub1 (length dimension_items)))]
                 [items (regexp-match (regexp "([A-Z]+)([0-9]+)") dest_item)]
                 [col_str (cadr items)]
                 [row_str (caddr items)])
            (set-field! dimension xlsx (cons (string->number row_str) (abc->number col_str))))
          )))
    (set-field! sheet_map xlsx data_map)
    (set-field! data_type_map xlsx type_map)))

(define (load-sheet-ref sheet_index xlsx)
  (load-sheet (list-ref (get-sheet-names xlsx) sheet_index) xlsx))

(define (get-cell-value item_name xlsx)
  (if (hash-has-key? (get-field sheet_map xlsx) item_name)
      (let* ([type (hash-ref (get-field data_type_map xlsx) item_name)]
             [type_t (car type)]
             [type_s (cdr type)]
             [value (hash-ref (get-field sheet_map xlsx) item_name)])
      (cond
       [(string=? type_t "s")
        (hash-ref (get-field shared_map xlsx) value)]
       [(string=? type_s "5")
        (format-time (string->number value))]
       [(or (string=? type_t "") (string=? type_s "2"))
        (string->number value)]))
      ""))

(define (get-sheet-dimension xlsx)
  (get-field dimension xlsx))
  
(define (with-row xlsx proc)
  (let ([rows (car (get-field dimension xlsx))]
        [cols (cdr (get-field dimension xlsx))])
    (let loop ([row_index 1])
      (when (<= row_index rows)
            (proc (map
                   (lambda (col)
                     (get-cell-value (string-append (number->abc col) (number->string row_index)) xlsx))
                   (number->list cols)))
            (loop (add1 row_index))))))
