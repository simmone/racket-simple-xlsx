#lang racket

(provide (contract-out 
          [with-input-from-excel-file (-> path-string? (-> void?) void?)]
          [get-sheet-names (-> list?)]
          [get-cell-value (-> string? any)]
          [load-sheet (-> string? void?)]
          [get-sheet-dimension (-> pair?)]
          [with-row (-> (-> list? any) any)]
          ))

(require xml)

(require "lib/lib.rkt")

(define excel_dir (make-parameter "excel_dir"))

(define shared_map (make-parameter #f))

(define sheet_map (make-parameter #f))

(define sheet_name_map (make-parameter #f))

(define relation_name_map (make-parameter #f))

(define data_type_map (make-parameter #f))

(define dimension (make-parameter #f))

(define (with-input-from-excel-file excel_file user_proc)
  (with-unzip 
   excel_file
   (lambda (tmp_dir)
     (parameterize* ([excel_dir tmp_dir]
                     [shared_map (get-shared-string)]
                     [sheet_name_map (get-sheet-name-map)]
                     [relation_name_map (get-relation-name-map)])
                    (user_proc)))))

(define (get-sheet-name-map)
  (let ([data_map (make-hash)])
    (with-input-from-file 
        (build-path (excel_dir) "xl" "workbook.xml")
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
    (set! sheet_name_map (make-parameter data_map))
    data_map))

(define (get-relation-name-map)
  (let ([data_map (make-hash)])
    (with-input-from-file 
        (build-path (excel_dir) "xl" "_rels" "workbook.xml.rels")
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
    (set! relation_name_map (make-parameter data_map))
    data_map))

(define (get-sheet-names)
  (map
   (lambda (item)
     (car item))
   (sort #:key cdr (hash->list (sheet_name_map)) string<?)))

(define (get-shared-string)
  (let ([data_map (make-hash)])
    (with-input-from-file 
        (build-path (excel_dir) "xl" "sharedStrings.xml")
      (lambda ()
        (let loop1 ([line (read-line)])
          (when (not (eof-object? line))
                (let loop2 ([split_items1 (regexp-split #rx"</si>" line)]
                            [index 0])
                  (when (not (null? split_items1))
                        (let ([split_items2 (regexp-split #rx"<si>" (car split_items1))])
                          (when (> (length split_items2) 1)
                                (let* ([xml (xml->xexpr (document-element (read-xml (open-input-string (string-append "<si>" (second split_items2) "</si>")))))]
                                       [v_list (xml-get-list 't xml)])
                                  (hash-set! data_map 
                                             (number->string index) 
                                             (regexp-replace* #rx" " (foldr (lambda (a b) (string-append a b)) "" v_list) " ")))))
                        (loop2 (cdr split_items1) (add1 index))))
                (loop1 (read-line))))))
  data_map))

(define (load-sheet sheet_name)
  (let ([data_map (make-hash)]
        [type_map (make-hash)])
    (with-input-from-file 
        (build-path (excel_dir) "xl" (hash-ref (relation_name_map) (hash-ref (sheet_name_map) sheet_name)))
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
            (set! dimension (make-parameter (cons (string->number row_str) (abc->number col_str)))))
          )))
    (set! sheet_map (make-parameter data_map))
    (set! data_type_map (make-parameter type_map))))

(define (get-cell-value item_name)
  (if (hash-has-key? (sheet_map) item_name)
      (let* ([type (hash-ref (data_type_map) item_name)]
             [type_t (car type)]
             [type_s (cdr type)]
             [value (hash-ref (sheet_map) item_name)])
      (cond
       [(string=? type_t "s")
        (hash-ref (shared_map) value)]
       [(string=? type_s "5")
        (format-time (string->number value))]
       [(or (string=? type_t "") (string=? type_s "2"))
        (string->number value)]))
      ""))

(define (get-sheet-dimension)
  (dimension))
  
(define (with-row proc)
  (let ([rows (car (dimension))]
        [cols (cdr (dimension))])
    (let loop ([row_index 1])
      (when (<= row_index rows)
            (proc (map
                   (lambda (col)
                     (get-cell-value (string-append (number->abc col) (number->string row_index))))
                   (number->list cols)))
            (loop (add1 row_index))))))
