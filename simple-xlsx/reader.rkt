#lang racket

(provide (contract-out
          [read-xlsx% class?]
          [with-input-from-xlsx-file (-> path-string? (-> (is-a?/c read-xlsx%) void?) void?)]
          [get-sheet-names (-> (is-a?/c read-xlsx%) list?)]
          [get-cell-value (-> string? (is-a?/c read-xlsx%) any)]
          [get-sheet-dimension (-> (is-a?/c read-xlsx%) pair?)]
          [load-sheet (-> string? (is-a?/c read-xlsx%) void?)]
          [load-sheet-ref (-> exact-nonnegative-integer? (is-a?/c read-xlsx%) void?)]
          [get-sheet-rows (-> (is-a?/c read-xlsx%) list?)]
          [sheet-name-rows (-> path-string? string? list?)]
          [sheet-ref-rows (-> path-string? exact-nonnegative-integer? list?)]
          ))

(require xml)

(require "lib/lib.rkt")

(define read-xlsx%
  (class object%
         (init-field [xlsx_dir ""]
                     [shared_map #f]
                     [sheet_map #f]
                     [sheet_name_map #f]
                     [relation_name_map #f]
                     [data_type_map #f]
                     [dimension #f])
         (super-new)))

(define (with-input-from-xlsx-file xlsx_file user_proc)
  (with-unzip
   xlsx_file
   (lambda (tmp_dir)
     (let ([new_shared_map #f]
           [new_sheet_name_map #f]
           [new_relation_name_map #f]
           [xlsx_obj #f])
     (set! new_shared_map (get-shared-string tmp_dir))

     (set! new_sheet_name_map (get-sheet-name-map tmp_dir))

     (set! new_relation_name_map (get-relation-name-map tmp_dir))

     (set! xlsx_obj
           (new read-xlsx%
                 (xlsx_dir tmp_dir)
                 (shared_map new_shared_map)
                 (sheet_name_map new_sheet_name_map)
                 (relation_name_map new_relation_name_map)))
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
          (let loop ([loop_list 
                      (xml-get-list 'Relationships xml)])
            (when (not (null? loop_list))
                  (if (not (list? (car loop_list)))
                      (loop (cdr loop_list))
                      (let ([attr_list (cadr (car loop_list))]
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
                      (loop (cdr loop_list)))))))
      data_map))

(define (get-sheet-names xlsx)
  (map
   (lambda (item)
     (car item))
   (sort #:key cdr (hash->list (get-field sheet_name_map xlsx)) string<?)))

(define (get-shared-string xlsx_dir)
  (let ([data_map (make-hash)]
        [shared_strings_file_name (build-path xlsx_dir "xl" "sharedStrings.xml")])
    (when (file-exists? shared_strings_file_name)
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
                                       [v_list (xml-get-list 't xml)]
                                       [ignore_rPh_map (make-hash)])
                                  
                                  (for-each
                                   (lambda (rPh_rec)
                                     (hash-set! ignore_rPh_map (third rPh_rec) ""))
                                   (xml-get-list 'rPh xml))
                                  
                                  (hash-set! data_map
                                             (number->string index)
                                             (regexp-replace* 
                                              #rx" " 
                                              (foldr (lambda (a b) (string-append a b)) "" 
                                                     (filter 
                                                      (lambda (item) 
                                                        (not (hash-has-key? ignore_rPh_map item))) 
                                                      (map
                                                       (lambda (v)
                                                         (cond
                                                          [(string? v)
                                                           v]
                                                          [(integer? v)
                                                           (string (integer->char v))]
                                                          [else
                                                           ""]))
                                                         v_list)))
                                              " ")))))
                        (loop2 (cdr split_items1) (add1 index))))))))
  data_map))

(define (load-sheet sheet_name xlsx)
  (let ([data_map (make-hash)]
        [type_map (make-hash)]
        [dimension_col 0]
        [dimension ""]
        [rows #f]
        [data_sheet_file_name
         (build-path (get-field xlsx_dir xlsx) "xl" (hash-ref (get-field relation_name_map xlsx) (hash-ref (get-field sheet_name_map xlsx) sheet_name)))])
    
    (when (string=? (path->string (fourth (explode-path data_sheet_file_name))) "worksheets")
          (let ([file_str (file->string data_sheet_file_name)])
            (set! rows
                  (let loop ([loop_list
                              (regexp-split #rx"<sheetData>|</sheetData>|<row" file_str)]
                             [result_list '()]
                             [index 1])
                    (if (not (null? loop_list))
                        (begin
                          (if (regexp-match #rx"</row>" (car loop_list))
                              (let ([row_index (second (regexp-match #rx" r=\"([0-9]+)\" " (car loop_list)))]
                                    [col_info (regexp-match* #rx" r=\"([A-Z]+)[0-9]+\" *" (car loop_list))])
                                (if (= (string->number row_index) index)
                                    (let ([row (xml->xexpr (document-element (read-xml (open-input-string (string-append "<row" (car loop_list))))))])
                                      (when (not (null? col_info))
                                            (let ([max_col_index (abc->number 
                                                                  (car (reverse 
                                                                        (map 
                                                                         (lambda (item)
                                                                           (second (regexp-match #rx"([A-Z]+)" item)))
                                                                         col_info))))])
                                              (when (> max_col_index dimension_col)
                                                    (set! dimension_col max_col_index))))
                                      (loop
                                       (cdr loop_list)
                                       (cons
                                        (xml->xexpr (document-element (read-xml (open-input-string (string-append "<row" (car loop_list)))))) 
                                        result_list)
                                       (add1 index)))
                                    (loop
                                     loop_list
                                     (cons null result_list)
                                     (add1 index))))
                              (loop (cdr loop_list) result_list index)))
                        (reverse result_list)))))
          
          (set-field! dimension xlsx (cons (length rows) dimension_col))
          
          (for-each
           (lambda (row_xml)
             (when (not (null? row_xml))
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
                                                  (set! para_t value)]
                                                 ))
                                              (loop (cdr para_list))))
                                      (hash-set! type_map para_r (cons para_t para_s))

                                      (let loop-cell ([cell_list (cdr cell_item)])
                                        (when (not (null? cell_list))
                                              (when (equal? (caar cell_list) 'v)
                                                    (hash-set! data_map para_r (caddar cell_list)))
                                              (loop-cell (cdr cell_list))))
                                      )))))
                    row_xml)))
           rows)
          )

    (set-field! sheet_map xlsx data_map)
    (set-field! data_type_map xlsx type_map)))

(define (load-sheet-ref sheet_index xlsx)
  (load-sheet (list-ref (get-sheet-names xlsx) sheet_index) xlsx))

(define (get-cell-value item_name xlsx)
  (let ([sheet_map (get-field sheet_map xlsx)]
        [data_type_map (get-field data_type_map xlsx)]
        [shared_map (get-field shared_map xlsx)])
    (if (and
         (hash-has-key? sheet_map item_name)
         (not (null? (hash-ref data_type_map item_name))))
        (let* ([type (hash-ref data_type_map item_name)]
               [type_t (car type)]
               [type_s (cdr type)]
               [value (hash-ref sheet_map item_name)])
          (cond
           [(string=? type_t "s")
            (hash-ref shared_map value)]
           [(string=? type_t "n")
            (string->number value)]
           [(string=? type_t "")
            (string->number value)]))
        "")))

(define (get-sheet-dimension xlsx)
  (get-field dimension xlsx))

(define (get-sheet-rows xlsx)
  (let* ([dimension (get-field dimension xlsx)]
         [rows (car dimension)]
         [cols (cdr dimension)])
    (let loop ([row_index 1]
               [result_list '()])
      (if (<= row_index rows)
          (loop 
           (add1 row_index)
           (cons 
            (map
             (lambda (col)
               (get-cell-value (string-append (number->abc col) (number->string row_index)) xlsx))
             (number->list cols))
            result_list))
          (reverse result_list)))))

(define (sheet-name-rows xlsx_file_path sheet_name)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet sheet_name xlsx)
     
     (get-sheet-rows xlsx))))

(define (sheet-ref-rows xlsx_file_path sheet_index)
  (with-input-from-xlsx-file
   xlsx_file_path
   (lambda (xlsx)
     (load-sheet-ref sheet_index xlsx)
     
     (get-sheet-rows xlsx))))

