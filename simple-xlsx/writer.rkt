#lang racket

(provide (contract-out
          [write-xlsx-file (-> path-string? void?)]
          [add-data-sheet (-> string? (listof list?) void?)]
          [add-chart-sheet (-> string? (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D) string? void?)]
          ))

(require racket/date)

(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")
(require "lib/lib.rkt")
(require "lib/dimension.rkt")
(require "new/content-type.rkt")
(require "new/_rels/rels.rkt")
(require "new/docProps/docprops-app.rkt")
(require "new/docProps/docprops-core.rkt")

(define (write-xlsx-file xlsx_file_name)
  (when (file-exists? xlsx_file_name)
        (delete-file xlsx_file_name))

  (dynamic-wind
      (lambda () (set-XLSX-xlsx_dir! (*CURRENT_XLSX*) (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
      (lambda ()
        ;; [Content_Types].xml
        (write-content-type-file)

        ;; _rels
        (write-rels-file)

        ;; docProps-app
        (write-docprops-app-file)

        ;; docProps-core
        (write-docprops-core-file)

        (zip-xlsx xlsx_file_name (XLSX-xlsx_dir (*CURRENT_XLSX*))))
      (lambda ()
        (delete-directory/files (XLSX-xlsx_dir (*CURRENT_XLSX*))))))

(define (add-data-sheet sheet_name sheet_data)
  (check-data-integrity sheet_data)

  (if (not (hash-has-key? (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name))
      (let ([seq (add1 (length (XLSX-sheet_list (*CURRENT_XLSX*))))]
            [type_seq (add1 (length (filter (lambda (rec) (DATA-SHEET? rec)) (XLSX-sheet_list (*CURRENT_XLSX*)))))]
            [shared_string_index (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*)))]
            [rvtsf_map (make-hash)])
        (let row-loop ([rows sheet_data]
                       [row_index 1])
          (when (not (null? rows))
                (let col-loop ([cols (car rows)]
                               [col_index 1])
                  (when (not (null? cols))
                        (cond
                         [(string? (car cols))
                          (set! shared_string_index (add-shared-strings-map (XLSX-shared_strings_map (*CURRENT_XLSX*)) (car cols) shared_string_index))
                          (hash-set! rvtsf_map
                                     (row_col->dimension row_index col_index) ;; r
                                     (list
                                      shared_string_index ;; v t s f
                                      "s"
                                      #f
                                      #f))
                          shared_string_index]
                         [(date? (car cols))
                          (let ([date_v (date->oa_date_number (car cols))])
                            (hash-set! rvtsf_map
                                       (row_col->dimension row_index col_index) ;; r
                                       (list
                                        date_v ;; v t s f
                                        "n"
                                        #f
                                        #f))
                            date_v)]
                         [(number? (car cols))
                          (hash-set! rvtsf_map
                                     (row_col->dimension row_index col_index) ;; r
                                     (list
                                      (car cols) ;; v t s f
                                      "n"
                                      #f
                                      #f))
                          (car cols)])

                      (col-loop (cdr cols) (add1 col_index))))

                (row-loop (cdr rows) (add1 row_index))))

        (let* ([sheet_index (XLSX-sheet_count (*CURRENT_XLSX*))]
               [id (number->string (add1 sheet_index))]
               [rId (format "rId~a" id)]
               [rel (format "worksheets/sheet~a" id)])

          (set-XLSX-sheet_count! (*CURRENT_XLSX*) (add1 (XLSX-sheet_count (*CURRENT_XLSX*))))
          
          (set-XLSX-sheet_list! (*CURRENT_XLSX*)
                                `(,@(XLSX-sheet_list (*CURRENT_XLSX*))
                                  ,(new-data-sheet
                                    (get-dimension sheet_data)
                                    rvtsf_map)))

          (hash-set! (XLSX-sheet_index_id_map (*CURRENT_XLSX*)) sheet_index id)
          (hash-set! (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) sheet_index sheet_name)
          (hash-set! (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name sheet_index)
          (hash-set! (XLSX-sheet_index_rid_map (*CURRENT_XLSX*)) sheet_index rId)
          (hash-set! (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) rId rel)
          (hash-set! (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index rel)))
      (error (format "duplicate sheet name[~a]" sheet_name))))

(define (add-chart-sheet sheet_name chart_type topic)
  (if (not (hash-has-key? (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name))
        (let* ([sheet_index (XLSX-sheet_count (*CURRENT_XLSX*))]
               [id (number->string (add1 sheet_index))]
               [rId (format "rId~a" id)]
               [rel (format "chartsheets/sheet~a" id)])

          (set-XLSX-sheet_count! (*CURRENT_XLSX*) (add1 sheet_index))
          (set-XLSX-sheet_list! (*CURRENT_XLSX*) `(,@(XLSX-sheet_list (*CURRENT_XLSX*)) ,(new-chart-sheet chart_type topic)))
          (hash-set! (XLSX-sheet_index_id_map (*CURRENT_XLSX*)) sheet_index id)
          (hash-set! (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) sheet_index sheet_name)
          (hash-set! (XLSX-sheet_name_index_map (*CURRENT_XLSX*)) sheet_name sheet_index)
          (hash-set! (XLSX-sheet_index_rid_map (*CURRENT_XLSX*)) sheet_index rId)
          (hash-set! (XLSX-sheet_rid_rel_map (*CURRENT_XLSX*)) rId rel)
          (hash-set! (XLSX-sheet_index_rel_map (*CURRENT_XLSX*)) sheet_index rel))
      (error (format "duplicate sheet name[~a]" sheet_name))))

