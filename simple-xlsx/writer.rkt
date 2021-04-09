#lang racket

(provide (contract-out
          [write-xlsx-file (-> (is-a?/c xlsx%) path-string? void?)]
          [add-data-sheet! (-> XLSX? string? (listof list?) void?)]
          ))

(require racket/date)

(require "xlsx/xlsx.rkt")
(require "xlsx/sheet.rkt")
(require "lib/lib.rkt")
(require "writer/content-type/content-type.rkt")
(require "writer/_rels/rels.rkt")
(require "writer/docProps/docprops-app.rkt")
(require "writer/docProps/docprops-core.rkt")
(require "writer/xl/_rels/workbook-xml-rels.rkt")
(require "writer/xl/printerSettings/printerSettings.rkt")
(require "writer/xl/sharedStrings.rkt")
(require "writer/xl/styles/styles.rkt")
(require "writer/xl/theme/theme.rkt")
(require "writer/xl/workbook.rkt")
(require "writer/xl/worksheets/_rels/rels.rkt")
(require "writer/xl/worksheets/worksheet/worksheet.rkt")
(require "writer/xl/charts/chart.rkt")
(require "writer/xl/chartsheets/chartsheet.rkt")
(require "writer/xl/chartsheets/_rels/chartsheet-rels.rkt")
(require "writer/xl/drawings/drawing.rkt")
(require "writer/xl/drawings/_rels/drawing-rels.rkt")

(define (write-xlsx-file xlsx xlsx_file_name)
  (when (file-exists? xlsx_file_name)
        (delete-file xlsx_file_name))

  (let ([tmp_dir #f])
    (dynamic-wind
        (lambda () (set! tmp_dir (make-temporary-file "xlsx_tmp_~a" 'directory ".")))
        (lambda ()
          ;; [Content_Types].xml
          (write-content-type-file tmp_dir xlsx)

          ;; _rels
          (write-rels-file (build-path tmp_dir "_rels"))

          ;; docProps
          (write-docprops-app-file (build-path tmp_dir "docProps") xlsx)
          (write-docprops-core-file (build-path tmp_dir "docProps") (current-date))
                
          ;; xl
          (write-workbook-xml-rels-file (build-path tmp_dir "xl" "_rels") xlsx)

          ;; printerSettings
          (create-printer-settings (build-path tmp_dir "xl" "printerSettings") xlsx)
                  
          ;; theme
          (write-theme-file (build-path tmp_dir "xl" "theme"))

          ;; sharedStrings
          (write-shared-strings-file (build-path tmp_dir "xl") xlsx)

          ;; styles
          (send xlsx burn-styles!)
          (write-styles-file 
           (build-path tmp_dir "xl") 
           (send xlsx get-style-list) 
           (send xlsx get-fill-list) 
           (send xlsx get-font-list)
           (send xlsx get-numFmt-list)
           (send xlsx get-border-list)
           )

          ;; workbook
          (write-workbook-file (build-path tmp_dir "xl") xlsx)

          ;; data-sheets
          (write-worksheets-rels-file (build-path tmp_dir "xl" "worksheets" "_rels") xlsx)
          (write-data-sheet-file (build-path tmp_dir "xl" "worksheets") xlsx)

          ;; charts
          (write-chart-file (build-path tmp_dir "xl" "charts") xlsx)

          ;; chartsheets
          (write-chart-sheet-rels-file (build-path tmp_dir "xl" "chartsheets" "_rels") xlsx)
          (write-chart-sheet-file (build-path tmp_dir "xl" "chartsheets") xlsx)

          ;; drawing
          (write-drawing-rels-file (build-path tmp_dir "xl" "drawings" "_rels") xlsx)
          (write-drawing-file (build-path tmp_dir "xl" "drawings") xlsx)

          (zip-xlsx xlsx_file_name tmp_dir))
        (lambda ()
          (delete-directory/files tmp_dir)))))

(define (add-data-sheet! xlsx sheet_name sheet_data)
  (check-data-integrity sheet_data)

  (if (not (hash-has-key? (XLSX-sheet_name_index_map xlsx) sheet_name))
      (let* ([seq (add1 (length (XLSX-sheet_list xlsx)))]
             [type_seq (add1 (length (filter (lambda (rec) (DATA-SHEET? rec)) (XLSX-sheet_list xlsx))))]
             [shared_string_index (hash-cont (XLSX-shared_strings_map xlsx))]
             [transformed_sheet_data 
              (let row-loop ([rows sheet_data]
                             [row_result '()])
                (if (not (null? rows))
                    (row-loop
                     (cdr loop_list)
                     (cons
                      (let col-loop ([cols (car rows)]
                                     [col_result '()])
                        (if (not (null? cols))
                            (col-loop
                             (cdr cols)
                             (cons
                              (cond
                               [(string? (car cols))
                                (set! shared_string_index (add-shared-strings-map (XLSX-shared_strings_map xlsx) (car cols) shared_string_index))
                                shared_string_index]
                               [(date? (car cols))
                                (date->oa_date_number (car cols))]
                               [else
                                (car cols)])
                              col_result))
                            (reverse col_result)))
                      row_result))
                    (reverse row_result)))])
        
        (set-XLSX-sheet_count! xlsx (add1 (XLSX-sheet_count xlsx)))

        (set-XLSX-sheet_list! xlsx
                              `(,@(XLSX-sheet_list xlsx)
                                ,(new-data-sheet 
                                  transformed_sheet_data
                                  (make-hash) (make-hash) '(0 . 0) (make-hash) (make-hash) (make-hash) (make-hash) (make-hash) (make-hash))))
        (hash-set! sheet_name_map sheet_name (sub1 seq)))
      (error (format "duplicate sheet name[~a]" sheet_name))))

                   (sheet_count natural?)
                   (sheet_index_id_map (hash/c natural? string?))
                   (sheet_index_name_map (hash/c natural? string?))
                   (sheet_name_index_map (hash/c string? natural?))
                   (sheet_index_rid_map (hash/c natural? string?))
                   (sheet_rid_rel_map (hash/c string? string?))
                   (sheet_index_rel_map (hash/c natural? string?))
                   (shared_strings_map (hash/c natural? string?))
                   (sheet_list (listof (or/c DATA-SHEET? CHART-SHEET?)))


