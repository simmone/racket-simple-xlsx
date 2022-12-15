#lang racket

(require simple-xml)

(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")
(require "lib/lib.rkt")

(provide (contract-out
          [to-content-type (-> list?)]
          [from-content-type (-> path-string? void?)]
          [write-content-type (->* () (path-string?) void?)]
          [read-content-type (->* () (path-string?) void?)]
          ))

(define (header)
  '("Types"
    ("xmlns" . "http://schemas.openxmlformats.org/package/2006/content-types")
    ("Default" ("Extension" . "bin") ("ContentType" . "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"))
    ("Default" ("Extension" . "rels") ("ContentType" . "application/vnd.openxmlformats-package.relationships+xml"))
    ("Default" ("Extension" . "xml") ("ContentType" . "application/xml"))
    ("Override" ("PartName" . "/xl/workbook.xml") ("ContentType" . "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))
    ("Override" ("PartName" . "/xl/theme/theme1.xml") ("ContentType" . "application/vnd.openxmlformats-officedocument.theme+xml"))
    ("Override" ("PartName" . "/xl/styles.xml") ("ContentType" . "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"))
    ("Override" ("PartName" . "/docProps/core.xml") ("ContentType" . "application/vnd.openxmlformats-package.core-properties+xml"))
    ("Override" ("PartName" . "/docProps/app.xml") ("ContentType" . "application/vnd.openxmlformats-officedocument.extended-properties+xml"))))

(define (xlsx->content-type)
  (let loop ([loop_list (XLSX-sheet_list (*XLSX*))]
             [data_sheet_count 1]
             [chart_sheet_count 1]
             [xml_list '()])
    (if (not (null? loop_list))
        (let ([sheet (car loop_list)])
          (cond
           [(DATA-SHEET? sheet)
            (loop
             (cdr loop_list)
             (add1 data_sheet_count)
             chart_sheet_count
             (cons
              (list "Override"
                    (cons
                     "PartName"
                     (format "/xl/worksheets/sheet~a.xml" data_sheet_count))
                    (cons "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))
              xml_list))]
           [(CHART-SHEET? sheet)
            (loop
             (cdr loop_list)
             data_sheet_count
             (add1 chart_sheet_count)
             `(
               ,(list "Override"
                      (cons
                       "PartName"
                       (format "/xl/chartsheets/sheet~a.xml" chart_sheet_count))
                      (cons "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"))
               ,(list "Override"
                      (cons
                       "PartName"
                       (format "/xl/drawings/drawing~a.xml" chart_sheet_count))
                      (cons "ContentType" "application/vnd.openxmlformats-officedocument.drawing+xml"))
               ,(list "Override"
                      (cons
                       "PartName"
                       (format "/xl/charts/chart~a.xml" chart_sheet_count))
                      (cons "ContentType" "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))
               ,@xml_list))]
           [else
            (loop (cdr loop_list) data_sheet_count chart_sheet_count xml_list)]))
        (reverse xml_list))))

(define (xlsx->shared-string)
  (if (> (hash-count (XLSX-shared_string->index_map (*XLSX*))) 0)
      '("Override"
        ("PartName" . "/xl/sharedStrings.xml")
        ("ContentType" . "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))
      '()))

(define (to-content-type)
  `(
    ,@(header)
    ,@(xlsx->content-type)
    ,(xlsx->shared-string)))

(define (from-content-type content_type_file)
  (let* ([xml_hash (xml->hash content_type_file)]
         [types_override_count (hash-ref xml_hash "Types1.Override's count" 0)])

    (when (> types_override_count 0)
          (let loop ([loop_count 1]
                     [data_sheet_count 1]
                     [chart_sheet_count 1])
            (when (<= loop_count types_override_count)
                  (let ([part_name (hash-ref xml_hash (format "Types1.Override~a.PartName" loop_count) "")])
                    (cond
                     [(regexp-match #rx"worksheets" part_name)
                      (add-data-sheet (format "Sheet~a" data_sheet_count) '(("none")))
                      (loop (add1 loop_count) (add1 data_sheet_count) chart_sheet_count)]
                     [(regexp-match #rx"chartsheets" part_name)
                      (add-chart-sheet (format "Chart~a" chart_sheet_count) 'UNKNOWN "" '())
                      (loop (add1 loop_count) data_sheet_count (add1 chart_sheet_count))]
                     [else
                      (loop (add1 loop_count) data_sheet_count chart_sheet_count)])))))))

(define (write-content-type [output_dir #f])
  (let ([dir (if output_dir output_dir (XLSX-xlsx_dir (*XLSX*)))])
    (make-directory* dir)

    (with-output-to-file (build-path dir "[Content_Types].xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->xml (to-content-type)))))))

(define (read-content-type [output_dir #f])
  (let ([dir (if output_dir output_dir (XLSX-xlsx_dir (*XLSX*)))])
    (from-content-type (build-path dir "[Content_Types].xml"))))
