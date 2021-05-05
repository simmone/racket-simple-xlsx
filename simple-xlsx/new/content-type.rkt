#lang racket

(require simple-xml)

(require "../xlsx/xlsx.rkt")
(require "../sheet/sheet.rkt")
(require "../lib/lib.rkt")

(provide (contract-out
          [content-type (-> list?)]
          [write-content-type-file (-> void?)]
          [read-content-type (-> void?)]
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
  (let loop ([loop_list (XLSX-sheet_list (*CURRENT_XLSX*))]
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
                       (format "/xl/charts/chart~a.xml" chart_sheet_count))
                      (cons "ContentType" "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))

               ,(list "Override"
                      (cons
                       "PartName"
                       (format "/xl/drawings/drawing~a.xml" chart_sheet_count)
                       (cons "ContentType" "application/vnd.openxmlformats-officedocument.drawing+xml")))

               ,(list "Override"
                      (cons
                       "PartName"
                       (format "/xl/chartsheets/sheet~a.xml" chart_sheet_count)
                       (cons "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml")))
               
               ,@xml_list))]
           [else
            (loop (cdr loop_list) data_sheet_count chart_sheet_count xml_list)]))
        (reverse xml_list))))

(define (xlsx->shared-string)
  (if (> (hash-count (XLSX-shared_strings_map (*CURRENT_XLSX*))) 0)
      '("Override"
        ("PartName" . "/xl/sharedStrings.xml")
        ("ContentType" . "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"))
      '()))

(define (content-type)
  `(
    ,@(header)
    ,@(xlsx->content-type)
    ,(xlsx->shared-string)))

(define (write-content-type-file)
  (with-output-to-file (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "[Content_Types].xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (lists->compact_xml (content-type))))))

(define (read-content-type)
  (void))
