#lang racket

(require simple-xml)

(require "../xlsx/xlsx.rkt")
(require "../sheet/sheet.rkt")

(provide (contract-out
          [to-workbook (-> list?)]
          [from-workbook (-> path-string? void?)]
          [write-workbook (->* () (path-string?) void?)]
          [read-workbook (->* () (path-string?) void?)]
          ))

(define (to-workbook)
  (append
   '("workbook"
     ("xmlns" . "http://schemas.openxmlformats.org/spreadsheetml/2006/main") ("xmlns:r" . "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
     ("fileVersion" ("appName" . "xl") ("lastEdited" . "4") ("lowestEdited" . "4") ("rupBuild" . "4505"))
     ("workbookPr" ("filterPrivacy" . "1") ("defaultThemeVersion" . "124226"))
     ("bookViews"
      ("workbookView" ("xWindow" . "0") ("yWindow" . "90") ("windowWidth" . "19200") ("windowHeight" . "10590"))))
   (list
    (append
     '("sheets")
     (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
                [index 0]
                [result_list '()])
       (if (not (null? sheets))
           (let ([sheet (car sheets)])
             (loop
              (cdr sheets)
              (add1 index)
              (cons
               (list
                "sheet"
                (cons "name" (get-sheet-name sheet))
                (cons "sheetId" (number->string (add1 index)))
                (cons "r:id" (format "rId~a" (add1 index))))
               result_list)))
           (reverse result_list)))))
  '(("calcPr" ("calcId" . "124519")))))

(define (from-workbook workbook_file)
  (when (file-exists? workbook_file)
    (let* ([xml_hash (xml->hash workbook_file)]
           [workbook_count (hash-ref xml_hash "workbook1.sheets1.sheet's count" 0)]
           [sheet_list (XLSX-sheet_list (*XLSX*))])

      (when (> workbook_count 0)
        (let loop ([loop_count 1])
          (when (<= loop_count workbook_count)
            (let ([sheet_name (hash-ref xml_hash (format "workbook1.sheets1.sheet~a.name" loop_count) "")])
              (cond
               [(DATA-SHEET? (list-ref sheet_list (sub1 loop_count)))
                (set-DATA-SHEET-sheet_name! (list-ref sheet_list (sub1 loop_count)) sheet_name)]
               [(CHART-SHEET? (list-ref sheet_list (sub1 loop_count)))
                (set-CHART-SHEET-sheet_name! (list-ref sheet_list (sub1 loop_count)) sheet_name)])
              (loop (add1 loop_count)))))))))

(define (write-workbook [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (make-directory* dir)

    (with-output-to-file (build-path dir "workbook.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (lists->xml (to-workbook)))))))

(define (read-workbook [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (from-workbook (build-path dir "workbook.xml"))))

