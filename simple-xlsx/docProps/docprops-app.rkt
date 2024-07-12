#lang racket

(require fast-xml
         "../xlsx/xlsx.rkt"
         "../sheet/sheet.rkt")

(provide (contract-out
          [to-docprops-app (-> list?)]
          [from-docprops-app (-> path-string? void?)]
          [write-docprops-app (->* () (path-string?) void?)]
          [read-docprops-app (->* () (path-string?) void?)]
          ))

(define (to-docprops-app)
  (let ([properties_list
         '("Properties"
           ("xmlns" . "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
           ("xmlns:vt" . "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
           ("Application" "Microsoft Excel")
           ("DocSecurity" "0")
           ("ScaleCrop" "false")
           ("LinksUpToDate" "false")
           ("SharedDoc" "false")
           ("HyperlinksChanged" "false")
           ("AppVersion" "12.0000"))]
        [heading_pairs_list #f]
        [titles_of_parts_list #f]
        [data_sheet_count (length (filter DATA-SHEET? (XLSX-sheet_list (*XLSX*))))]
        [chart_sheet_count (length (filter CHART-SHEET? (XLSX-sheet_list (*XLSX*))))])

    (set! heading_pairs_list
          `("HeadingPairs"
            ,(append
              (list
               "vt:vector"
               (cons "size" (number->string
                             (+ (if (> data_sheet_count 0) 2 0) (if (> chart_sheet_count 0) 2 0))))
               (cons "baseType" "variant"))

                (if (> data_sheet_count 0)
                    (list
                     (list "vt:variant" (list "vt:lpstr" "工作表"))
                     (list "vt:variant" (list "vt:i4" (number->string data_sheet_count))))
                    '())

                (if (> chart_sheet_count 0)
                    (list
                     (list "vt:variant" (list "vt:lpstr" "图表"))
                     (list "vt:variant" (list "vt:i4" (number->string chart_sheet_count))))
                    '()))))

    (set! titles_of_parts_list
          `("TitlesOfParts"
            ,(append
              (list "vt:vector"
                    (cons "size" (number->string (length (XLSX-sheet_list (*XLSX*)))))
                    (cons "baseType" "lpstr"))

              (let loop ([sheets (XLSX-sheet_list (*XLSX*))]
                         [index 0]
                         [result_list '()])
                (if (not (null? sheets))
                    (loop (cdr sheets) (add1 index) (cons (list "vt:lpstr" (list-ref (get-sheet-name-list) index)) result_list))
                    (reverse result_list))))))

    `(,@properties_list ,heading_pairs_list ,titles_of_parts_list)))

(define (from-docprops-app doc_props_app_file)
  (when (file-exists? doc_props_app_file)
    (let* ([xml_hash (xml-file-to-hash
                      doc_props_app_file
                      '(
                        "Properties.TitlesOfParts.vt:vector.vt:lpstr"
                        ))]
           [sheet_list (XLSX-sheet_list (*XLSX*))])
      
      (let ([lpstr_count (hash-ref xml_hash "Properties1.TitlesOfParts1.vt:vector1.vt:lpstr's count" 0)])
        (let loop ([lpstr_index 1])
          (when (<= lpstr_index lpstr_count)
            (let ([sheet_name (hash-ref xml_hash (format "Properties1.TitlesOfParts1.vt:vector1.vt:lpstr~a" lpstr_index) "")])
              (cond
               [(DATA-SHEET? (list-ref sheet_list (sub1 lpstr_index)))
                (set-DATA-SHEET-sheet_name! (list-ref sheet_list (sub1 lpstr_index)) sheet_name)]
               [(CHART-SHEET? (list-ref sheet_list (sub1 lpstr_index)))
                (set-CHART-SHEET-sheet_name! (list-ref sheet_list (sub1 lpstr_index)) sheet_name)])
              (loop (add1 lpstr_index)))))))))

(define (write-docprops-app [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "docProps"))])
    (make-directory* dir)

    (with-output-to-file (build-path dir "app.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists-to-xml (to-docprops-app)))))))

(define (read-docprops-app [input_dir #f])
  (let ([dir (if input_dir input_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "docProps"))])
    (from-docprops-app (build-path dir "app.xml"))))
