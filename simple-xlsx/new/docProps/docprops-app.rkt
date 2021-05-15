#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")

(provide (contract-out
          [docprops-app (-> list?)]
          [write-docprops-app-file (-> void?)]
          [read-docpros-app (-> void?)]
          ))

(define (docprops-app)
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
        [titles_of_parts_list #f])
    
    (set! heading_pairs_list
          `("HeadingPairs"
            ,(append
              (list
               "vt:vector"
               (cons "size" (number->string (XLSX-sheet_count (*CURRENT_XLSX*))))
               (cons "baseType" "variant"))
            
              (let ([data_sheet_count (length (filter DATA-SHEET? (XLSX-sheet_list (*CURRENT_XLSX*))))])
                  (if (> data_sheet_count 0)
                      (list
                       (list "vt:variant" (list "vt:lpstr" "工作表"))
                       (list "vt:variant" (list "vt:i4" (number->string data_sheet_count))))
                      '()))

              (let ([chart_sheet_count (length (filter CHART-SHEET? (XLSX-sheet_list (*CURRENT_XLSX*))))])
                (if (> chart_sheet_count 0)
                    (list
                     (list "vt:variant" (list "vt:lpstr" "图表"))
                     (list "vt:variant" (list "vt:i4" (number->string chart_sheet_count))))
                    '())))))

    (set! titles_of_parts_list
          `("TitlesOfParts"
            ,(append
              (list "vt:vector"
                    (cons "size" (number->string (XLSX-sheet_count (*CURRENT_XLSX*))))
                    (cons "baseType" "lpstr"))

              (let loop ([sheets (XLSX-sheet_list (*CURRENT_XLSX*))]
                         [index 0]
                         [result_list '()])
                (if (not (null? sheets))
                    (loop (cdr sheets) (add1 index) (cons (list "vt:lpstr" (hash-ref (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) index)) result_list))
                    (reverse result_list))))))
    
    `(,@properties_list ,heading_pairs_list ,titles_of_parts_list)))

(define (write-docprops-app-file)
  (let ([docprops_dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "docProps")])
    (make-directory* docprops_dir)

    (with-output-to-file (build-path docprops_dir "app.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->compact_xml (docprops-app)))))))

(define (read-docpros-app)
  (void))
