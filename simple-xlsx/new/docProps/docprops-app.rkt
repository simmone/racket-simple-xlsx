#lang racket

(require simple-xml)

(require "../../xlsx/xlsx.rkt")
(require "../../sheet/sheet.rkt")

(provide (contract-out
          [docprops-app (-> list?)]
          [write-docprops-app-file (-> void?)]
          [read-docpros-app (-> void?)]
          ))

(define (print-sheet-variant)
  (with-output-to-string
    (lambda ()
      (printf "<HeadingPairs><vt:vector size=\"~a\" baseType=\"variant\">" (length (XLSX-sheet_list (*CURRENT_XLSX*))))

      (let ([data_sheet_count (length (filter (lambda (sheet) (DATA-SHEET? sheet)) (XLSX-sheet_list (*CURRENT_XLSX*))))])
        (when (> data_sheet_count 0)
              (printf "<vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>~a</vt:i4></vt:variant>" data_sheet_count)))

      (let ([chart_sheet_count (length (filter (lambda (sheet) (CHART-SHEET? sheet)) (XLSX-sheet_list (*CURRENT_XLSX*))))])
        (when (> chart_sheet_count 0)
           (printf "<vt:variant><vt:lpstr>图表</vt:lpstr></vt:variant><vt:variant><vt:i4>~a</vt:i4></vt:variant>" chart_sheet_count))))))
 
(define (docprops-app)
  (append
   '("Properties"
     ("xmlns" . "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
     ("xmlns:vt" . "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
     ("Application" . "Microsoft Excel")
     ("DocSecurity" . "0")
     ("ScaleCrop" ."false"))

   (list "HeadingPairs"
         (list "vt:vector"
               (cons "size" (number->string (XLSX-sheet_count (*CURRENT_XLSX*))))
               (cons "baseType" "variant")

               (let ([data_sheet_count (length (filter (lambda (sheet) (DATA-SHEET? sheet)) (XLSX-sheet_list (*CURRENT_XLSX*))))])
                 (if (> data_sheet_count 0)
                     (list
                      (list "vt:variant" (list "vt:lpstr" "工作表"))
                      (list "vt:variant" (list "vt:i4" (number->string data_sheet_count))))
                     '()))

               (let ([chart_sheet_count (length (filter (lambda (sheet) (CHART-SHEET? sheet)) (XLSX-sheet_list (*CURRENT_XLSX*))))])
                 (if (> chart_sheet_count 0)
                     (list
                      (list "vt:variant" (list "vt:lpstr" "图表"))
                      (list "vt:variant" (list "vt:i4" (number->string chart_sheet_count))))
                     '()))))

   (list "TitlesOfParts"
     (list "vt:vector"
           (cons "size" (number->string (XLSX-sheet_count (*CURRENT_XLSX*))))
           (cons "baseType" "lpstr"))

     (let loop ([sheets (XLSX-sheet_list (*CURRENT_XLSX*))]
                [index 0]
                [result_list '()])
       (if (not (null? sheets))
           (loop (cdr sheets) (add1 index) (cons (cons "vt:lpstr" (hash-ref (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) index)) result_list))
           (reverse result_list))))

   '("LinksUpToDate" "false")
   '("SharedDoc" "false")
   '("HyperlinksChanged" "false")
   '("AppVersion" "12.0000")))

(define (write-docprops-app-file)
  (let ([docprops_dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "docProps")])
    (make-directory* docprops_dir)

    (with-output-to-file (build-path docprops_dir "app.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (lists->compact_xml (docprops-app))))))

(define (read-docpros-app)
  (void))
