#lang at-exp racket/base

(require racket/port)
(require racket/class)
(require racket/file)
(require racket/list)
(require racket/contract)

(require "../../xlsx/xlsx.rkt")
(require "../../xlsx/sheet.rkt")

(provide (contract-out
          [write-docprops-app (-> string?)]
          [write-docprops-app-file (-> void?)]
          [read-docpros-app (-> void?)]
          ))

(define S string-append)

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
 
(define (write-docprops-app) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>@|(print-sheet-variant sheet_list)|</vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="@|(number->string (length (XLSX-sheet_list (*CURRENT_XLSX*))))|" baseType="lpstr">@|(with-output-to-string
  (lambda ()
    (let loop ([sheets (XLSX-sheet_list (*CURRENT_XLSX*))]
               [index 0])
      (when (not (null? sheets))
        (printf "<vt:lpstr>~a</vt:lpstr>" (hash-ref (XLSX-sheet_index_name_map (*CURRENT_XLSX*)) index))
        (loop (cdr sheets) (add1 index))))))|</vt:vector></TitlesOfParts><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000</AppVersion></Properties>
})

(define (write-docprops-app-file)
  (let ([docprops_dir (build-path (XLSX-xlsx_dir (*CURRENT_XLSX*)) "docProps")])
    (make-directory* docprops_dir)

    (with-output-to-file (build-path docprops_dir "app.xml")
      #:exists 'replace
      (lambda ()
        (printf "~a" (write-docprops-app (get-field sheets xlsx)))))))

(define (read-docpros-app)
  (void))
