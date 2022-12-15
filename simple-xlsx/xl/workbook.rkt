#lang racket

(require simple-xml)

(require "../xlsx/xlsx.rkt")
(require "../sheet/sheet.rkt")

(provide (contract-out
          [workbook (-> list?)]
          [write-workbook (->* () (path-string?) void?)]
          [read-workbook (-> void?)]
          ))

(define (workbook)
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

(define (write-workbook [output_dir #f])
  (let ([dir (if output_dir output_dir (build-path (XLSX-xlsx_dir (*XLSX*)) "xl"))])
    (make-directory* dir)

    (with-output-to-file (build-path dir "workbook.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (lists->xml (workbook)))))))

(define (read-workbook)
  (void))

