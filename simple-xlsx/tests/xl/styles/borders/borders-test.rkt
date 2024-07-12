#lang racket

(require fast-xml
         rackunit/text-ui rackunit
         "../../../../xlsx/xlsx.rkt"
         "../../../../sheet/sheet.rkt"
         "../../../../style/style.rkt"
         "../../../../style/styles.rkt"
         "../../../../style/border-style.rkt"
         "../../../../style/assemble-styles.rkt"
         "../../../../style/set-styles.rkt"
         "../../../../lib/lib.rkt"
         "../../../../xl/styles/borders.rkt"
         racket/runtime-path)

(define-runtime-path borders_file "borders.xml")
(define-runtime-path side_borders_file "side_borders.xml")

(provide (contract-out
          [set-border-styles (-> void?)]
          [check-border-styles (-> void?)]
          [borders_file path-string?]
          ))

(define (set-border-styles)
  (set-cell-range-border-style "A1" "all" "FF0000" "dashed")
  (set-cell-range-border-style "B1" "all" "000000" "dashed")
  (set-cell-range-border-style "C1" "bottom" "0000FF" "thick")
  (set-cell-range-border-style "D1" "top" "000000" "thin"))

(define (check-border-styles)
  (check-equal? (length (STYLES-border_list (*STYLES*))) 5)

  (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 0)
                (BORDER-STYLE #f #f #f #f #f #f #f #f))
  (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 1)
                (BORDER-STYLE "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed" "FF0000" "dashed"))
  (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 2)
                (BORDER-STYLE "000000" "dashed" "000000" "dashed" "000000" "dashed" "000000" "dashed"))
  (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 3)
                (BORDER-STYLE #f #f #f #f #f #f "0000FF" "thick"))
  (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 4)
                (BORDER-STYLE #f #f #f #f "000000" "thin" #f #f)))

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-to-borders"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (with-sheet (lambda () (set-border-styles)))

       (strip-styles)
       (assemble-styles)

       (call-with-input-file borders_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content
             (to-borders (STYLES-border_list (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual))))))))

   (test-case
    "test-from-borders"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(("month1" "month2" "month3" "month4" "real") (201601 100 110 1110 6.9)))
       (add-data-sheet "Sheet2" '((1)))
       (add-data-sheet "Sheet3" '((1)))

       (from-borders
        (xml-port-to-hash
         (open-input-string (format "<styleSheet>~a</styleSheet>" (file->string borders_file)))
         '(
           "styleSheet.borders.border"
           "styleSheet.borders.border.left.color.rgb"
           "styleSheet.borders.border.right.color.rgb"
           "styleSheet.borders.border.top.color.rgb"
           "styleSheet.borders.border.bottom.color.rgb"
           "styleSheet.borders.border.left.style"
           "styleSheet.borders.border.right.style"
           "styleSheet.borders.border.top.style"
           "styleSheet.borders.border.bottom.style"
           )
         ))

       (check-border-styles)
       )))

   (test-case
    "test-side-borders"

    (with-xlsx
     (lambda ()
       (add-data-sheet "Sheet1"
                       '(
                         ("month1" "month2" "month3")
                         ("month1" "month2" "month3")
                         ("month1" "month2" "month3")))

       (with-sheet
        (lambda ()
          (set-cell-range-border-style "A1-C3" "side" "000000" "thick")))

       (strip-styles)
       (assemble-styles)

       (check-equal? (length (STYLES-border_list (*STYLES*))) 9)
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 0)
                     (BORDER-STYLE #f #f #f #f #f #f #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 1)
                     (BORDER-STYLE "000000" "thick" #f #f #f #f #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 2)
                     (BORDER-STYLE #f #f "000000" "thick" #f #f #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 3)
                     (BORDER-STYLE "000000" "thick" #f #f "000000" "thick" #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 4)
                     (BORDER-STYLE  #f #f #f #f "000000" "thick" #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 5)
                     (BORDER-STYLE #f #f "000000" "thick" "000000" "thick" #f #f))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 6)
                     (BORDER-STYLE "000000" "thick" #f #f #f #f "000000" "thick"))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 7)
                     (BORDER-STYLE  #f #f #f #f #f #f "000000" "thick"))
       (check-equal? (list-ref (STYLES-border_list (*STYLES*)) 8)
                     (BORDER-STYLE #f #f "000000" "thick" #f #f "000000" "thick"))

       (call-with-input-file side_borders_file
         (lambda (expected)
           (call-with-input-string
            (lists-to-xml_content
             (to-borders (STYLES-border_list (*STYLES*))))
            (lambda (actual)
              (check-lines? expected actual))))))))
   ))

(run-tests test-styles)
