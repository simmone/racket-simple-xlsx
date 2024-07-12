#lang racket

(require rackunit/text-ui
         rackunit
         "../../../main.rkt"
         "../../../style/styles.rkt"
         racket/runtime-path
         "../cell-fill-style/cell-fill-style-verify.rkt"
         "../cell-alignment-style/cell-alignment-style-verify.rkt"
         "../cell-border-style/cell-border-style-verify.rkt"
         "../cell-date-style/cell-date-style-verify.rkt"
         "../cell-font-style/cell-font-style-verify.rkt"
         "../cell-number-style/cell-number-style-verify.rkt"
         "../col-style/col-style-verify.rkt"
         "../freeze/freeze-verify.rkt"
         "../merge-cells/merge-cells-verify.rkt"
         "../row-col-overlap/row-col-overlap-verify.rkt"
         "../row-style/row-style-verify.rkt"
         "../width-height/width-height-verify.rkt")

(define-runtime-path styles_file "_styles.xlsx")
(define-runtime-path styles_read_write_file "_styles_read_write.xlsx")
(define-runtime-path gsheet_file "gsheet.xlsx")
(define-runtime-path gsheet_read_write_file "_gsheet_read_write.xlsx")
(define-runtime-path msoffice_file "msoffice.xlsx")
(define-runtime-path msoffice_read_write_file "_msoffice_read_write.xlsx")
(define-runtime-path wps_file "wps.xlsx")
(define-runtime-path wps_read_write_file "_wps_read_write.xlsx")
(define-runtime-path libre_file "libre.xlsx")
(define-runtime-path libre_read_write_file "_libre_read_write.xlsx")


(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles"

    (dynamic-wind
        (lambda () (void))
        (lambda ()
          (printf "common style write\n")
          (write-xlsx
           styles_file
           (lambda ()
             (add-cell-fills-sheet)
             (add-cell-alignment-sheet)
             (add-cell-border-sheet)
             (add-cell-date-sheet)
             (add-cell-font-sheet)
             (add-cell-number-sheet)
             (add-col-style-sheet)
             (add-freeze-style-sheet)
             (add-merge-cells-style-sheet)
             (add-row-col-style-sheet)
             (add-col-row-style-sheet)
             (add-row-style-sheet)
             (add-width-height-style-sheet)
             ))

          (printf "common style read and write\n")
          (read-and-write-xlsx
           styles_file
           styles_read_write_file
           (lambda ()
             (set-more-cell-fills)
             (set-more-cell-alignments)
             (set-more-cell-borders)
             (set-more-cell-dates)
             (set-more-cell-fonts)
             (set-more-cell-numbers)
             (set-more-col-styles)
             (set-more-freeze-styles)
             (set-more-merge-cells-styles)
             (set-more-row-col-styles)
             (set-more-col-row-styles)
             (set-more-row-styles)
             (set-more-width-height-styles)
             ))


          (printf "google sheet style read and write\n")
          (read-and-write-xlsx
           gsheet_file
           gsheet_read_write_file
           (lambda ()
             (set-more-cell-fills)
             (set-more-cell-alignments)
             (set-more-cell-borders)
             (set-more-cell-dates)
             (set-more-cell-fonts)
             (set-more-cell-numbers)
             (set-more-col-styles)
             (set-more-merge-cells-styles)
             (set-more-row-col-styles)
             (set-more-col-row-styles)
             (set-more-row-styles)
             (set-more-width-height-styles)
             ))

          (printf "wps style read and write\n")
          (read-and-write-xlsx
           wps_file
           wps_read_write_file
           (lambda ()
             (set-more-cell-fills)
             (set-more-cell-alignments)
             (set-more-cell-borders)
             (set-more-cell-dates)
             (set-more-cell-fonts)
             (set-more-cell-numbers)
             (set-more-col-styles)
             (set-more-freeze-styles)
             (set-more-merge-cells-styles)
             (set-more-row-col-styles)
             (set-more-col-row-styles)
             (set-more-row-styles)
             (set-more-width-height-styles)
             ))

          (printf "libre office style read and write\n")
          (read-and-write-xlsx
           libre_file
           libre_read_write_file
           (lambda ()
             (set-more-cell-fills)
             (set-more-cell-alignments)
             (set-more-cell-borders)
             (set-more-cell-dates)
             (set-more-cell-fonts)
             (set-more-cell-numbers)
             (set-more-col-styles)
             (set-more-freeze-styles)
             (set-more-merge-cells-styles)
             (set-more-row-col-styles)
             (set-more-col-row-styles)
             (set-more-row-styles)
             (set-more-width-height-styles)
             ))

          (printf "microsoft excel style read and write\n")
          (read-and-write-xlsx
           msoffice_file
           msoffice_read_write_file
           (lambda ()
             (set-more-cell-fills)
             (set-more-cell-alignments)
             (set-more-cell-borders)
             (set-more-cell-dates)
             (set-more-cell-fonts)
             (set-more-cell-numbers)
             (set-more-col-styles)
             (set-more-freeze-styles)
             (set-more-merge-cells-styles)
             (set-more-row-col-styles)
             (set-more-col-row-styles)
             (set-more-row-styles)
             (set-more-width-height-styles)
             ))

          )
        (lambda ()
          ;(void)
          (delete-file styles_file)
          (delete-file styles_read_write_file)
          (delete-file wps_read_write_file)
          (delete-file libre_read_write_file)
          (delete-file gsheet_read_write_file)
          (delete-file msoffice_read_write_file)
          )))
   ))

(run-tests test-styles)
