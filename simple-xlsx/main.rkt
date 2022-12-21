#lang racket

(require "lib/dimension.rkt")
(require "lib/lib.rkt")
(require "style/set-styles.rkt")
(require "style/lib.rkt")
(require "style/border-style.rkt")
(require "style/fill-style.rkt")
(require "style/alignment-style.rkt")
(require "xlsx/xlsx.rkt")
(require "sheet/sheet.rkt")
(require "writer.rkt")
(require "reader.rkt")
(require "reader-writer.rkt")
(require "lib/sheet-lib.rkt")

(provide (contract-out
          [write-xlsx (-> path-string? procedure? any)]
          [read-xlsx (-> path-string? procedure? any)]
          [read-and-write-xlsx (-> path-string? path-string? procedure? any)]
          [add-data-sheet (->* (string? (listof list?)) (cell?) void?)]
          [get-sheet-name-list (-> (listof string?))]
          [with-sheet-ref (-> natural? procedure? any)]
          [with-sheet (-> procedure? any)]
          [with-sheet-name (-> string? procedure? any)]
          [with-sheet-*name* (-> string? procedure? any)]
          [get-sheet-dimension (-> string?)]
          [set-row-range-height (-> string? natural? void?)]
          [set-col-range-width (-> string? natural? void?)]
          [set-freeze-row-col-range (-> natural? natural? void?)]
          [set-merge-cell-range (-> cell-range? void?)]
          [rgb? (-> string? boolean?)]
          [set-cell-range-border-style (-> string? border-direction? rgb? border-mode? void?)]
          [set-cell-range-font-style (-> string? natural? string? rgb? void?)]
          [set-row-range-font-style (-> string? natural? string? rgb? void?)]
          [set-col-range-font-style (-> string? natural? string? rgb? void?)]
          [set-cell-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)]
          [set-row-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)]
          [set-col-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)]
          [set-cell-range-number-style (-> string? string? void?)]
          [set-row-range-number-style (-> string? string? void?)]
          [set-col-range-number-style (-> string? string? void?)]
          [set-cell-range-date-style (-> string? string? void?)]
          [set-row-range-date-style (-> string? string? void?)]
          [set-col-range-date-style (-> string? string? void?)]
          [set-cell-range-fill-style (-> string? rgb? fill-pattern? void?)]
          [set-row-range-fill-style (-> string? rgb? fill-pattern? void?)]
          [set-col-range-fill-style (-> string? rgb? fill-pattern? void?)]
          [add-chart-sheet (-> string?
                               (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D)
                               string?
                               (listof (list/c string? string? string? string? string?)) void?)]

          [cell-value? (-> (or/c string? number? date?) boolean?)]

          [get-rows-count (-> natural?)]
          [get-sheet-ref-rows-count (-> natural? natural?)]
          [get-sheet-name-rows-count (-> string? natural?)]
          [get-sheet-*name*-rows-count (-> string? natural?)]

          [get-cols-count (-> natural?)]
          [get-sheet-ref-cols-count (-> natural? natural?)]
          [get-sheet-name-cols-count (-> string? natural?)]
          [get-sheet-*name*-cols-count (-> string? natural?)]

          [get-row-cells (-> natural? (listof string?))]
          [get-sheet-ref-row-cells (-> natural? natural? (listof string?))]
          [get-sheet-name-row-cells (-> string? natural? (listof string?))]
          [get-sheet-*name*-row-cells (-> string? natural? (listof string?))]

          [get-col-cells (-> (or/c natural? string?) (listof string?))]
          [get-sheet-ref-col-cells (-> natural? (or/c natural? string?) (listof string?))]
          [get-sheet-name-col-cells (-> string? (or/c natural? string?) (listof string?))]
          [get-sheet-*name*-col-cells (-> string? (or/c natural? string?) (listof string?))]

          [get-cell (-> string? cell-value?)]
          [get-sheet-ref-cell (-> natural? string? cell-value?)]
          [get-sheet-name-cell (-> string? string? cell-value?)]
          [get-sheet-*name*-cell (-> string? string? cell-value?)]

          [set-cell! (-> string? cell-value? void?)]
          [set-sheet-ref-cell! (-> natural? string? cell-value? void?)]
          [set-sheet-name-cell! (-> string? string? cell-value? void?)]
          [set-sheet-*name*-cell! (-> string? string? cell-value? void?)]

          [get-row (-> natural? (listof cell-value?))]
          [get-sheet-ref-row (-> natural? natural? (listof cell-value?))]
          [get-sheet-name-row (-> string? natural? (listof cell-value?))]
          [get-sheet-*name*-row (-> string? natural? (listof cell-value?))]

          [set-row! (-> natural? (listof cell-value?) void?)]
          [set-sheet-ref-row! (-> natural? natural? (listof cell-value?) void?)]
          [set-sheet-name-row! (-> string? natural? (listof cell-value?) void?)]
          [set-sheet-*name*-row! (-> string? natural? (listof cell-value?) void?)]

          [get-rows (-> (listof (listof cell-value?)))]
          [get-sheet-ref-rows (-> natural? (listof (listof cell-value?)))]
          [get-sheet-name-rows (-> string? (listof (listof cell-value?)))]
          [get-sheet-*name*-rows (-> string? (listof (listof cell-value?)))]

          [set-rows! (-> (listof (listof cell-value?)) void?)]
          [set-sheet-ref-rows! (-> natural? (listof (listof cell-value?)) void?)]
          [set-sheet-name-rows! (-> string? (listof (listof cell-value?)) void?)]
          [set-sheet-*name*-rows! (-> string? (listof (listof cell-value?)) void?)]

          [get-col (-> (or/c natural? string?) (listof cell-value?))]
          [get-sheet-ref-col (-> natural? (or/c natural? string?) (listof cell-value?))]
          [get-sheet-name-col (-> string? (or/c natural? string?) (listof cell-value?))]
          [get-sheet-*name*-col (-> string? (or/c natural? string?) (listof cell-value?))]

          [set-col! (-> (or/c natural? string?) (listof cell-value?) void?)]
          [set-sheet-ref-col! (-> natural? (or/c natural? string?) (listof cell-value?) void?)]
          [set-sheet-name-col! (-> string? (or/c natural? string?) (listof cell-value?) void?)]
          [set-sheet-*name*-col! (-> string? (or/c natural? string?) (listof cell-value?) void?)]

          [get-cols (-> (listof (listof cell-value?)))]
          [get-sheet-ref-cols (-> natural? (listof (listof cell-value?)))]
          [get-sheet-name-cols (-> string? (listof (listof cell-value?)))]
          [get-sheet-*name*-cols (-> string? (listof (listof cell-value?)))]

          [set-cols! (-> (listof (listof cell-value?)) void?)]
          [set-sheet-ref-cols! (-> natural? (listof (listof cell-value?)) void?)]
          [set-sheet-name-cols! (-> string? (listof (listof cell-value?)) void?)]
          [set-sheet-*name*-cols! (-> string? (listof (listof cell-value?)) void?)]

          [get-range-values (-> string? (listof cell-value?))]
          [get-sheet-ref-range-values (-> natural? string? (listof cell-value?))]
          [get-sheet-name-range-values (-> string? string? (listof cell-value?))]
          [get-sheet-*name*-range-values (-> string? string? (listof cell-value?))]

          [oa_date_number->date (->* (number?) (boolean?) date?)]
          ))
