#lang racket

(require rackunit/text-ui rackunit)

(require"../../../style/style.rkt")
(require"../../../style/set-styles.rkt")
(require"../../../style/font-style.rkt")
(require"../../../style/fill-style.rkt")
(require"../../../style/alignment-style.rkt")
(require"../../../style/border-style.rkt")
(require"../../../style/number-style.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-update-one-kind-style"

    (let* ([style (new-style)])
      (check-equal? (STYLE-hash_code style) "<s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "<p><p><p><p><p><p><p>")))
      (check-equal? (STYLE-hash_code style) "<p><p><p><p><p><p><p><s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "f00000<p>thin<p><p><p><p><p><p>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p><p><p><p><p><p><s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "<p><p>0f0000<p>thick<p><p><p><p>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p><p><p><p><s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "<p><p><p><p>00f000<p>double<p><p>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p><p><s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "<p><p><p><p><p><p>000f00<p>dashed")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "ffffff<p>thin<p><p><p><p><p><p>")))
      (check-equal? (STYLE-hash_code style) "FFFFFF<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>")

      (set! style (update-style style (border-style-from-hash-code "<p><p>0f0000<p>thick<p><p><p><p>")))
      (check-equal? (STYLE-hash_code style) "FFFFFF<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>")

      (set! style (update-style style (font-style-from-hash-code "10<p>Arial<p>0000ff")))
      (check-equal?
       (STYLE-hash_code style)
       "FFFFFF<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s><s><s>")

      (set! style (update-style style (alignment-style-from-hash-code "left<p>top")))
      (check-equal?
       (STYLE-hash_code style)
       "FFFFFF<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s><s>")

      (set! style (update-style style (number-style-from-hash-code "#,###.00")))
      (check-equal?
       (STYLE-hash_code style)
       "FFFFFF<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s>#,###.00<s>")

      (set! style (update-style style (fill-style-from-hash-code "0000FF<p>solid")))
      (check-equal?
       (STYLE-hash_code style)
       "FFFFFF<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s>#,###.00<s>0000FF<p>solid")
      ))

   (test-case
    "test-update-whole-style"

    (let* ([style (new-style)])
      (check-equal? (STYLE-hash_code style) "<s><s><s><s>")

      (set! style (update-style style (style-from-hash-code "<p><p><p><p><p><p><p><s><s><s><s>")))
      (check-equal? (STYLE-hash_code style) "<p><p><p><p><p><p><p><s><s><s><s>")

      (set! style (update-style style (style-from-hash-code "f00000<p>thin<p><p><p><p><p><p><s><s><s><s>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p><p><p><p><p><p><s><s><s><s>")

      (set! style (update-style style (style-from-hash-code "<p><p>0f0000<p>thick<p><p><p><p><s><s><s><s>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p><p><p><p><s><s><s><s>")

      (set! style (update-style style (style-from-hash-code "<p><p><p><p>00F000<p>double<p><p><s><s><s><s>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p><p><s><s><s><s>")

      (set! style (update-style style (style-from-hash-code "<p><p><p><p><p><p>000F00<p>dashed<s><s><s><s>")))
      (check-equal? (STYLE-hash_code style) "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s><s><s><s>")

      (set! style (update-style style (style-from-hash-code "<s>10<p>Arial<p>0000ff<s><s><s>")))
      (check-equal?
       (STYLE-hash_code style)
       "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s><s><s>")

      (set! style (update-style style (style-from-hash-code "<s><s>left<p>top<s><s>")))
      (check-equal?
       (STYLE-hash_code style)
       "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s><s>")

      (set! style (update-style style (style-from-hash-code "<s><s><s>#,###.00<s>")))
      (check-equal?
       (STYLE-hash_code style)
       "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s>#,###.00<s>")

      (set! style (update-style style (style-from-hash-code "<s><s><s><s>0000FF<p>solid")))
      (check-equal?
       (STYLE-hash_code style)
       "F00000<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s>#,###.00<s>0000FF<p>solid")

      (set!
       style
       (update-style
        style
        (style-from-hash-code "fffff0<p>thin<p><p><p><p><p><p><s><s><s><s>000FFF<p>solid")))
      (check-equal?
       (STYLE-hash_code style)
       "FFFFF0<p>thin<p>0F0000<p>thick<p>00F000<p>double<p>000F00<p>dashed<s>10<p>Arial<p>0000FF<s>left<p>top<s>#,###.00<s>000FFF<p>solid")
    ))
   ))

(run-tests test-styles)
