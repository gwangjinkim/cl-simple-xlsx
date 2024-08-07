(defpackage cl-simple-xlsx/tests/main
  (:use :cl
        :cl-simple-xlsx
        :rove))
(in-package :cl-simple-xlsx/tests/main)

;; NOTE: To run this test file, execute `(asdf:test-system :cl-simple-xlsx)' in your Lisp.

(deftest test-target-1
  (testing "should (= 1 1) to be true"
    (ok (= 1 1))))
