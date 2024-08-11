(ql:quickload :uiop)
(ql:quickload :asdf)
(ql:quickload :cl-simple-xlsx)
(ql:quickload :cl-fast-xml)
(ql:quickload :cl-fad)
(ql:quickload :local-time)
(ql:quickload :fiveam)

(defpackage cl-simple-xlsx/tests/main
  (:use :cl
        :cl-simple-xlsx
	:cl-fast-xml
        :cl-fad
	:local-time
        :uiop
        :fiveam)
  (:shadowing-import-from :cl-simple-xlsx))

(in-package :cl-simple-xlsx/tests/main)

(defun ht2al (ht)
  (let ((alist '()))
    (maphash (lambda (k v) (push (cons k v) alist)) ht)
    alist))

;; I want a function which can list all xml files in a directory recursively

;; I want a function which lists all file/direcotry objects in a directory recursively or for a given depth

;; (list-directories)
;; (list-directories #p"path")
;; (list-directories :pattern "regex") ;; :regex-pattern   :recursively t :level

;; (list-files)
;; (list-files #p"path")

;; probe-file returns name with "/" ending when directory and without when file

;; actually, I want `ls` and `glob`.
;; `tree`

;; loop over every file in a directory and apply a function on it.

;; giving functions as arguments is very powerful, because one can pack an entire program into a function.
;; one can chain functions.

;; (defun path-as-string (path)
;;   (format nil "~a" path))     => namestring

(defun path-listform-p (path)
  (and (listp path)
       (eq :absolute (car path))))

(defun path-to-string (path)
  (if (path-listform-p path)
      (namestring (make-pathname :directory path))
      (format nil "~a" (probe-file path))))

(defun directoryp (path)
  (char= #\/ (car (last (coerce (namestring (probe-file path)) 'list)))))

(defun filep (path)
  (char/= #\/ (car (last (coerce (namestring (probe-file path)) 'list)))))

(defun ls (&optional (dir-path "."))
  (directory (make-pathname :name :wild :type :wild :directory (path-to-string dir-path))))

(defun list-directories (&optional (dir-path "."))
  (remove-if-not #'directoryp (ls dir-path)))

(defun list-files (&optional (dir-path "."))
  (remove-if #'directoryp (ls dir-path)))

(defun pattern-matches-p (str &key (pattern ".*"))
  (cl-ppcre:scan pattern str))

(defun ls-pattern (&key (dir-path ".") (pattern ".*"))
  (remove-if-not (lambda (p) (pattern-matches-p (path-to-string p) :pattern pattern)) (ls dir-path)))

(defun list-directories-pattern (&key (dir-path ".") (pattern ".*"))
  (remove-if-not (lambda (p) (and (directoryp p)
				  (pattern-matches-p (path-to-string p) :pattern pattern)))
		 (ls dir-path)))

(defun list-files-pattern (&key (dir-path ".") (pattern ".*"))
  (remove-if-not (lambda (p) (and (filep p)
				  (pattern-matches-p (path-to-string p) :pattern pattern)))
		 (ls dir-path)))

(defun list-files-recursively-pattern (&key (dir-path ".") (pattern ".*"))
    (let ((dirs (list-directories dir-path))
	  (files (list-files-pattern :dir-path dir-path :pattern pattern)))
      (cond ((null dirs) files)
	    (t (append files
		       (mapcan (lambda (dir) (list-files-recursively-pattern :dir-path dir :pattern pattern))
		       dirs))))))


(defun parent (path)
  (path-to-string (pathname-directory path)))

(defun filename (path)
  (file-namestring path))

(defun split-path (path)
  "Return values directory and filename component of a path."
  (let ((dir (parent path))
	(filename (filename path)))
    (values dir filename)))

(defun paths-to-hash-table (paths)
  "Returns a hash-table with file name as key and a list of absolute paths as value"
  (let ((ht (make-hash-table :test 'equal)))
    (loop for path in paths
	  do (multiple-value-bind (dir filename) (split-path path)
	       (declare (ignore dir))
	       (setf (gethash filename ht) (cons path (gethash filename ht nil))))
	  finally (return ht))))

(defun assess-xml-rel-files (&optional (dir-path "."))
  (list-files-recursively-pattern :dir-path dir-path :pattern ".*\\.(xml|rels)$"))

(defun prepare-xml-rel-hash-table (&optional (dir-path "."))
  (paths-to-hash-table (assess-xml-rel-files dir-path)))

(defparameter *paths-hash-table* (prepare-xml-rel-hash-table "."))

;; this is really a huge help for dealing witht he xml and rels paths!

;; I should create a package for path names with seibel's and my functions!

;; linux functions mv rm rmdir cp mkdir -p
;; and eventually the file meta info modifier functions are needed

(def-suite lib
  :description "Test lib functions")

(in-suite lib)
(test test-dimension-lib
  (is (cell-range-p "A1-B2"))
  (is (cell-range-p "A1:B2"))
  (is (cell-range-p "A1"))
  (is (not (cell-range-p "ksdkf")))

  (is (cellp "A1"))
  (is (not (cellp "a")))
  (is (not (cellp "23")))

  (is (string= (row-col->cell 1 1) "A1"))
  (is (string= (row-col->cell 2 1) "A2"))
  (is (string= (row-col->cell 2 5) "E2"))

  (let ((row-col (cell->row-col "A1")))
    (is (= (car row-col) 1))
    (is (= (cdr row-col) 1)))

  (let ((row-col (cell->row-col "A2")))
    (is (= (car row-col) 2))
    (is (= (cdr row-col) 1)))

  (let ((row-col (cell->row-col "E2")))
    (is (= (car row-col) 2))
    (is (= (cdr row-col) 5)))

  (let ((row-col (cell->row-col "C10")))
    (is (= (car row-col) 10))
    (is (= (cdr row-col) 3)))

  (let ((row-col (cell->row-col "AB23")))
    (is (= (car row-col) 23))
    (is (= (cdr row-col) 28)))

  (is (equal (range->row-col-pair "A1-B2") '((1 . 1) . (2 . 2))))
  (is (equal (range->row-col-pair "A2-D7") '((2 . 1) . (7 . 4))))
  (is (equal (range->row-col-pair "B2-A1") '((1 . 1) . (1 . 1))))
  (is (equal (range->row-col-pair "A1") '((1 . 1) . (1 . 1))))
  (is (equal (range->row-col-pair "A2") '((2 . 1) . (2 . 1))))
  (is (equal (range->row-col-pair "A2-A1") '((1 . 1) . (1 . 1))))
  (is (equal (range->row-col-pair "B1-A1") '((1 . 1) . (1 . 1))))
  
  (is (equal (range->capacity "A1:F4") '(4 . 6)))
  (is (equal (range->capacity "B2:F4") '(3 . 5)))

  (is (equal (capacity->range '(4 . 6)) "A1:F4"))
  (is (equal (capacity->range '(4 . 6) "B1") "B1:G4"))
  (is (equal (capacity->range '(4 . 6) "C2") "C2:H5"))

  (is (= (col-abc->number "A") 1))
  (is (= (col-abc->number "B") 2))
  (is (= (col-abc->number "Z") 26))
  (is (= (col-abc->number "AA") 27))
  (is (= (col-abc->number "AB") 28))
  (is (= (col-abc->number "AZ") 52))
  (is (= (col-abc->number "BA") 53))
  (is (= (col-abc->number "YZ") 676))
  (is (= (col-abc->number "ZA") 677))
  (is (= (col-abc->number "ZZ") 702))
  (is (= (col-abc->number "AAA") 703))

  (is (string= (col-number->abc 1) "A"))
  (is (string= (col-number->abc 26) "Z"))
  (is (string= (col-number->abc 27) "AA"))
  (is (string= (col-number->abc 28) "AB"))
  (is (string= (col-number->abc 29) "AC"))
  (is (string= (col-number->abc 51) "AY"))
  (is (string= (col-number->abc 52) "AZ"))
  (is (string= (col-number->abc 53) "BA"))
  (is (string= (col-number->abc 676) "YZ"))
  (is (string= (col-number->abc 677) "ZA"))
  (is (string= (col-number->abc 702) "ZZ"))
  (is (string= (col-number->abc 703) "AAA")))

(test to-col-range
  (is (equal (to-col-range "A") '(1 . 1)))
  (is (equal (to-col-range "B") '(2 . 2)))
  (is (equal (to-col-range "2") '(2 . 2)))
  (is (equal (to-col-range "C-D") '(3 . 4)))
  (is (equal (to-col-range "3-4") '(3 . 4)))
  (is (equal (to-col-range "1-26") '(1 . 26)))
  (is (equal (to-col-range "26-1") '(1 . 1)))
  (is (equal (to-col-range "A-Z") '(1 . 26)))
  (is (equal (to-col-range "Z-A") '(1 . 1)))
  (is (equal (to-col-range "A-B-C") '(1 . 1)))
  (is (equal (to-col-range "A-ksdk344") '(1 . 1)))
  (is (equal (to-col-range "sdksjdkf-%^%$#") '(1 . 1)))
  (is (equal (to-col-range "A-Z") '(1 . 26)))
  (is (equal (to-col-range "A-z") '(1 . 26)))
  (is (equal (to-col-range "A-B") '(1 . 2)))
  (is (equal (to-col-range "A") '(1 . 1)))
  (is (equal (to-col-range "7-12") '(7 . 12)))
  (is (equal (to-col-range "10") '(10 . 10))))

(test to-row-range
  (is (equal (to-row-range "1-4") '(1 . 4)))
  (is (equal (to-row-range "7-12") '(7 . 12)))
  (is (equal (to-row-range "10") '(10 . 10)))
  (is (equal (to-row-range "2-1") '(1 . 1)))
  (is (equal (to-row-range "A-B") '(1 . 1)))
  (is (equal (to-row-range "12-7") '(1 . 1))))


(test test-range->range-xml
  (is (string= (range->range-xml "C2-C10") "$C$2:$C$10"))
  (is (string= (range->range-xml "C2-Z2") "$C$2:$Z$2"))
  (is (string= (range->range-xml "AB20-AB100") "$AB$20:$AB$100")))

(test test-range-xml->range
  (is (string= (range-xml->range "$C$2:$C$10") "C2-C10"))
  (is (string= (range-xml->range "$C$2:$Z$2") "C2-Z2"))
  (is (string= (range-xml->range "$AB$20:$AB$100") "AB20-AB100")))


(test test-cell-range->cell-list
  (is (equal (cell-range->cell-list "A1-D1") '("A1" "B1" "C1" "D1")))
  (is (equal (cell-range->cell-list "B1-D2") '("B1" "C1" "D1" "B2" "C2" "D2")))
  (is (equal (cell-range->cell-list "A2-A1") '("A1")))
  (is (equal (cell-range->cell-list "B1-A1") '("A1"))))

(test test-get-cell-range-side-cells
  (multiple-value-bind (left-cells right-cells top-cells bottom-cells) (get-cell-range-four-sides-cells "A1")
    (is (equal top-cells '("A1")))
    (is (equal bottom-cells '("A1")))
    (is (equal left-cells '("A1")))
    (is (equal right-cells '("A1"))))

  (multiple-value-bind (left-cells right-cells top-cells bottom-cells) (get-cell-range-four-sides-cells "A1-B1")
    (is (equal top-cells '("A1" "B1")))
    (is (equal bottom-cells '("A1" "B1")))
    (is (equal left-cells '("A1")))
    (is (equal right-cells '("B1"))))

  (multiple-value-bind (left-cells right-cells top-cells bottom-cells) (get-cell-range-four-sides-cells "A1-A2")
    (is (equal top-cells '("A1")))
    (is (equal bottom-cells '("A2")))
    (is (equal left-cells '("A1" "A2")))
    (is (equal right-cells '("A1" "A2"))))

  (multiple-value-bind (left-cells right-cells top-cells bottom-cells) (get-cell-range-four-sides-cells "A1-A2")
    (is (equal top-cells '("A1")))
    (is (equal bottom-cells '("A2")))
    (is (equal left-cells '("A1" "A2")))
    (is (equal right-cells '("A1" "A2"))))

  (multiple-value-bind (left-cells right-cells top-cells bottom-cells) (get-cell-range-four-sides-cells "A1-C3")
    (is (equal top-cells '("A1" "B1" "C1")))
    (is (equal bottom-cells '("A3" "B3" "C3")))
    (is (equal left-cells '("A1" "A2" "A3")))
    (is (equal right-cells '("C1" "C2" "C3")))))

(defparameter *content-type-file* (gethash "[Content_Types].xml" *paths-hash-table*))


(test maintain-sheet-data-consistency
  (signals simple-error (maintain-sheet-data-consistency '() ""))
  (signals type-error (maintain-sheet-data-consistency '((1) 4) "")) ;; 4 should be a list

  (is (equal (maintain-sheet-data-consistency '((1) (1 2)) "") '((1 "") (1 2))))
  (is (equal (maintain-sheet-data-consistency '((1) (1 2)) 0) '((1 0) (1 2))))
  (is (equal (maintain-sheet-data-consistency '((1 2) (3 4)) 0) '((1 2) (3 4)))))


(test test-check-lines1
  (let ((expected-port (make-string-input-stream "abc"))
        (test-port (make-string-input-stream "abc")))
    (is (check-lines-p expected-port test-port)))

  (let ((expected-port (make-string-input-stream (format nil "abc~%11")))
        (test-port (make-string-input-stream (format nil "abc~%11"))))
    (is (check-lines-p expected-port test-port)))

  (let ((expected-port (make-string-input-stream (format nil " a~%~% b~%~%c~%")))
        (test-port (make-string-input-stream (format nil " a~%~% b~%~%c~%"))))
    (is (check-lines-p expected-port test-port))))

(let ((port1 (make-string-input-stream (format nil "abc~%~%11")))
      (port2 (make-string-input-stream (format nil "abc~%11"))))
  (let* ((lines1 (cl-simple-xlsx::port->lines port1))
	 (lines2 (cl-simple-xlsx::port->lines port2))
	 (n1 (length lines1))
	 (n2 (length lines2))
	 (delta (- n1 n2)))
    (list lines1 lines2 n1 n2 delta)))

;; (let ((port1 (make-string-input-stream (format nil "abc~%~%~%")))
;;       (port2 (make-string-input-stream (format nil "abc~%~%"))))
;;   (check-lines-p port1 port2))


;; Define paths
(defparameter *zip-xlsx-temp-directory* (asdf:system-relative-pathname "cl-simple-xlsx" "test/test-directory/"))
(defparameter *content-type-file* (make-pathname :directory (pathname-directory *zip-xlsx-temp-directory*)
						 :name "[Content_Types]"
						 :type "xml"))
(defparameter *rels-directory* (merge-pathnames "_rels/" *zip-xlsx-temp-directory*))
(defparameter *doc-props-directory* (merge-pathnames"docProps/"  *zip-xlsx-temp-directory*))
(defparameter *xl-directory* (merge-pathnames "xl/"  *zip-xlsx-temp-directory*))
(defparameter *zip-xlsx-file* "test.xlsx")

(test test-format-w3cdtf
  (let* ((timestamp (local-time:encode-timestamp 996159076 44 17 13 2 1 2015))
	 (offset-string (local-time:format-timestring nil timestamp :format '(:gmt-offset))))
    (is (string= (format-w3cdtf (local-time:encode-timestamp 996159076 44 17 13 2 1 2015)) 
		 (format nil "~a~a" "2015-01-02T13:17:44" offset-string)))))

(test test-date->oa-date-number
  (let ((date1 (local-time:encode-timestamp 0 0 0 0 17 9 2018))
        (date2 (local-time:encode-timestamp 0 0 0 0 16 9 2018)))
    (is (= (date->oa-date-number date1) 43360))
    (is (= (date->oa-date-number date2) 43359))))

(defun timestamp-equal (x y)
  (string= (format nil "~a" x) (format nil "~a" y)))

(test test-oa-date-number->date
  (let ((expected-date1 (local-time:encode-timestamp 0 0 0 0 18 9 2018 :timezone local-time:+gmt-zone+))
        (expected-date2 (local-time:encode-timestamp 0 0 0 0 17 9 2018 :timezone local-time:+gmt-zone+)))
    (is (timestamp-equal  (oa-date-number->date 43360 :local-time-p nil) expected-date1))
    (is (timestamp-equal (oa-date-number->date 43359 :local-time-p nil) expected-date2))
    (is (timestamp-equal  (oa-date-number->date 43359.1212121 :local-time-p nil ) expected-date2))))


(test test-zip-and-unzip
  (unwind-protect
      (progn
        ;; Setup phase
        (ensure-directories-exist *zip-xlsx-temp-directory*)
        (with-open-file (stream *content-type-file*
                                :direction :output
                                :if-does-not-exist :create
                                :if-exists :supersede)
          (format stream ""))
        (ensure-directories-exist *rels-directory*)
        (ensure-directories-exist *xl-directory*)
        (ensure-directories-exist *doc-props-directory*)
        (zip-xlsx *zip-xlsx-file* *zip-xlsx-temp-directory*)
	;; (cl-fad:delete-directory-and-files *zip-xlsx-temp-directory*)
        
        ;; Test phase
        (unzip-xlsx *zip-xlsx-file* (format nil "~a" *zip-xlsx-temp-directory*))
        (is (probe-file *content-type-file*))
        (is (probe-file *rels-directory*))
        (is (probe-file *doc-props-directory*))
        (is (probe-file *xl-directory*)))
    
    ;; Cleanup phase
    (when (cl-fad:file-exists-p *zip-xlsx-file*)
        (delete-file *zip-xlsx-file*))
    (cl-fad:delete-directory-and-files *zip-xlsx-temp-directory*)))


(run! 'lib)
