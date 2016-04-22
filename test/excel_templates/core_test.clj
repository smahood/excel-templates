(ns excel-templates.core-test
  (:use clojure.test)
  (:require [excel-templates.build]))



(deftest temp-xlsx-file-operations
    ; create temp xlsx file
  ; write changes to temp xlsx file
  ; remove temp-xlsx-file
    (let [tempfile (excel-templates.build/create-temp-xlsx-file "excel-templates-test")]

    (println tempfile)

  ))


(deftest can-write-changes-to-temp-xlsx-file)

(deftest can-remove-temp-xlsx-file)

