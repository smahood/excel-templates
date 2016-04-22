(ns excel-unzip.core
  (:import
    [net.lingala.zip4j.core ZipFile]
    [java.io File FileInputStream]
    ;[java.util.zip ZipEntry ZipFile ZipOutputStream ZipInputStream]
    )
  (:require [clojure.zip :as zip]
            [clojure.java.io :as io]))







(comment
  (defn entries [zip-stream]
    (take-while #(not (nil? %))
                (repeatedly #(.getNextEntry zip-stream))))

  (defn walk-zip [input-file]
    (with-open [z (ZipFile. input-file)]
      z
      (doseq [e (entries z)]
        (println "Name")
        (println (.getName e))
        (println (.getSize e))
        (.read e)

        (.closeEntry z))
      ))


  (walk-zip "resources/blank.xlsx")



  (defn zipfile-to-data [input-file]
    (let [buf-size 65536
          buf (byte-array buf-size)]
      (with-open [zip-file (ZipFile. input-file)]
        zip-file)))
  )


(comment (zipfile-to-data "resources/blank.xlsx"))



(defn unzip [src dest]
  (try
    (let [zipfile (ZipFile. src)]
      (.extractAll zipfile dest))
    (catch Exception e
      (println (.getMessage e)))))

(comment (unzip "foo2.xlsx" "foo_unzipped"))


(defn dir-tree-to-html [src]
  (try
    (let [dir (clojure.java.io/file src)
          dir-tree (file-seq dir)]
      (clojure.string/join (map #(str
                                  "<div>"
                                  %
                                  "</div>")
                                dir-tree)))))


(defn zip-file [file]
  (let [xml-zipper (zip/xml-zip (slurp file))]
    {:filename   (str file)
     :xml-zipper xml-zipper
     }))


(defn zip-files [src]
  (try
    (let [dir (clojure.java.io/file src)
          dir-tree (file-seq dir)
          file-tree (filter #(.contains (str %) ".")
                            dir-tree)]
      (map #(zip-file %)
           file-tree))
    (catch Exception e
      (println (.getMessage e)))))

(defn xml-to-html [src dest]

  (with-open [wrtr (io/writer dest)]
    (.write wrtr

            (str
              "<html><head><link rel='stylesheet' href='http://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.3.0/styles/default.min.css'><script src='http://cdnjs.cloudflare.com/ajax/libs/highlight.js/9.3.0/highlight.min.js'></script></head><body><script>hljs.initHighlightingOnLoad();</script>"
              (dir-tree-to-html src)

              (clojure.string/join
                (map
                  #(str "<div>"
                        "<h1>"
                        (:filename %)
                        "</h1>"

                        "<pre><code class='xml'>"
                        (-> (first (:xml-zipper %))
                            (clojure.string/replace "<" "&lt;")
                            (clojure.string/replace ">" "&gt;<br/>")
                            (clojure.string/replace "\n" "<br/>")
                            (clojure.string/replace "' " "' <br/>____")
                            (clojure.string/replace "\"" "'"))
                        "</code></pre>"
                        "</div>")
                  (zip-files src)))
              "</body></html>"))))



(unzip "resources/blank.xlsx" "resources/blank")

(xml-to-html "resources/blank" "resources/blank.html")


(defn zip-to-html [src folder-to-delete dest]
  (unzip src folder-to-delete)
  (xml-to-html folder-to-delete dest))


(zip-to-html "resources/legal-draft-invoice-template.xlsx" "resources/legal-draft-invoice-template_temp" "resources/legal-draft-invoice-template.html")


(zip-to-html "resources/bar3.xlsx" "resources/bar3_temp" "resources/bar3.html")



