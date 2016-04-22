(ns excel-unzip.core
  (:import [net.lingala.zip4j.core ZipFile])
  (:require [clojure.zip :as zip]
            [clojure.java.io :as io]))























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
           dir-tree))

      )))


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
                            (clojure.string/replace "' " "' <br/> &nbsp;&nbsp;&nbsp;&nbsp;")
                            (clojure.string/replace "\"" "'"))
                        "</code></pre>"
                        "</div>")
                  (zip-files src)))
              "</body></html>"))))



(unzip "resources/bs-stakeout-stats-template.xlsx" "resources/bs-stakeout-stats-template-unzipped")

(xml-to-html "resources/bs-stakeout-stats-template-unzipped" "resources/bs-stakeout-stats-template.html")

(dir-tree-to-html "New Microsoft Excel Worksheet")