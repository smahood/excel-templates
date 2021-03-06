(defproject com.infolace/excel-templates "0.3.2"
  :description "Build Excel files by combining a template with plain old data"
  :url "https://github.com/tomfaulhaber/excel-templates"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.apache.commons/commons-lang3 "3.4"]
                 [org.apache.poi/poi-ooxml "3.10-FINAL"]
                 [org.apache.poi/ooxml-schemas "1.1"]
                 [org.clojure/data.zip "0.1.1" :exclusions [[org.clojure/clojure]]]
                 [joda-time "2.7"]
                 [org.clojure/clojure "1.7.0"]
                 [net.lingala.zip4j/zip4j "1.3.2"]
                 [org.clojure/data.xml "0.1.0-beta1"]
                 ]
  :profiles {:repl {:dependencies [[org.clojure/clojure "1.7.0"]]}})
