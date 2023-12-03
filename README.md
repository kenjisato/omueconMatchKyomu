# OMU 経済学部 専門演習マッチング業務のためのコード

1. マッチング用のプログラムをダウンロードする。 <https://github.com/kenjisato/omueconMatchKyomu/archive/refs/heads/main.zip>
1. [omueconMatchKyomu-main.zip] というZIPファイルがダウンロードされるので、展開する。フォルダの名前は変えても大丈夫です。例えば、「マッチング2023」のようにします。
1. ファイルの準備をします。
  1. `Data/Admin` に教務事務が管理するゼミリスト(faculty.xlsx) 、学生リスト（students.xlsx）を置いてください。
  1. `Data/Faculty` に NextCloud からダウンロードした教員作成のファイル（[omuid].xlsx）をすべて置いてください。
  1. `Data/Students` に Moodle からダウンロードした学生作成のファイル（[学籍番号....***].xlsx）をすべて置いてください。名前を変更する必要はありません。
1. `matching.RData` というファイルをダブルクリックして、R を起動します。 
1. 下のコマンドを順に実行します。

```r
install.packages("remotes", repos = "https://cloud.r-project.org")
remotes::install_github("kenjisato/omueconMatch", upgrade = "never")
source("R/matching.R", encoding = "utf-8")
```

`Data/Admin` フォルダに結果のファイル、チェック用のファイルが色々と入っていると思います。
