# OMU 経済学部 専門演習マッチング業務のためのコード

1. マッチング用のプログラムをダウンロードする。 <https://github.com/kenjisato/omueconMatchKyomu/archive/refs/heads/main.zip>
1. [omueconMatchKyomu-main.zip] というZIPファイルがダウンロードされるので、展開する。フォルダの名前は変えても大丈夫です。例えば、「マッチング2023」のようにします。
1. ファイルの準備をします。
  1. `Data/Admin` に教務事務が管理するゼミリスト(faculty.xlsx) 、学生リスト（students.xlsx）を置いてください。
  1. `Data/Faculty` に NextCloud からダウンロードした教員作成のファイル（[omuid].xlsx）をすべて置いてください。
  1. `Data/Students` に Moodle からダウンロードした学生作成のファイル（[学籍番号....***].xlsx）をすべて置いてください。名前を変更する必要はありません。
1. `matching.RData` というファイルをダブルクリック（またはRのショートカットにドラッグ＆ドロップ）して、R を起動します。 
1. 下のコマンドを順に実行します。



### マッチング計算からはじめる

```r
source("R/0-setup.R", encoding = "utf-8")
source("R/1-matching.R", encoding = "utf-8")
source("R/2-document.R", encoding = "utf-8")
```

`Data/Admin` フォルダに結果のファイル、チェック用のファイルが色々と入っていると思います。（出力先は `config.yml` で変更可能です）


### マッチング計算は完了している

マッチングは完了していて、結果の分析、図の再生成だけしたい場合はこちらです。

```r
source("R/0-setup.R", encoding = "utf-8")
source("R/2-document.R", encoding = "utf-8")
```

### 追加募集の反映

追加募集や諸般の事情で1次結果から変更になるケースです。
`Data/Admin` に教務事務が管理する追加・変更申請者のリスト(change.xlsx) 、を置いてください。change.xlsx は、Student、Seminar、Name の3列構成です。
それぞれ、学籍番号、配属先教員のID、学生氏名を入力します。

なお、学生氏名は必須ではありません。学生リスト（`students.xlsx`）に記載のない学生のみ必要です。学生リストに記載のある学生の名前は、入力しても無視されます。

```r
source("R/0-setup.R", encoding = "utf-8")
source("R/3-amendment.R", encoding = "utf-8")
```

