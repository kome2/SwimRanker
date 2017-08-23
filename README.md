# SwimRanker
Swimming Ranking Generator which combines the results of high school swim meets of each prefectures

日本国内における、都道府県の高校県大会を結合し、一つのエクセルのデータを作成できる。
出力できるデータはdataディレクトリ内のconfig.txtを編集することで任意の大会に変更が可能。

config.txtは、通し番号と都道府県名、そして各県の大会コードの下3桁に対応している。
大会コードはスイムレコードどっとこむの該当大会URLから参照できる。

また、標準記録調査も同時に行っており、予選種目においては、インターハイ標準記録を突破した記録のみを抽出、決勝とタイム決勝は全競技抽出としている。

標準記録についても、standards.txtを編集することで任意の記録に変更可能である。
標準記録は秒表示となっている。
