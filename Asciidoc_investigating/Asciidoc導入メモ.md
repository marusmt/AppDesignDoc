# Asciidoc導入メモ

作成：丸山利和(marusmt@me.com)
作成：2018/2/17
更新：2019/5/12
Rubyは導入済の状態
* ruby 2.4.3p205 (2017-12-14 revision 61247) [x64-mingw32]
* gem 2.6.14

以下を参考にインストール

* [asciidoctor-pdfで社内ドキュメントを書こう][1]

#### Asciidoctorのインストール

```
gem install asciidoctor
```
Asciidoctor 1.5.6.1をインストール

#### Asciidoctor-pdfのインストール


```
gem install --pre asciidoctor-pdf
```

Asciidoctor PDF 1.5.0.alpha.16をインストール

[1]:https://qiita.com/gho4d76g/items/302e1ff91754b9b50f34

---

## MacでRubyを最新化してみる

今のrubyの状態はこんな感じ

macOS Hight Sierra 10.13.3

```
ruby 2.3.3p222 (2016-11-21 revision 56859) [universal.x86_64-darwin17]
~ : $ gem list

*** LOCAL GEMS ***

bigdecimal (1.2.8)
CFPropertyList (2.2.8)
did_you_mean (1.0.0)
io-console (0.4.5)
json (1.8.3)
libxml-ruby (2.9.0)
minitest (5.8.5)
net-telnet (0.1.1)
nokogiri (1.5.6)
power_assert (0.2.6)
psych (2.1.0)
rake (10.4.2)
rdoc (4.2.1)
sqlite3 (1.3.11)
test-unit (3.1.5)
```

Macには最初からRubyがインストールされているため、rbenvで複数のバージョンを管理できる様にする。

```
~ : $ brew install rbenv ruby-build
```

バージョンを確認
```
  ~ : $ rbenv --version
  rbenv 1.1.1
```

インストール可能なバージョンを確認する

```
~ : $ rbenv install -l
Available versions:
  1.8.5-p52
  1.8.5-p113
  1.8.5-p114
  1.8.5-p115
  1.8.5-p231
  1.8.6
  1.8.6-p36
  (中略)
  2.5.0-dev
  2.5.0-preview1
  2.5.0-rc1
  2.5.0
  2.6.0-dev
```
2.5.0をインストールしてみる

```
~ : $ rbenv install 2.5.0
```

.bash_profileにパスを追加する。

``` shell
export PATH="$HOME/.rbenv/shims:$PATH"
```
パス設定していないとrbenvでインストールしたrubyのパスが通らない。

インストールしたversionがあるかを確認する。

```
~ : $ rbenv versions
  system
* 2.5.0 (set by /Users/MARUYAMA/.rbenv/version)
```
systemは最初からインストールされているRuby

バージョンを切り替える。

```
~ : $ rbenv global 2.5.0
~ : $ ruby -v
ruby 2.5.0p0 (2017-12-25 revision 61468) [x86_64-darwin17]
```
globalにするとシステム全体に適用される。
localを指定すると一部の適用となる。（そのユーザだけ？）
元々インストールしているバージョンに戻す場合はrbenvでsystemを指定する。

```
rbenv global system
```
[参考1][2]

[参考2][3]

[2]:https://qiita.com/Arashi/items/689e91389235c25088a5
[3]:https://qiita.com/kogache/items/5886a6b62f036c1f94c9
