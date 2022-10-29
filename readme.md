秀丸エディからVisual Studioを制御するマクロ
========

# 始めに
- 秀丸エディタから VisualStudio を制御するプログラマ向けのマクロです。
- ソースコードを秀丸エディタで書いて、ビルドする度に VisualStudio へ移動するのが面倒なので、「ビルド・リビルド・実行...etc」を秀丸エディタから出来るようにしました。

# 対応しているVisualStudio
- Microsoft VisualStudio Community 2022
- Microsoft VisualStudio Community 2017
- VisualStudio 2008／VisualStudio 2010／VisualStudio 2013／VisualStudio 2015それぞれExpress版は未対応）

それぞれ、C++/C#/VBのソリューションで動作を確認しています。

# 動作環境
- 秀丸エディタ ver8.03以降

# ディレクトリ構成
		macro       秀丸マクロ
		src         pythonのソースコード
		doc         ドキュメント類

# 動作イメージ
- ![Alt text](http://cdn-ak.f.st-hatena.com/images/fotolife/o/ohtorii/20110402/20110402135007.png)

# 謝辞
[visual_studio.vim](https://www.vim.org/scripts/script.php?script_id=864)のソースコードをかなり参考にさせて頂きました、ありがとうございます。

以上
