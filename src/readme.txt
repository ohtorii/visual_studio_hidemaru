============================================================================
visual studio hidemaru
============================================================================
（ファイル）
readme.txt					このファイル
make_exe.bat				実行ファイル(.exe)を作るバッチファイル
setup.py					py2exeで使用する設定ファイル
visual_studio_hidemaru.py	本体
license.txt					ライセンス


（必要なもの）
python 	ver 2.7.1
py2exe	ver 0.6.9

python ver3系では動作しません、必ずver2系を使用してください。

（実行ファイルの作り方）
コマンドプロンプトから make_exe.bat を起動してください。
distディレクトリに実行ファイル(.exe)が作られます、秀丸のマクロディレクトリ
にコピーしてください。

（内部実装）
COM の RunningObjectTable から起動しているVisual Studioを列挙します。複数のVisual Studioが起動していても一意に識別できます。
 （_get_dte_from_pid関数 / cmd_dte_list関数など）
これが全てと言っても過言ではありません。


RunningObjectTable は IROTVIEW.EXE で確認できます。（最近のVisual Studioには入っていません、確かVisual Studio6に入っていたような・・・）

cmd_te_*関数はソースコードの絶対パス名からVisual Studioを探し出して、ビルドや実行を行なうテキストエディタ向けの関数です。

（その他）
今回はpythonで作りましたが、C#の方がサンプルも多く作りやすいかもしれません。


（謝辞）
visual_studio.vim を参考にさせてもらいました。
http://www.vim.org/scripts/script.php?script_id=864


（連絡先）
http://d.hatena.ne.jp/ohtorii/
