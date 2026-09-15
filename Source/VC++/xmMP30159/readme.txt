===============================================================================
【名　　称】　MP3操作DLL VBMP3 ソースファイル Ver 1.59
【圧 縮 名】　vbm3s159.lzh
【制作者名】　ミケ
【動作環境】　Windows95/98/NT
　　　　　　　VisualC++ 6.0(SP3) で動作確認済み
【作成方法】　LHA X vbm3s159.lzh
【備    考】　Microsoft VisualC++ を別途用意してください。
-------------------------------------------------------------------------------

  MP3操作DLL VBMP3 ソースファイル

  GPL に基づき VBMP3 のソースを公開します。
  かなり汚いソースで申し訳ありませんが参考になれば嬉しいです。

  ※VCMP3 は VBMP3 と同一の物です。
    VCMP3 lib 版は、定義が多少異なりますのでプロジェクトも含めた
    パッケージ(vcmp3lib_src.lzh)として同梱しました。


■使用方法

  ディレクトリの例
  
  \PRG
    \mp3dec      : デコードライブラリmp3dec用ディレクトリ
      \include   : mp3dec.h を格納
      \Release   : mp3dec.lib を格納
    \VBMP3       : VBMP3用ディレクトリ(プロジェクトファイル等を格納)
      \src       : cpp ファイルなど格納
      \include   : ヘッダファイル格納
    \VCMP3       : VCMP3用ディレクトリ(プロジェクトファイル等を格納)
      \src       : cpp ファイルなど格納
      \include   : ヘッダファイル格納

  ・前準備
    (1)別途、ホームページ(*1)に公開している mp3dec.lib を
       ダウンロードしておき、上記の例に沿い格納
    (2)VBMP3を上記の例にそって格納(ディレクトリ付き解凍)

  ・プロジェクトファイル使用
    (1)VC++ よりプロジェクトファイル VBMP3.dsw を読み込む

  ・新規にプロジェクト作成
    (1)Win32 Dynamic-Link Library を選択しプロジェクトを作成
    (2) (1)で作成されたディレクトリ内に解凍
    (3)ソースをプロジェクトに追加
    (4)プロジェクトの設定変更
       i   [プロジェクト]-[設定] を選択
       ii  C/C++ タブを選択
       iii カテゴリリストより「コード生成」を選択
       iv  使用するランタイムライブラリリストより「マルチスレッド」
           (もしくは「マルチスレッド(デバッグ)」)を選択

   *1 : 注意事項の「4.サポート」を参照して下さい

■注意事項

  1. 著作権

  ・ このライブラリの著作権は、Xing Technology, GoodNoise, 高橋政雄氏, 
     ミケこと、苅込大輔が保有しています

  ・ このライブラリはフリーウェアですが、オリジナルの転載は禁止とします。
     その他、改変、配布、公開する場合、GPLに基づいて行って下さい。
     （3. GPL について)


  2. 免責

  ・  このライブラリの使用による、いかなる損害、
      その他についても、制作者はその責任を負いません。

  ・ バグ等があったとしても、これを訂正する義務は負いません。


  3. GPL について

     本ライブラリのデコードエンジンは、当初は Visual C++で作るMP3 Player に
     付属している Xingの古いデコードエンジンを元にしたライブラリを利用してい
     ました。（＊この時点では、出版社側が許可を得ています。）

     しかし、途中のバージョンより、VBRやグラフィックイコライザ対応のため、
     デコードエンジン部分に更に変更を加えました。
     GPLに基づいているため、オープンソースとします。
     このソースコードを改変して配布する場合は、GPLに基づきますので、注意して
     ください。

     GPLの詳細についてはGNU 一般公有使用許諾書をお読み下さい。

     MP3 Decoder originally Copyright (C) 1995-1997 Xing Technology
     Corp.  http://www.xingtech.com

     Portions Copyright (C) 1998 GoodNoise
     Portions Copyright (C) 1999 Masao Takahashi
     Portions Copyright (C) 1999-2000 Daisuke Karikomi

     This program is free software; you can redistribute it and/or modify
     it under the terms of the GNU General Public License as published by
     the Free Software Foundation; either version 2 of the License, or
     (at your option) any later version.

     This program is distributed in the hope that it will be useful,
     but WITHOUT ANY WARRANTY; without even the implied warranty of
     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
     GNU General Public License for more details.

     You should have received a copy of the GNU General Public License
     along with this program; if not, write to the Free Software
     Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA


  4. サポート

     最新版の公開等は、以下のページより行っています。
     http://www.angel.ne.jp/~mike/

     このライブラリに関し、ご意見、ご質問等がありましたら
     掲示板もしくはミケ宛にメールして下さい。
     しかし、すべてのロジックを理解しているというわけでは
     ありませんので、お答えできないこともあります。
     あらかじめご了承願います。
    （忙しく返事が出来ないときはご容赦ください）

     e-mail : mike@angel.ne.jp

■参考文献

  このライブラリを改変する上で以下の書籍を参考にさせて頂きました。

  ●MP3関連
    高橋政雄著 VisualC++ 6.0 でつくる MP3 Playerプログラミング エーアイ出版

    月刊Ｃマガジン ソフトバンク パブリッシング社

  ●DLL関連
    北山洋幸著 技術者のための VisualC++ 実践プログラミング技法 技術評論社

■謝辞

  ・オリジナルのデコードライブラリをフリーで公開して下さいました
    Xing Technology に心より感謝致します。

  ・FreeAmp と言う有用なソフトのソースを公開して下さった GoodNoise の皆様に
    心から感謝致します。
    http://www.freeamp.org

  ・「VisualC++ 6.0 でつくる MP3 Playerプログラミング」という VC++ の知識に
     乏しい私でも分かりやすい本を書いていただき、更に MPEG 2.5 / VBR に対応した
     デコードライブラリを提供して下さいました高橋 政雄氏に心から感謝致します。

  ・スペクトラムアナライザを実現するために Reliable Software の
    FFTアルゴリズムを利用させて頂きました。
    高速なアルゴリズムを提供して下さいました Reliable Software の皆様に
    心から感謝致します。


[vbm3s159.lzh]
