===============================================================================
【名　　称】　Mpeg audio Decode Library mp3dec
【圧 縮 名】　mp3dec.lzh
【制作者名】　ミケ
            　高橋 政雄氏
            　GoodNoise
            　Xing Technology
【動作環境】　Windows95/98/NT
　　　　　　　VisualC++ 6.0(SP3) で動作確認済み
【作成方法】　LHA X mp3dec.lzh
【備    考】　Microsoft VisualC++ を別途用意してください。
-------------------------------------------------------------------------------

  Mpeg audio Decode Library mp3dec

  VBMP3 で使用させて頂いている mp3dec.lib を GPL に基づきソース公開します。


■特徴

  ・MPEG-1/2 Layer1～3 及び、MPEG2.5, VBR 形式にも対応

  ・デコードオプションによりダウンサンプリング等が可能

  ・グラフィックイコライザに対応


■使用方法

  ・プロジェクトファイル使用
    (1)ファイルを適当なディレクトリに解凍(ディレクトリ付き解凍)
    (2)VC++ よりプロジェクトファイル mp3dec.dsw を読み込む

  ・新規にプロジェクト作成
    (1)Win32 Static Library を選択し、次の画面は何もチェックせず終了
    (2) (1)で作成されたディレクトリ内に解凍
    (3)ソースをプロジェクトに追加
    (4)プロジェクトの設定変更
       i   [プロジェクト]-[設定] を選択
       ii  C/C++ タブを選択
       iii カテゴリリストより「コード生成」を選択
       iv  使用するランタイムライブラリリストより「マルチスレッド」
           (もしくは「マルチスレッド(デバッグ)」)を選択

   ※参考
     VBMP3 では、mp3dec.h で定義された関数のみ使用しています。
     各関数のインタフェースは後述の参考文献に付属のライブラリと
     ほぼ変わりませんので参考にされることをお勧めします。


■注意事項

  1. 著作権

  ・ このライブラリの著作権は、Xing Technology, GoodNoise, 高橋政雄氏, 
     ミケこと、苅込大輔が保有しています

  ・ このライブラリはフリーウェアですが、改変、転載、配布、公開する場合、
     GPLに従って下さい。（3. GPL について)


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
     デコードエンジン部分は、GPLに基づいているため、オープンソースとします。
     このソースコードを改変して配布する場合は、GPLに基づきますので、注意して
     ください。

     GPLの詳細についてはGNU 一般公有使用許諾書をお読み下さい。

     MP3 Decoder originally Copyright (C) 1995-1997 Xing Technology
     Corp.  http://www.xingtech.com

     Portions Copyright (C) 1998 GoodNoise
     Portions Copyright (C) 1999 Masao Takahashi
     Portions Copyright (C) 1999 Daisuke Karikomi

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

     このライブラリに関し、ご質問等がありましたらミケ宛にメールして下さい。
     しかし、私自身ブラックボックス的な処理が多く明確に回答できないことが
     ありますのであらかじめご了承願います。
    （忙しく返事が出来ないときはご容赦ください）

     e-mail : mike@angel.ne.jp

■参考文献

  このライブラリを改変する上で以下の書籍を参考にさせて頂きました。

  ●MP3関連
    高橋政雄著 VisualC++ 6.0 でつくる MP3 Playerプログラミング エーアイ出版


■謝辞

  ・オリジナルのデコードライブラリをフリーで公開して下さいました
    Xing Technology に心より感謝致します。

  ・FreeAmp と言う有用なソフトのソースを公開して下さった GoodNoise の皆様に
    心から感謝致します。
    http://www.freeamp.org

  ・「VisualC++ 6.0 でつくる MP3 Playerプログラミング」という VC++ の知識に
    乏しい私でも分かりやすい本を書いていただき、更に MPEG 2.5 / VBR に対応した
    デコードライブラリを提供して下さいました高橋 政雄氏に心から感謝致します。

[mp3dec.lzh]
