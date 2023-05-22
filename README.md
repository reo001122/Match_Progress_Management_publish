# Match_Progress_Management_publish
使用例: https://docs.google.com/spreadsheets/d/1f4s9WsdIkBBvGNPAKiar72qTjhW9R18IDlQFZs3a1C4/edit?usp=sharing
## シートの使い方
### トリガーの設定
- check_by_time():時間ベース（1分毎）
- input_time():スプレッドシートから（変更時）
- input_time_reserve():スプレッドシートから（変更時）
### 試合進行シート（シート名が日付のみもの）
#### シートの見方
シート上部にあるのが、各面の試合進行を表した表です。各試合の2x4のセルについて<br>

<table>
    <tr>
      <td>ID</td>
      <td>(試合種)</td>
      <td>(選手名1)</td>
      <td>(スコア1)</td>
    </tr>
    <tr>
      <td>（ID番号）</td>
      <td>(更新時刻・試合状態)</td>
      <td>(選手名2)</td>
      <td>(スコア2)</td>
    </tr>
 </table>

といった情報を持っています。<br>
更新時刻・試合状態は、スコアが更新されると自動的に記入されます。最終のスコア更新から一定時刻(デフォルトでは7分)経過すると赤字になるので(使用例の8面8Rのように)、
本部から更新忘れないか確認に向かわせると良いでしょう。<br>
セルの色は以下の状態を表しています。

|白色|水色|青色|
|-|-|-|
|試合開始前|試合終了|試合中|

#### シートの操作方法（審判）
審判は担当試合のスコアの更新のみを行うようにしてください。
もし担当するはずの試合が表にない場合は自分で編集せず、本部に問い合わせるようにしてください。

#### シートの操作方法（本部）
基本的に操作は、控えの追加以外はシート下部のボタンからのみ行うようにしてください。<br>
また、試合や控えの追加・変更・削除が行われたときには、該当選手への伝達を忘れないように注意してください。<br>
名簿に存在しない名前が扱われたときには、エラーメッセージが表示され、その名前が名簿シートの方に表示されます。
その場合には正しい名前に訂正し、カウンタも正しく訂正してください。
##### 「ON/OFF」ボタン
こちらのボタンからON状態にしないと、動かないので注意してください。
##### 「適用」ボタン
こちらからコート数を指定できます。
##### 「追加」ボタン
こちらから上部の表へ試合の追加ができます。移動先ID、試合種、選手名を正しく記入した上でボタンを押してください。
##### 「控え移動」ボタン
こちらから、控え一覧から上部の表へ試合を移すことができます。移動させたい試合の控えIDと移動先IDを正しく記入した上でボタンを押してください。
##### 「試合移動」ボタン
こちらからテーブル内で試合の移動をすることができます。試合を行う面が変更する際に利用してください。移動元、移動先のIDを正しく記入した上でボタンを押してください。
##### 「削除」ボタン
こちらから上部の表にある試合を削除することができます。削除したい試合のIDを正しく記入した上でボタンを押してください。
##### 「控え解除」ボタン
こちらから控え一覧にある試合を解除することができます。解除したい控えのIDを正しく記入した上でボタンを押してください。

### 名簿シート（シート名が日付+名簿のもの）
#### シートの見方
シート上部にあるのは、試合進行シートで名簿にない名前が操作されたときにその名前が表示されるボックスです。
ここに名前が表示された場合には手動で変更してください。<br>
シート下部の左と右は、ぞれぞれOBOGと現役の名簿リストです。試合進行シートを操作する前に埋めてください。
セルの色は以下の状態を表しています

|白色|薄緑色|緑色|
|-|-|-|
|控え・試合に入っていない|控えに入っている|試合中|

控えを入れるときは、白色のセルの選手から試合に入った回数や疲労度を考慮して入れるようにしてください。

#### シートの操作方法（本部）
##### 「再カウント」ボタン
カウント表示が何かおかしい場合に、このボタンで再カウントすることができます。ただし、実行に十数秒かかるので注意してください。
