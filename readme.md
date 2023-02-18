# power_plan_controler

特定のアプリが起動しているとき、Windowsの電源プランを「高パフォーマンス」または「バランス」に設定し、そうでないときは「省電力」に設定するスクリプト。
タスクトレイに常駐しており、一時的に任意の電源プランに変更することも可能。


## 使い方
スクリプトの冒頭には、```high_performance_app_map```と```balanced_app_map```というリストがあります。
```high_performance_app_map```には、「高パフォーマンス」で実行したいプロセス名を、```balanced_app_map```には「バランス」で実行したいプロセス名をそれぞれ入力してください。

``` python
high_performance_app_map : list = ["r5apex"]
balanced_app_map         : list = ["firefox","Chrome","Code"]
```

```high_performance_app_map```にあるアプリケーションが起動している場合、電源プランを高パフォーマンスに設定します。そうでない場合、```balanced_app_map```にあるアプリケーションが起動している場合、電源プランをバランスに設定します。どちらでもない場合は、電源プランを省電力に設定します。

ログイン時に自動的に実行するよう、タスクスケジューラーなどを利用すると良いでしょう。

## 環境

- Windows Vista以上
- Python 3.7以上
- Pyside 6

