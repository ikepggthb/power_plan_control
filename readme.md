# power_plan_controler

特定のアプリが起動しているとき、Windowsの電源プランを高パフォーマンスに設定するスクリプト。

# note

## 使い方
main関数内の "high_performance_process" に、高パフォーマンスで実行したいプロセス名を入れて実行する

``` python
def main():
    # 高パフォーマンスで実行したいプロセス名を入れる
    high_performance_process : list = ["firefox.exe","r5apex.exe"]
    main_loop(high_performance_process)
```

タスクスケジューラー等でログイン時に自動的に実行されるようにすると良い。

# 環境

- Windows Vista以上
- Python 3.7以上

