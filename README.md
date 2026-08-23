# 裝彈、射擊與防禦遊戲：無限期零和賽局分析

這個專案包含高中科展報告所使用的精確驗證、浮點數探索、圖表產生與 Word 報告產生程式。

## 遊戲規則

每回合兩位玩家同時選擇：

- `L`（Load）：裝一發子彈。
- `S`（Shoot）：有子彈時消耗一發並射擊。
- `D`（Defend）：擋住對方該回合的射擊，但不能連續使用。

若一方射擊，而另一方既未射擊也未防禦，射擊者立即獲勝；雙方同時射擊或射擊遇到防禦時，遊戲繼續。

## 原始碼檔案

- `exact_game.py`：以 `fractions.Fraction` 做有限期上下截止遞迴，並精確驗證低資源 37 個有序狀態的漂移與逃離機率不等式。
- `explore_game.cpp`：使用雙精度浮點數探索有限期上下界、零截止值與第二回合策略的數值趨勢。
- `make_report_figures.py`：自動編譯 `explore_game.cpp`，產生 `report_data.json`、`figure_gap.png` 與 `figure_strategy.png`。
- `build_full_report.py`：根據資料與圖片建立完整 Word 科展報告。
- `requirements.txt`：Python 套件需求。

## 已附與可重新產生的輸出

- `report_data.json`：本次報告使用的數值輸出，並可由 `make_report_figures.py` 重新產生。
- `figure_gap.png`、`figure_strategy.png`：由數值資料產生的兩張圖。
- `game_theory_science_fair_report.docx`：由 `build_full_report.py` 產生的報告成品；若 GitHub 儲存庫未附二進位成品，仍可依下方步驟在本機重建。

## 環境需求

- Python 3.10 以上
- 支援 C++17 的 `g++`
- Python 套件：`python-docx`、`matplotlib`

安裝 Python 套件：

```bash
python -m pip install -r requirements.txt
```

## 重現步驟

### 1. 執行精確核心驗證

```bash
python exact_game.py 12 verify
```

重要輸出應包含：

```text
core N 12 states 37 bad 0
CERTIFIED level 10, 4 steps: min_escape > 9/100
CERTIFIED level 12, 2 steps: min_escape > 1/2
```

這部分使用有理數分數運算；`bad 0` 與兩個 `CERTIFIED` 判斷不是由四捨五入小數決定。

### 2. 重新產生浮點資料與圖表

```bash
python make_report_figures.py
```

程式會自動執行：

```bash
g++ -std=c++17 -O2 -DNDEBUG explore_game.cpp -o explore_game
```

並產生：

- `report_data.json`
- `figure_gap.png`
- `figure_strategy.png`

### 3. 重新建立 Word 報告

```bash
python build_full_report.py
```

輸出檔為 `game_theory_science_fair_report.docx`。報告指定中文標楷體、英文 Times New Roman；若電腦沒有標楷體，Word 可能使用替代字型。

## 嚴格證明與數值探索的區別

- `exact_game.py 12 verify` 使用精確分數來認證有限核心中的不等式，是完整收斂證明的一個有限電腦輔助步驟。
- `explore_game.cpp` 與兩張圖使用雙精度浮點數，只用來顯示趨勢及提供後續區間算術的目標。
- 因此，像 `0.2791962356`、`0.3076983951`、`0.4131053693` 這些數字目前是高精度浮點估計，不應寫成已嚴格證明到十位小數。

## 主要結論

報告證明悲觀截止值與樂觀截止值在每個合法狀態收斂到同一極限，因此平手截止的有限期值也收斂，而且該共同極限是無限期遊戲的值。在雙方第一回合都裝彈的條件下，第二回合狀態為 `(3,3)`；若

```text
A = V(5,2),  B = V(2,1),
```

則該狀態的極限混合策略為

```text
(L, S, D) = (B, A, 1) / (1 + A + B).
```
