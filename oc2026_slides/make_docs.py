# -*- coding: utf-8 -*-
"""spec.json から進行表と写真差し込み一覧を生成する"""
import json, sys

spec = json.load(open("spec.json", encoding="utf-8"))
S = spec["slides"]

def mmss(sec):
    return f"{sec // 60}:{sec % 60:02d}"

# ---------------------------------------------------------------- 進行表
cuts = sorted([s for s in S if s.get("cut_rank")], key=lambda x: x["cut_rank"])
full = sum(s["est_seconds"] for s in S)

lines = ["# 当日の進行表 ― 2026/8/23 オープンキャンパス 学科紹介", "",
         f"全26枚・フルで話すと **約{mmss(full)}**（質疑を除く）。",
         "会場の持ち時間に合わせて、下の「削る順番」の上から落としてください。", "",
         "## 時間配分（フル版）", "",
         "| # | 区分 | 見出し | 目安 | 累計 | 削れる |",
         "|---:|---|---|---:|---:|:--:|"]
acc = 0
for s in S:
    acc += s["est_seconds"]
    mark = f"{s['cut_rank']}番目" if s.get("cut_rank") else ("" if not s.get("cuttable") else "○")
    title = s["title"].replace("\n", " ")
    lines.append(f"| {s['no']} | {s['section']} | {title} | {s['est_seconds']}秒 | {mmss(acc)} | {mark} |")

lines += ["", "## チェックポイント（時計を見るのはこの3回だけ）", ""]
marks = {}
for s in S:
    if "4年間の階段" in s["title"]:
        marks["階段を話し終える"] = s["no"]
    if "お金の話" in s["title"]:
        marks["95万円を話し終える"] = s["no"]
    if "授業料 半額免除" in s["title"] and s["section"] == "3":
        marks["入試を話し終える"] = s["no"]
acc = 0
cum = {}
for s in S:
    acc += s["est_seconds"]
    cum[s["no"]] = acc
for k, no in marks.items():
    lines.append(f"- **{mmss(cum[no])}** … {k}（{no}枚目）")

lines += ["", "## 押したときに削る順番", "",
          "上から順に落とすと、いちばん効く部分（階段・95万円・英語が苦手・入試・まとめ）が最後まで残ります。", ""]
tot = full
for s in cuts:
    tot -= s["est_seconds"]
    lines.append(f"{s['cut_rank']}. **{s['no']}枚目「{s['title']}」** を飛ばす（−{s['est_seconds']}秒 → 約{mmss(tot)}）")
extra = [(6, [5, 6], "経済学クイズと答え（2枚まとめて）"),
         (7, [9], "教員8名 ―「先生は8人です」と口頭だけにする"),
         (8, [2], "アイスブレイク")]
by_no = {s["no"]: s for s in S}
for rank, nos, label in extra:
    tot -= sum(by_no[n]["est_seconds"] for n in nos)
    lines.append(f"{rank}. **{'・'.join(str(n) for n in nos)}枚目「{label}」** を飛ばす → 約{mmss(tot)}")

hide = [s["cut_rank"] and s["no"] for s in S if s.get("cut_rank")]
hide = sorted(hide) + [5, 6, 9, 2]
hide = sorted(set(hide))
left = full - sum(by_no[n]["est_seconds"] for n in hide)
lines += ["", "## 20分コース（持ち時間が20分のとき）", "",
          f"開始前に次の**{len(hide)}枚を非表示**にしてください（PowerPoint：スライド一覧で右クリック → 非表示スライド）。",
          "",
          "```",
          "非表示にする枚：" + "、".join(str(n) for n in hide),
          "```",
          "",
          f"残り**{len(S) - len(hide)}枚・約{mmss(left)}**。質疑を入れて20〜22分に収まります。",
          "",
          "残る枚：" + "、".join(f"{s['no']}（{s['title'].replace(chr(10),' ')}）"
                                 for s in S if s["no"] not in hide),
          "",
          "## 絶対に削らない8枚", ""]
keep = ["1（表紙・写真の伏線）", "3（アジェンダ・数字の伏線）", "10（区切り／ここからが本題）",
        "11（4年間の階段）", "14（TOP10の写真・伏線回収）", "16（95万円）",
        "18（英語が苦手でも大丈夫）", "24（入試・授業料半額免除）", "25（まとめ）", "26（クロージング）"]
lines += [f"- {k}" for k in keep]
lines += ["", "## 沈黙を置く3か所（4か所以上にすると「間」ではなく「滞り」になります）", "",
          "- 14枚目：写真を出したあと **5秒**",
          "- 16枚目：95万円を出したあと **3秒**",
          "- 24枚目：「語学の条件をもう満たしています」のあと **2秒**", ""]
open("進行表.md", "w", encoding="utf-8").write("\n".join(lines))

# ---------------------------------------------------------------- 写真一覧
pl = ["# 写真の差し込み一覧", "",
      "旧ファイルの写真はこの環境から取り出せないため、スライドには**色面のプレースホルダ**を置いてあります。",
      "旧版（`国際経済学科紹介_オープンキャンパス2026_改.pptx` など）から写真をコピーして、",
      "下の位置に貼り付けてください。**貼り付けたら、枠の中の灰色の指示テキストは削除**してください。", "",
      "写真を入れなくても色面として成立するので、時間がなければそのままでも破綻しません。", "",
      "| 枚 | 見出し | 入れる写真 | キャプション |", "|---:|---|---|---|"]
for s in S:
    if s.get("photo_slot"):
        pl.append(f"| {s['no']} | {s['title']}".replace("\n", " ")
                  + f" | {s['photo_slot']}".replace("\n", " ")
                  + f" | {s.get('photo_caption','')}".replace("\n", " ") + " |")
pl += ["", "## 貼り付けの手順（PowerPoint）", "",
       "1. 旧ファイルを開き、使う写真を選んで **Ctrl+C**",
       "2. 新ファイルの該当スライドで **Ctrl+V**",
       "3. 写真を選んだまま **図の形式 → トリミング → 塗りつぶし** で枠に合わせる",
       "4. **右クリック → 最背面へ移動** すると、キャプションが写真の上に出ます",
       "5. 灰色の「［写真を差し込む位置］…」のテキストボックスを選んで削除",
       "",
       "## 表紙の写真だけは必ず入れてください", "",
       "表紙のニュージーランドの写真は、14枚目「実際に、行ってきました」で",
       "「最初にお見せした写真、覚えていますか」と**答え合わせをする伏線**になっています。",
       "この2枚は同じ写真、または同じ研修の写真にしてください。", ""]
open("写真の差し込み一覧.md", "w", encoding="utf-8").write("\n".join(pl))
print("進行表.md / 写真の差し込み一覧.md を書き出しました")
