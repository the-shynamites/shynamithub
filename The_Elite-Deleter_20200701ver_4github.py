# The_Elite-Deleter-20200701ver.py

# coding:utf-8
import re
import time

# open and read a file in HPML_mix.
# ファイル名を入力してEnterで確定する。
print("PLZ choose your HPML-TXT file, and press Enter-key.")
HPML_Text = input()
file = open(HPML_Text, 'r', encoding='utf-8')

str = file.read()

# 4秒間の待機
time.sleep(4)

print("Now Running.")

file.close()


# 前処理として、HPMLファイルの記号を整える。

print("Now Regulating your HPML.")

# 半角カギ括弧を全角カギ括弧に置き換える
text_mod1 = re.sub('｢', "「", str)
text_mod2 = re.sub('｣', "」", text_mod1)

# 2秒間の待機
time.sleep(2)

# 全角丸括弧を半角丸括弧に置き換える
text_mod3 = re.sub('（', "(", text_mod2)
text_mod4 = re.sub('）', ")", text_mod3)

# 2秒間の待機
time.sleep(2)

# 半角句読点を全角句読点に置き換える
text_mod5 = re.sub('､', "、", text_mod4)
text_mod6 = re.sub('｡', "。", text_mod5)

# 2秒間の待機
time.sleep(2)


# open and write a new file in regulated HPML.
file = open(HPML_Text+'_regulated.txt', 'w', encoding = 'utf-8')

file.write(text_mod6)

file.close()


# from regulated-HPML to SB.
text_mod7 = re.sub('「|」', "", text_mod6)

# 4秒間の待機
time.sleep(4)


print("from regulated HPML to SB mix.")

# open and write a file in SB_mix.
file = open(HPML_Text+'_SB_mix.txt', 'w', encoding = 'utf-8')

file.write(text_mod7)

file.close()

# from SB to F.
text_mod8 = re.sub(r'\([^)]+\)', "", text_mod7)

# 4秒間の待機
time.sleep(4)


print("from regulated HPML to F mix.")

# open and write a file in F_mix.
file = open(HPML_Text+'_F_mix.txt', 'w', encoding = 'utf-8')

file.write(text_mod8)

file.close()

# “The Elite-Deleter in Python 3.8” produced by K.masamix “KIXAN” as the SHYNAMITES/La CHENAMITOJ, 2020/07/01.

