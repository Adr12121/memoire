# -*- coding: utf-8 -*-
with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

import re
titles = re.findall(r'init_standard_slide\([^\)]+\)', text)
for t in titles:
    print(t)
