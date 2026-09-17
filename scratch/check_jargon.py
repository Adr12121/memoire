# -*- coding: utf-8 -*-
import re

terms = ['désambiguïsation', 'paradigme', 'synergie', 'holistique', 'dresse un panorama', 'il est essentiel', 'crucial', 'souveraineté', 'disruption', 'deontologie', 'déontologie']
with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

found = False
for t in terms:
    matches = re.findall(rf'.{{0,30}}{t}.{{0,30}}', text, re.IGNORECASE)
    if matches:
        found = True
        print(f'Term "{t}" found ({len(matches)}):')
        for m in matches[:5]:
            print('  ...', m.strip().replace('\n', ' '), '...')

if not found:
    print('No jargon terms found!')
