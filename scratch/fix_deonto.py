with open('generate_final_180s_deck.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('("Rigueur juridique et déontologique",', '("Rigueur réglementaire et secret professionnel",')

with open('generate_final_180s_deck.py', 'w', encoding='utf-8') as f:
    f.write(text)

print("Remplacement reussi !")
