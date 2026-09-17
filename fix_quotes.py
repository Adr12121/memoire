with open("patch_target_slides.py", "r", encoding="utf-8") as f:
    text = f.read()

# replace r\'\'\' with r""" and ''' with """
text = text.replace("r\\\'\\\'\\\'", 'r"""')
text = text.replace("'''", '"""')

with open("patch_target_slides.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Fixed quotes to r\"\"\"")
