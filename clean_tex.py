import os
import re

directory = r'c:\Users\Topo_4\Documents\AT_PFE\memoire'
modified_files = []

for root, dirs, files in os.walk(directory):
    for file in files:
        if file.endswith('.tex'):
            filepath = os.path.join(root, file)
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Replace ' --- ' with ' - '
            new_content = content.replace(' --- ', ' - ')
            new_content = new_content.replace('--- ', '- ')
            new_content = new_content.replace(' ---', ' -')
            
            # Replace double spaces with single space, EXCEPT at the beginning of the line
            lines = new_content.split('\n')
            new_lines = []
            for line in lines:
                match = re.match(r'^(\s*)(.*)', line)
                if match:
                    leading = match.group(1)
                    rest = match.group(2)
                    rest = re.sub(r' {2,}', ' ', rest)
                    new_lines.append(leading + rest)
                else:
                    new_lines.append(line)
            
            new_content = '\n'.join(new_lines)
            
            if new_content != content:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(new_content)
                modified_files.append(file)

print('Modified files:', modified_files)
