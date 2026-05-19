with open('templates/admin_sections.html', 'r', encoding='utf-8') as f:
    content = f.read()

script_start = content.find('<script>')
script_end = content.rfind('</script>')
script = content[script_start:script_end]

lines = script.split('\n')
depth = 0
# Print every line where depth increases, to find the unclosed one
depth_log = []
for i, line in enumerate(lines):
    before = depth
    for ch in line:
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
    if depth != before:
        depth_log.append((i+1, depth, line.strip()[:80]))

# Print the last 30 depth changes
for entry in depth_log[-30:]:
    print(f'Line {entry[0]:4d} depth={entry[1]:3d}: {entry[2]}')

print('\nFinal depth:', depth)
