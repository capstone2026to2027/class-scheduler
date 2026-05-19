import re

with open('templates/admin_password.html', 'r', encoding='utf-8') as f:
    content = f.read()
    lines = f.readlines() if False else content.splitlines(keepends=True)

# Find <script> and </script> positions dynamically
script_start = None
script_end = None
for i, line in enumerate(lines):
    if '<script>' in line and script_start is None:
        script_start = i
    if '</script>' in line and script_start is not None:
        script_end = i
        break

print(f'Script block: lines {script_start+1} to {script_end+1}')

js_lines = lines[script_start:script_end+1]
js = ''.join(js_lines)
js = js.replace('<script>', '').replace('</script>', '')
js = js.replace('{% if current_user.is_super_admin %}true{% else %}false{% endif %}', 'true')
js = js.replace('{{ current_user.id }}', '1')

# Remove template literals first
no_strings = re.sub(r'`[^`\\]*(?:\\.[^`\\]*)*`', '""', js, flags=re.DOTALL)
no_strings = re.sub(r"'(?:[^'\\]|\\.)*'", '""', no_strings)
no_strings = re.sub(r'"(?:[^"\\]|\\.)*"', '""', no_strings)
no_strings = re.sub(r'//[^\n]*', '', no_strings)
no_strings = re.sub(r'/\*.*?\*/', '', no_strings, flags=re.DOTALL)

opens = no_strings.count('{')
closes = no_strings.count('}')
print(f'Open braces: {opens}, Close braces: {closes}, Diff: {opens - closes}')

depth = 0
for i, line in enumerate(no_strings.split('\n')):
    for ch in line:
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1

print(f'Final depth: {depth}')
if depth == 0:
    print('✓ All braces are balanced!')
else:
    print(f'✗ Imbalanced by {depth}')
