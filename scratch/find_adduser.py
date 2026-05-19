with open('templates/admin_password.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()
for i, line in enumerate(lines):
    if 'adduser-panel' in line or 'adduser_panel' in line or 'Add User' in line or 'createUserForm' in line:
        print(f'{i+1}: {line.strip()}')
