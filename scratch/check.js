const isSuperAdmin = {% if current_user.is_super_admin %}true{% else %}false{% endif %};
const currentUserId = {{ current_user.id }};
function handleResetToDefault() {
    const username = document.getElementById('userInput').value.trim();
    if (!username) {
        showToast('Please select or search a target user account first.', 'error');
        return;
    }
    
    if (confirm(`Are you sure you want to reset the password for user "${username}" to the default value "123456"?`)) {
        // Set password value to 123456
        const passwordInput = document.getElementById('newPasswordInput');
        passwordInput.value = '123456';
        passwordInput.removeAttribute('required'); // Prevent validation blockage
        
        // Trigger AJAX submit via submit event dispatch
        const form = document.getElementById('securityUpdateForm');
        form.dispatchEvent(new Event('submit'));
    }
}

function toggleManualPasswordFlow() {
    const section = document.getElementById('manualPasswordSection');
    const btn = document.getElementById('manualFlowBtn');
    const passwordInput = document.getElementById('newPasswordInput');
    
    if (section.style.display === 'none' || section.style.display === '') {
        section.style.display = 'block';
        passwordInput.setAttribute('required', 'true');
        btn.classList.add('active');
        
        // Scroll smoothly to section
        setTimeout(() => {
            section.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 100);
    } else {
        section.style.display = 'none';
        passwordInput.removeAttribute('required');
        btn.classList.remove('active');
    }
}

function selectUserForReset(username) {
    const userSelect = document.getElementById('userSelect');
    const userInput = document.getElementById('userInput');
    
    if (userSelect) {
        userSelect.value = username;
    }
    if (userInput) {
        userInput.value = username;
    }
    
    // Smoothly scroll to target input section
    document.getElementById('securityUpdateForm').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    showToast(`Target account set to: ${username}`);
}

// Client Side Pagination State
let currentPage = 1;
const pageSize = 10;

function filterAccounts() {
    const query = document.getElementById('dirSearchInput').value.toLowerCase().trim();
    const role = document.getElementById('dirRoleSelect').value.toLowerCase();
    
    const rows = Array.from(document.querySelectorAll('.account-row'));
    let filteredRows = [];
    
    rows.forEach(row => {
        const username = row.getAttribute('data-username').toLowerCase();
        const userRole = row.getAttribute('data-role').toLowerCase();
        
        const matchesQuery = username.includes(query);
        const matchesRole = role === '' || userRole === role;
        
        if (matchesQuery && matchesRole) {
            filteredRows.push(row);
        } else {
            row.style.display = 'none';
        }
    });
    
    // Calculate pages
    const totalFiltered = filteredRows.length;
    const totalPages = Math.ceil(totalFiltered / pageSize) || 1;
    
    if (currentPage > totalPages) {
        currentPage = totalPages;
    }
    if (currentPage < 1) {
        currentPage = 1;
    }
    
    // Show only active page rows
    const startIdx = (currentPage - 1) * pageSize;
    const endIdx = startIdx + pageSize;
    
    rows.forEach(row => {
        row.style.display = 'none';
    });
    
    filteredRows.forEach((row, idx) => {
        if (idx >= startIdx && idx < endIdx) {
            row.style.display = '';
        }
    });
    
    // Update count dynamically
    const countBadge = document.getElementById('totalAccountsCount');
    if (countBadge) {
        countBadge.textContent = `${totalFiltered} of ${rows.length} Accounts`;
    }
    
    // Update Pagination controls
    updatePaginationUI(currentPage, totalPages);
}

function updatePaginationUI(page, totalPages) {
    const container = document.getElementById('paginationContainer');
    if (!container) return;
    
    if (totalPages <= 1) {
        container.style.display = 'none';
        return;
    }
    
    container.style.display = 'flex';
    container.innerHTML = `
        <button type="button" onclick="changePage(${page - 1})" class="paginate-btn" ${page === 1 ? 'disabled style="opacity: 0.5; pointer-events: none;"' : ''}>
            <i data-lucide="chevron-left" style="width: 16px; height: 16px;"></i> Prev
        </button>
        <span style="font-size: 0.875rem; font-weight: 600; color: var(--text-main);">Page ${page} of ${totalPages}</span>
        <button type="button" onclick="changePage(${page + 1})" class="paginate-btn" ${page === totalPages ? 'disabled style="opacity: 0.5; pointer-events: none;"' : ''}>
            Next <i data-lucide="chevron-right" style="width: 16px; height: 16px;"></i>
        </button>
    `;
    if (window.lucide) {
        lucide.createIcons();
    }
}

function changePage(newPage) {
    currentPage = newPage;
    filterAccounts();
}

function clearDirFilters() {
    document.getElementById('dirSearchInput').value = '';
    document.getElementById('dirRoleSelect').value = '';
    currentPage = 1;
    filterAccounts();
}

function switchAccessTab(panelId, btn) {
    document.querySelectorAll('.tab-panel').forEach(p => p.style.display = 'none');
    document.getElementById(panelId).style.display = 'block';
    
    document.querySelectorAll('.tab-btn').forEach(b => {
        b.classList.remove('active');
        b.style.color = 'var(--text-muted)';
        b.style.borderBottomColor = 'transparent';
        b.style.fontWeight = '600';
    });
    
    btn.classList.add('active');
    btn.style.color = 'var(--primary)';
    btn.style.borderBottomColor = 'var(--primary)';
    btn.style.fontWeight = '700';
    
    if (window.lucide) {
        lucide.createIcons();
    }
}

function toggleAddUserRoleFields(role) {
    const fieldsRow = document.getElementById('createUserRoleFields');
    const teacherDept = document.getElementById('teacherDeptField');
    const studentSection = document.getElementById('studentSectionField');
    
    if (role === 'admin') {
        fieldsRow.style.display = 'none';
        teacherDept.style.display = 'none';
        studentSection.style.display = 'none';
    } else if (role === 'teacher') {
        fieldsRow.style.display = 'grid';
        teacherDept.style.display = 'block';
        studentSection.style.display = 'none';
    } else if (role === 'student') {
        fieldsRow.style.display = 'grid';
        teacherDept.style.display = 'none';
        studentSection.style.display = 'block';
    }
    
    if (window.lucide) {
        lucide.createIcons();
    }
}

function setAddUserPasswordMode(mode) {
    const autoBtn = document.getElementById('createUserPassAutoBtn');
    const manualBtn = document.getElementById('createUserPassManualBtn');
    const passBlock = document.getElementById('createUserPasswordBlock');
    const passInput = document.getElementById('createUserPasswordInput');
    const modeInput = document.getElementById('createUserPassMode');
    const submitBtn = document.getElementById('createUserSubmitBtn');
    
    modeInput.value = mode;
    
    if (mode === 'auto') {
        autoBtn.className = 'reset-action-btn active';
        autoBtn.style.borderColor = 'var(--primary)';
        autoBtn.style.color = 'var(--primary)';
        autoBtn.style.background = '#eff6ff';
        
        manualBtn.className = 'change-action-btn';
        manualBtn.style.borderColor = 'var(--border)';
        manualBtn.style.color = '#475569';
        manualBtn.style.background = '#f8fafc';
        
        passBlock.style.display = 'none';
        passInput.removeAttribute('required');
        
        if (submitBtn) {
            submitBtn.removeAttribute('disabled');
            submitBtn.style.opacity = '1';
            submitBtn.style.pointerEvents = 'auto';
        }
    } else {
        manualBtn.className = 'change-action-btn active';
        manualBtn.style.borderColor = 'var(--primary)';
        manualBtn.style.color = 'var(--primary)';
        manualBtn.style.background = '#eff6ff';
        
        autoBtn.className = 'reset-action-btn';
        autoBtn.style.borderColor = 'var(--border)';
        autoBtn.style.color = '#475569';
        autoBtn.style.background = '#f8fafc';
        
        passBlock.style.display = 'block';
        passInput.setAttribute('required', 'true');
        
        if (passInput) {
            passInput.dispatchEvent(new Event('input'));
        }
        
        setTimeout(() => {
            passBlock.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }, 100);
    }
    
    if (window.lucide) {
        lucide.createIcons();
    }
}

document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('securityUpdateForm');
    const passwordInput = document.getElementById('newPasswordInput');
    const applyManualBtn = document.getElementById('applyManualBtn');
    
    const ruleLength = document.getElementById('ruleLength');
    const ruleLetters = document.getElementById('ruleLetters');
    const ruleNumbers = document.getElementById('ruleNumbers');
    const validatorStatusText = document.getElementById('validatorStatusText');
    
    // Add User Form Elements
    const createUserForm = document.getElementById('createUserForm');
    const createPasswordInput = document.getElementById('createUserPasswordInput');
    const createUserSubmitBtn = document.getElementById('createUserSubmitBtn');
    
    const createRuleLength = document.getElementById('createRuleLength');
    const createRuleLetters = document.getElementById('createRuleLetters');
    const createRuleNumbers = document.getElementById('createRuleNumbers');
    const createValidatorStatusText = document.getElementById('createValidatorStatusText');
    
    // Initial pagination load
    filterAccounts();
    
    // AJAX Submit Handler for Password Update
    if (form) {
        form.onsubmit = function(e) {
            e.preventDefault();
            
            const formData = new FormData(this);
            
            fetch(this.action, {
                method: 'POST',
                headers: { 'X-Requested-With': 'XMLHttpRequest' },
                body: formData
            })
            .then(r => r.json())
            .then(res => {
                if (res.success) {
                    showToast(res.msg || 'Password updated successfully!');
                    
                    // Clear only the password field after success
                    if (passwordInput) {
                        passwordInput.value = '';
                        passwordInput.dispatchEvent(new Event('input'));
                    }
                    
                    // Close the manual password section if it was shown
                    const manualSection = document.getElementById('manualPasswordSection');
                    if (manualSection && manualSection.style.display !== 'none') {
                        toggleManualPasswordFlow();
                    }

                    // Dynamically update the Registered Accounts table row
                    const row = document.querySelector(`.account-row[data-username="${res.username}"]`);
                    if (row) {
                        const cell = row.querySelector('.last-reset-cell');
                        if (cell) {
                            cell.innerHTML = `
                                <div style="display: flex; align-items: center; gap: 0.4rem; color: #0f766e;">
                                    <i data-lucide="calendar" style="width: 14px; height: 14px;"></i>
                                    <span>${res.updated_at}</span>
                                </div>
                            `;
                            if (window.lucide) {
                                lucide.createIcons();
                            }
                        }
                        
                        row.style.transition = 'background 0.3s ease';
                        row.style.background = '#f0fdf4';
                        setTimeout(() => {
                            row.style.background = '';
                        }, 2000);
                    }
                } else {
                    showToast(res.error || 'Failed to update password.', 'error');
                }
            })
            .catch(err => {
                showToast('Network error: ' + err, 'error');
            });
        };
    }
    

    
    if (passwordInput) {
        passwordInput.addEventListener('input', function() {
            const val = this.value;
            const isLengthValid = val.length >= 8;
            const isLettersValid = /[a-zA-Z]/.test(val);
            const isNumbersValid = /[0-9]/.test(val);
            
            updateRuleUI(ruleLength, isLengthValid);
            updateRuleUI(ruleLetters, isLettersValid);
            updateRuleUI(ruleNumbers, isNumbersValid);
            
            const allValid = isLengthValid && isLettersValid && isNumbersValid;
            if (allValid) {
                applyManualBtn.removeAttribute('disabled');
                applyManualBtn.style.opacity = '1';
                applyManualBtn.style.pointerEvents = 'auto';
                validatorStatusText.textContent = 'Strong password';
                validatorStatusText.className = 'validator-status-success';
            } else {
                applyManualBtn.setAttribute('disabled', 'true');
                applyManualBtn.style.opacity = '0.6';
                applyManualBtn.style.pointerEvents = 'none';
                if (val.length === 0) {
                    validatorStatusText.textContent = 'Enter a secure password';
                    validatorStatusText.className = 'validator-status-pending';
                } else {
                    validatorStatusText.textContent = 'Weak password';
                    validatorStatusText.className = 'validator-status-error';
                }
            }
        });
    }
    
    if (createPasswordInput) {
        createPasswordInput.addEventListener('input', function() {
            const val = this.value;
            const isLengthValid = val.length >= 8;
            const isLettersValid = /[a-zA-Z]/.test(val);
            const isNumbersValid = /[0-9]/.test(val);
            
            updateRuleUI(createRuleLength, isLengthValid);
            updateRuleUI(createRuleLetters, isLettersValid);
            updateRuleUI(createRuleNumbers, isNumbersValid);
            
            const allValid = isLengthValid && isLettersValid && isNumbersValid;
            if (allValid) {
                createUserSubmitBtn.removeAttribute('disabled');
                createUserSubmitBtn.style.opacity = '1';
                createUserSubmitBtn.style.pointerEvents = 'auto';
                createValidatorStatusText.textContent = 'Strong password';
                createValidatorStatusText.className = 'validator-status-success';
            } else {
                createUserSubmitBtn.setAttribute('disabled', 'true');
                createUserSubmitBtn.style.opacity = '0.6';
                createUserSubmitBtn.style.pointerEvents = 'none';
                if (val.length === 0) {
                    createValidatorStatusText.textContent = 'Enter a secure password';
                    createValidatorStatusText.className = 'validator-status-pending';
                } else {
                    createValidatorStatusText.textContent = 'Weak password';
                    createValidatorStatusText.className = 'validator-status-error';
                }
            }
        });
    }
});

function updateRuleUI(ruleEl, isValid) {
    if (!ruleEl) return;
    const emptyIcon = ruleEl.querySelector('.rule-icon-empty');
    const validIcon = ruleEl.querySelector('.rule-icon-valid');
    
    if (isValid) {
        ruleEl.classList.add('valid');
        if (emptyIcon) emptyIcon.style.display = 'none';
        if (validIcon) validIcon.style.display = 'inline-block';
    } else {
        ruleEl.classList.remove('valid');
        if (emptyIcon) emptyIcon.style.display = 'inline-block';
        if (validIcon) validIcon.style.display = 'none';
    }
}

function submitCreateUserForm(event) {
    if (event) event.preventDefault();
    
    const form = document.getElementById('createUserForm');
    if (!form) return;
    
    const passMode = document.getElementById('createUserPassMode').value;
    if (passMode === 'manual') {
        const passInput = document.getElementById('createUserPasswordInput');
        const val = passInput ? passInput.value : '';
        const isLengthValid = val.length >= 8;
        const isLettersValid = /[a-zA-Z]/.test(val);
        const isNumbersValid = /[0-9]/.test(val);
        if (!(isLengthValid && isLettersValid && isNumbersValid)) {
            showToast('Please enter a secure password that meets all requirements.', 'error');
            return;
        }
    }
    
    const formData = new FormData(form);
    const submitBtn = document.getElementById('createUserSubmitBtn');
    if (submitBtn) {
        submitBtn.setAttribute('disabled', 'true');
        submitBtn.style.opacity = '0.6';
    }
    
    fetch(form.action, {
        method: 'POST',
        headers: { 'X-Requested-With': 'XMLHttpRequest' },
        body: formData
    })
    .then(r => {
        if (!r.ok) {
            return r.json().then(errData => {
                throw new Error(errData.error || 'Server error occurred.');
            }).catch(e => {
                throw new Error(e.message || 'Server error occurred.');
            });
        }
        return r.json();
    })
    .then(res => {
        if (res.success) {
            showToast(res.msg || 'User account created successfully!', 'success');
            
            // Reset form
            form.reset();
            toggleAddUserRoleFields('admin');
            setAddUserPasswordMode('auto');
            
            const createPasswordInput = document.getElementById('createUserPasswordInput');
            if (createPasswordInput) {
                createPasswordInput.value = '';
                createPasswordInput.dispatchEvent(new Event('input'));
            }
            
            // Dynamically insert into list
            const tbody = document.querySelector('#accountsTable tbody');
            if (tbody) {
                const emptyRow = tbody.querySelector('tr td[colspan]');
                if (emptyRow) {
                    tbody.innerHTML = '';
                }
                
                const tr = document.createElement('tr');
                tr.className = 'account-row';
                tr.setAttribute('data-username', res.username);
                tr.setAttribute('data-role', res.role);
                tr.style.borderBottom = '1px solid var(--border)';
                tr.style.transition = 'background 0.3s ease';
                tr.style.background = '#ecfdf5';
                
                let roleBadge = '';
                if (res.role === 'admin') {
                    roleBadge = '<span class="badge-role badge-admin">Administrator</span>';
                } else if (res.role === 'teacher') {
                    roleBadge = '<span class="badge-role badge-teacher">Teacher</span>';
                } else {
                    roleBadge = '<span class="badge-role badge-student">Student</span>';
                }
                
                let actionButtonsHtml = `
                    <button type="button" onclick="selectUserForReset('${res.username}')" class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.4rem 0.8rem; border-radius: 8px; font-weight: 600;">
                        Select Account
                    </button>
                `;
                if (isSuperAdmin) {
                    actionButtonsHtml += `
                        <button type="button" onclick="openEditUserModal('${res.user_id}', '${res.username}')" class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.4rem 0.8rem; border-radius: 8px; font-weight: 600; color: var(--primary); border-color: var(--primary);" title="Edit Username">
                            <i data-lucide="edit-2" style="width: 14px; height: 14px;"></i>
                        </button>
                    `;
                    if (res.user_id != currentUserId) {
                        actionButtonsHtml += `
                            <button type="button" onclick="openDeleteUserModal('${res.user_id}', '${res.username}')" class="btn btn-secondary" style="font-size: 0.75rem; padding: 0.4rem 0.8rem; border-radius: 8px; font-weight: 600; color: #ef4444; border-color: #fca5a5;" title="Delete Account">
                                <i data-lucide="trash-2" style="width: 14px; height: 14px;"></i>
                            </button>
                        `;
                    }
                }

                tr.innerHTML = `
                    <td style="padding: 1rem; font-weight: 600; color: var(--text-main); font-size: 0.9375rem;">
                        <div style="display: flex; align-items: center; gap: 0.5rem;">
                            <i data-lucide="user" style="width: 16px; height: 16px; color: var(--text-muted);"></i>
                            <span>${res.username}</span>
                        </div>
                    </td>
                    <td style="padding: 1rem;">
                        ${roleBadge}
                    </td>
                    <td style="padding: 1rem; font-size: 0.875rem;" class="last-reset-cell">
                        <span style="color: var(--text-muted); font-style: italic;">Never updated</span>
                    </td>
                    <td style="padding: 1rem; text-align: right; white-space: nowrap;">
                        <div style="display: inline-flex; gap: 0.5rem; justify-content: flex-end; align-items: center;">
                            ${actionButtonsHtml}
                        </div>
                    </td>
                `;
                
                tbody.insertBefore(tr, tbody.firstChild);
                
                // Force refresh datalist and datatable select options
                const usersList = document.getElementById('usersList');
                if (usersList) {
                    const opt = document.createElement('option');
                    opt.value = res.username;
                    opt.textContent = `${res.username} (${res.role})`;
                    usersList.appendChild(opt);
                }
                const userSelect = document.getElementById('userSelect');
                if (userSelect) {
                    const opt = document.createElement('option');
                    opt.value = res.username;
                    opt.textContent = `${res.username} (${res.role})`;
                    userSelect.appendChild(opt);
                }
                
                currentPage = 1;
                filterAccounts();
                
                if (window.lucide) {
                    lucide.createIcons();
                }
                
                setTimeout(() => {
                    tr.style.background = '';
                }, 2000);
            }
        } else {
            showToast(res.error || 'Failed to create user account.', 'error');
        }
    })
    .catch(err => {
        showToast(err.message || 'Network error occurred.', 'error');
    })
    .finally(() => {
        if (submitBtn) {
            submitBtn.removeAttribute('disabled');
            submitBtn.style.opacity = '1';
        }
    });
// Edit User Modal Handlers
function openEditUserModal(userId, username) {
    const modal = document.getElementById('editUserModal');
    const idInput = document.getElementById('editUserId');
    const oldInput = document.getElementById('editOldUsername');
    const newInput = document.getElementById('editNewUsername');
    const comparison = document.getElementById('editComparisonDisplay');
    
    if (modal && idInput && oldInput && newInput) {
        idInput.value = userId;
        oldInput.value = username;
        newInput.value = '';
        if (comparison) comparison.style.display = 'none';
        modal.style.display = 'flex';
        
        setTimeout(() => newInput.focus(), 150);
        
        if (window.lucide) {
            lucide.createIcons();
        }
    }
}

function closeEditUserModal() {
    const modal = document.getElementById('editUserModal');
    if (modal) {
        modal.style.display = 'none';
    }
}

function updateEditComparison() {
    const oldVal = document.getElementById('editOldUsername').value;
    const newVal = document.getElementById('editNewUsername').value.trim();
    const comparison = document.getElementById('editComparisonDisplay');
    const previewOld = document.getElementById('previewOld');
    const previewNew = document.getElementById('previewNew');
    
    if (newVal.length > 0 && newVal !== oldVal) {
        previewOld.textContent = oldVal;
        previewNew.textContent = newVal;
        comparison.style.display = 'block';
    } else {
        comparison.style.display = 'none';
    }
}

function submitEditUserForm(event) {
    if (event) event.preventDefault();
    
    const userId = document.getElementById('editUserId').value;
    const oldUsername = document.getElementById('editOldUsername').value;
    const newUsername = document.getElementById('editNewUsername').value.trim();
    
    if (!newUsername) {
        showToast('New username cannot be empty.', 'error');
        return;
    }
    
    if (newUsername === oldUsername) {
        showToast('New username must be different from the old username.', 'error');
        return;
    }
    
    const saveBtn = document.getElementById('saveEditBtn');
    if (saveBtn) {
        saveBtn.setAttribute('disabled', 'true');
        saveBtn.style.opacity = '0.6';
    }
    
    const formData = new FormData();
    formData.append('username', newUsername);
    
    fetch(`/admin/users/edit/${userId}`, {
        method: 'POST',
        headers: { 'X-Requested-With': 'XMLHttpRequest' },
        body: formData
    })
    .then(r => {
        if (!r.ok) {
            return r.json().then(errData => {
                throw new Error(errData.error || 'Server error occurred.');
            }).catch(e => {
                throw new Error(e.message || 'Server error occurred.');
            });
        }
        return r.json();
    })
    .then(res => {
        if (res.success) {
            showToast(res.msg || 'Username updated successfully!', 'success');
            closeEditUserModal();
            
            // Dynamically update username in the table rows
            const rows = document.querySelectorAll(`.account-row[data-username="${oldUsername}"]`);
            rows.forEach(row => {
                row.setAttribute('data-username', res.username);
                const span = row.querySelector('td div span');
                if (span) {
                    span.textContent = res.username;
                }
                
                // Update actions buttons with new username parameter
                const selectBtn = row.querySelector('button[onclick^="selectUserForReset"]');
                if (selectBtn) {
                    selectBtn.setAttribute('onclick', `selectUserForReset('${res.username}')`);
                }
                const editBtn = row.querySelector('button[onclick^="openEditUserModal"]');
                if (editBtn) {
                    editBtn.setAttribute('onclick', `openEditUserModal('${res.user_id}', '${res.username}')`);
                }
                const deleteBtn = row.querySelector('button[onclick^="openDeleteUserModal"]');
                if (deleteBtn) {
                    deleteBtn.setAttribute('onclick', `openDeleteUserModal('${res.user_id}', '${res.username}')`);
                }
                
                row.style.transition = 'background 0.3s ease';
                row.style.background = '#eff6ff';
                setTimeout(() => row.style.background = '', 2000);
            });
            
            // Also update any datalists or browse selects
            const usersList = document.getElementById('usersList');
            if (usersList) {
                const opt = usersList.querySelector(`option[value="${oldUsername}"]`);
                if (opt) {
                    opt.value = res.username;
                    opt.textContent = `${res.username} (${opt.textContent.split('(')[1] || ''}`;
                }
            }
            const userSelect = document.getElementById('userSelect');
            if (userSelect) {
                const opt = userSelect.querySelector(`option[value="${oldUsername}"]`);
                if (opt) {
                    opt.value = res.username;
                    opt.textContent = `${res.username} (${opt.textContent.split('(')[1] || ''}`;
                }
            }
            
            filterAccounts();
        } else {
            showToast(res.error || 'Failed to update username.', 'error');
        }
    })
    .catch(err => {
        showToast(err.message || 'Network error occurred.', 'error');
    })
    .finally(() => {
        if (saveBtn) {
            saveBtn.removeAttribute('disabled');
            saveBtn.style.opacity = '1';
        }
    });
}

// Delete User Modal Handlers
function openDeleteUserModal(userId, username) {
    const modal = document.getElementById('deleteUserModal');
    const idInput = document.getElementById('deleteUserId');
    const targetText = document.getElementById('deleteTargetUsername');
    
    if (modal && idInput && targetText) {
        idInput.value = userId;
        targetText.textContent = username;
        modal.style.display = 'flex';
        
        if (window.lucide) {
            lucide.createIcons();
        }
    }
}

function closeDeleteUserModal() {
    const modal = document.getElementById('deleteUserModal');
    if (modal) {
        modal.style.display = 'none';
    }
}

function submitDeleteUserForm(event) {
    if (event) event.preventDefault();
    
    const userId = document.getElementById('deleteUserId').value;
    const username = document.getElementById('deleteTargetUsername').textContent;
    
    const confirmBtn = document.getElementById('confirmDeleteBtn');
    if (confirmBtn) {
        confirmBtn.setAttribute('disabled', 'true');
        confirmBtn.style.opacity = '0.6';
    }
    
    fetch(`/admin/users/delete/${userId}`, {
        method: 'POST',
        headers: { 'X-Requested-With': 'XMLHttpRequest' }
    })
    .then(r => {
        if (!r.ok) {
            return r.json().then(errData => {
                throw new Error(errData.error || 'Server error occurred.');
            }).catch(e => {
                throw new Error(e.message || 'Server error occurred.');
            });
        }
        return r.json();
    })
    .then(res => {
        if (res.success) {
            showToast(res.msg || 'User account successfully deleted.', 'success');
            closeDeleteUserModal();
            
            // Animate and remove the row from the directory table
            const row = document.querySelector(`.account-row[data-username="${username}"]`);
            if (row) {
                row.style.transition = 'all 0.4s ease';
                row.style.background = '#fee2e2';
                row.style.opacity = '0';
                row.style.transform = 'translateX(-20px)';
                setTimeout(() => {
                    row.remove();
                    currentPage = 1;
                    filterAccounts();
                }, 400);
            }
            
            // Remove from datalists or browse selects
            const usersList = document.getElementById('usersList');
            if (usersList) {
                const opt = usersList.querySelector(`option[value="${username}"]`);
                if (opt) opt.remove();
            }
            const userSelect = document.getElementById('userSelect');
            if (userSelect) {
                const opt = userSelect.querySelector(`option[value="${username}"]`);
                if (opt) opt.remove();
            }
        } else {
            showToast(res.error || 'Failed to delete user account.', 'error');
        }
    })
    .catch(err => {
        showToast(err.message || 'Network error occurred.', 'error');
    })
    .finally(() => {
        if (confirmBtn) {
            confirmBtn.removeAttribute('disabled');
            confirmBtn.style.opacity = '1';
        }
    });
}
