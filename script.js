// Global State
let participants = [];
let winners = [];
let prizes = [];
let absentPeople = [];
let excelColumns = [];
let displayColumns = [];
let isSpinning = false;
let selectedPrizeId = null;
let currentWinners = [];
let rollInterval = null;
let gridSize = { rows: 0, cols: 0, total: 0 };

// ADMIN - Enhanced
let adminSelectedPerson = null;
let adminSelectedPrize = null;
let isAdminPanelOpen = false;
let adminPresetList = []; // Danh sách setup trước: [{person, prize}, ...]

// Panel states
let isControlPanelOpen = false;
let isPrizePanelOpen = false;

// Music state
let currentMusicState = {
    prizeId: null,
    isPlaying: false
};

// ============= NOTIFICATION SYSTEM =============
function showNotification(message, type = 'info', duration = 4000) {
    const container = document.getElementById('notificationContainer');
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    
    // Map icon based on type
    const icons = {
        'success': '✅',
        'error': '❌',
        'warning': '⚠️',
        'info': 'ℹ️'
    };
    
    notification.innerHTML = `
        <span class="notification-icon">${icons[type] || icons['info']}</span>
        <span class="notification-message">${message}</span>
        <button class="notification-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    
    container.appendChild(notification);
    
    // Auto remove after duration
    if (duration > 0) {
        setTimeout(() => {
            notification.classList.add('exit');
            setTimeout(() => notification.remove(), 300);
        }, duration);
    }
}

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);
    document.getElementById('prizeSelect').addEventListener('change', function(e) {
        selectedPrizeId = e.target.value ? Number(e.target.value) : null;
        updateSpinButton();
        playPrizeMusic(selectedPrizeId);
    });
    
    // Music file input listener
    document.getElementById('prizeMusic').addEventListener('change', function(e) {
        const fileName = e.target.files[0]?.name || 'Chọn nhạc';
        document.getElementById('musicFileName').textContent = fileName;
    });
    
    calculateGridSize();
    createGrid();
    
    updateStats();
    updateSpinButton();
    
    // Admin Panel - Ctrl+Shift+Z
    document.addEventListener('keydown', function(e) {
        if (e.ctrlKey && e.shiftKey && (e.key === 'Z' || e.key === 'z')) {
            e.preventDefault();
            toggleAdminPanel();
        }
        // ESC to close panels
        if (e.key === 'Escape') {
            if (isControlPanelOpen) toggleControlPanel();
            if (isPrizePanelOpen) togglePrizePanel();
            if (isAdminPanelOpen) closeAdminPanel();
        }
    });
    
    // Resize listener
    window.addEventListener('resize', function() {
        calculateGridSize();
        createGrid();
        if (participants.length > 0) {
            initializeRollGrid();
        }
    });
});

// ============= MUSIC CONTROL =============

function playPrizeMusic(prizeId) {
    if (!prizeId) {
        stopPrizeMusic();
        return;
    }

    const prize = prizes.find(p => p.id === prizeId);
    if (!prize || !prize.music) {
        stopPrizeMusic();
        return;
    }

    const audio = document.getElementById('prizeAudio');
    
    // If same prize is already playing, do nothing
    if (currentMusicState.prizeId === prizeId && currentMusicState.isPlaying) {
        return;
    }

    try {
        // Set audio source based on music type
        if (prize.musicType === 'file') {
            // For file music, create blob URL
            const blob = new Blob([prize.music], { type: 'audio/mpeg' });
            audio.src = URL.createObjectURL(blob);
        } else if (prize.musicType === 'url') {
            // For URL music
            audio.src = prize.music;
        }

        audio.loop = true;
        audio.play().catch(err => {
            console.log('Music playback failed:', err);
        });

        currentMusicState.prizeId = prizeId;
        currentMusicState.isPlaying = true;
    } catch (error) {
        console.error('Error playing music:', error);
        showNotification('Không thể phát nhạc!', 'error');
    }
}

function stopPrizeMusic() {
    const audio = document.getElementById('prizeAudio');
    audio.pause();
    audio.currentTime = 0;
    currentMusicState.prizeId = null;
    currentMusicState.isPlaying = false;
}

// ============= PANEL CONTROLS =============

function toggleControlPanel() {
    isControlPanelOpen = !isControlPanelOpen;
    const panel = document.getElementById('controlPanel');
    const btn = document.getElementById('controlBtn');
    
    if (isControlPanelOpen) {
        panel.classList.add('active');
        btn.classList.add('active');
        if (isPrizePanelOpen) togglePrizePanel();
    } else {
        panel.classList.remove('active');
        btn.classList.remove('active');
    }
}

function togglePrizePanel() {
    isPrizePanelOpen = !isPrizePanelOpen;
    const panel = document.getElementById('prizePanel');
    const btn = document.getElementById('prizeBtn');
    
    if (isPrizePanelOpen) {
        panel.classList.add('active');
        btn.classList.add('active');
        if (isControlPanelOpen) toggleControlPanel();
    } else {
        panel.classList.remove('active');
        btn.classList.remove('active');
    }
}

// ============= GRID CALCULATION =============

function calculateGridSize() {
    const width = window.innerWidth;
    let cols, rows;
    
    if (width >= 1920) {
        cols = 9; rows = 5;
    } else if (width >= 1440) {
        cols = 7; rows = 5;
    } else if (width >= 1024) {
        cols = 5; rows = 4;
    } else if (width >= 768) {
        cols = 4; rows = 3;
    } else {
        cols = 3; rows = 3;
    }
    
    if (cols % 2 === 0) cols--;
    if (rows % 2 === 0) rows--;
    
    gridSize = {
        rows: rows,
        cols: cols,
        total: rows * cols,
        center: Math.floor(rows * cols / 2)
    };
}

function createGrid() {
    const gridEl = document.getElementById('rollGrid');
    if (!gridEl) return;
    
    gridEl.innerHTML = '';
    gridEl.style.gridTemplateColumns = `repeat(${gridSize.cols}, 1fr)`;
    gridEl.style.gridTemplateRows = `repeat(${gridSize.rows}, 1fr)`;
    
    for (let i = 0; i < gridSize.total; i++) {
        if (i === gridSize.center) {
            const centerCell = document.createElement('div');
            centerCell.className = 'roll-cell-new center-cell-new';
            centerCell.dataset.cell = i;
            gridEl.appendChild(centerCell);
        } else {
            const cell = document.createElement('div');
            cell.className = 'roll-cell-new';
            cell.dataset.cell = i;
            cell.innerHTML = `
                <div class="roll-content-new">
                    <div class="employee-code-new">---</div>
                    <div class="employee-name-new">---</div>
                    <div class="employee-dept-new">---</div>
                </div>
            `;
            gridEl.appendChild(cell);
        }
    }
}

// ============= ADMIN FUNCTIONS - ENHANCED =============

function toggleAdminPanel() {
    isAdminPanelOpen = !isAdminPanelOpen;
    const panel = document.getElementById('adminPanel');
    
    if (isAdminPanelOpen) {
        panel.classList.add('active');
        updateAdminPrizeSelect();
        updateAdminParticipantList();
        updateAdminPresetList();
    } else {
        panel.classList.remove('active');
    }
}

function closeAdminPanel() {
    isAdminPanelOpen = false;
    document.getElementById('adminPanel').classList.remove('active');
}

function updateAdminPrizeSelect() {
    const select = document.getElementById('adminPrizeSelect');
    if (!select) return;
    
    let html = '<option value="">-- Chọn giải --</option>';
    prizes.forEach(prize => {
        html += `<option value="${prize.id}">${prize.name} (Còn ${prize.remaining})</option>`;
    });
    select.innerHTML = html;
}

function updateAdminParticipantList(searchTerm = '') {
    const listDiv = document.getElementById('adminParticipantList');
    
    if (participants.length === 0) {
        listDiv.innerHTML = '<p class="admin-empty">Chưa tải dữ liệu</p>';
        return;
    }
    
    let available = getAvailableParticipants();
    
    if (searchTerm) {
        available = available.filter(person => {
            const searchableText = excelColumns.map(col => 
                String(person[col] || '').toLowerCase()
            ).join(' ');
            return searchableText.includes(searchTerm);
        });
    }
    
    if (available.length === 0) {
        listDiv.innerHTML = '<p class="admin-empty">Không tìm thấy</p>';
        return;
    }
    
    let html = '';
    available.forEach(person => {
        const isSelected = adminSelectedPerson && adminSelectedPerson.id === person.id;
        const selectedClass = isSelected ? 'selected' : '';
        
        html += `
            <div class="admin-person-item ${selectedClass}" onclick="selectAdminPerson(${person.id})">
                <div class="admin-person-info">
                    <div class="admin-person-id">ID: ${person.id}</div>
                    <div class="admin-person-details">
        `;
        
        // Hiển thị TẤT CẢ các cột (không giới hạn 3)
        displayColumns.forEach(col => {
            if (person[col]) {
                html += `<div><strong>${col}:</strong> ${person[col]}</div>`;
            }
        });
        
        html += `
                    </div>
                </div>
                ${isSelected ? '<div class="admin-check">✓</div>' : ''}
            </div>
        `;
    });
    
    listDiv.innerHTML = html;
}

function filterAdminList() {
    const searchTerm = document.getElementById('adminSearch').value.toLowerCase();
    updateAdminParticipantList(searchTerm);
}

function selectAdminPerson(personId) {
    const person = participants.find(p => p.id === personId);
    if (!person) return;
    
    adminSelectedPerson = person;
    updateAdminParticipantList(document.getElementById('adminSearch').value.toLowerCase());
    updateAdminCurrentSelection();
}

function selectAdminPrize() {
    const selectEl = document.getElementById('adminPrizeSelect');
    const prizeId = selectEl.value ? Number(selectEl.value) : null;
    
    if (prizeId) {
        adminSelectedPrize = prizes.find(p => p.id === prizeId);
    } else {
        adminSelectedPrize = null;
    }
    
    updateAdminCurrentSelection();
}

function updateAdminCurrentSelection() {
    const displayDiv = document.getElementById('adminCurrentSelection');
    
    if (!adminSelectedPerson && !adminSelectedPrize) {
        displayDiv.innerHTML = '<span class="admin-no-selection">Chưa chọn người và giải</span>';
        return;
    }
    
    let html = '<div class="admin-current-info">';
    
    if (adminSelectedPerson) {
        html += '<div class="admin-current-person">';
        html += '<div class="admin-current-label">👤 Người:</div>';
        displayColumns.forEach(col => {
            if (adminSelectedPerson[col]) {
                html += `<div><strong>${col}:</strong> ${adminSelectedPerson[col]}</div>`;
            }
        });
        html += '</div>';
    }
    
    if (adminSelectedPrize) {
        html += '<div class="admin-current-prize">';
        html += `<div class="admin-current-label">🏆 Giải: ${adminSelectedPrize.name}</div>`;
        html += '</div>';
    }
    
    html += '</div>';
    displayDiv.innerHTML = html;
}

function addToPreset() {
    if (!adminSelectedPerson || !adminSelectedPrize) {
        showNotification('Vui lòng chọn cả người và giải!', 'warning');
        return;
    }
    
    // Kiểm tra đã có trong preset chưa
    const exists = adminPresetList.some(p => 
        p.person.id === adminSelectedPerson.id && p.prize.id === adminSelectedPrize.id
    );
    
    if (exists) {
        showNotification('Đã có trong danh sách setup!', 'warning');
        return;
    }
    
    adminPresetList.push({
        person: { ...adminSelectedPerson },
        prize: { ...adminSelectedPrize }
    });
    
    // Clear selection
    adminSelectedPerson = null;
    adminSelectedPrize = null;
    document.getElementById('adminPrizeSelect').value = '';
    
    updateAdminCurrentSelection();
    updateAdminParticipantList(document.getElementById('adminSearch').value.toLowerCase());
    updateAdminPresetList();
    
    // Notification
    showNotification('Đã thêm vào danh sách!', 'success');
}

function removeFromPreset(index) {
    adminPresetList.splice(index, 1);
    updateAdminPresetList();
    showNotification('Đã xóa khỏi danh sách!', 'success');
}

function clearAllPresets() {
    if (confirm('Xóa toàn bộ danh sách setup?')) {
        adminPresetList = [];
        updateAdminPresetList();
    }
}

function updateAdminPresetList() {
    const listDiv = document.getElementById('adminPresetList');
    const countEl = document.getElementById('presetCount');
    
    if (countEl) {
        countEl.textContent = adminPresetList.length;
    }
    
    if (adminPresetList.length === 0) {
        listDiv.innerHTML = '<p class="admin-empty">Chưa có setup nào</p>';
        return;
    }
    
    let html = '';
    adminPresetList.forEach((preset, index) => {
        html += `
            <div class="admin-preset-item">
                <div class="admin-preset-content">
                    <div class="admin-preset-person">
                        <strong>👤</strong>
        `;
        
        displayColumns.forEach(col => {
            if (preset.person[col]) {
                html += `<span>${preset.person[col]}</span>`;
            }
        });
        
        html += `
                    </div>
                    <div class="admin-preset-arrow">→</div>
                    <div class="admin-preset-prize">
                        <strong>🏆 ${preset.prize.name}</strong>
                    </div>
                </div>
                <button onclick="removeFromPreset(${index})" class="admin-preset-delete">
                    <svg class="icon-sm" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" 
                            d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                    </svg>
                </button>
            </div>
        `;
    });
    
    listDiv.innerHTML = html;
}

function clearAdminSelection() {
    adminSelectedPerson = null;
    adminSelectedPrize = null;
    document.getElementById('adminPrizeSelect').value = '';
    updateAdminCurrentSelection();
    updateAdminParticipantList(document.getElementById('adminSearch').value.toLowerCase());
}

// ============= FILE UPLOAD =============

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            if (jsonData.length === 0) {
                showNotification('File Excel không có dữ liệu!', 'error');
                return;
            }

            excelColumns = Object.keys(jsonData[0]);
            displayColumns = excelColumns.filter(col => 
                col.toLowerCase() !== 'stt'
            ).slice(0, 3);
            
            participants = jsonData.map((row, idx) => ({
                id: idx + 1,
                ...row
            }));

            showNotification(`Đã tải ${participants.length} người!`, 'success');
            updateStats();
            updateSpinButton();
            updateAbsentCount();
            initializeRollGrid();
            
            if (isAdminPanelOpen) {
                updateAdminPrizeSelect();
                updateAdminParticipantList();
            }
        } catch (error) {
            showNotification('Lỗi khi đọc file Excel!', 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ============= GRID FUNCTIONS =============

function initializeRollGrid() {
    const available = getAvailableParticipants();
    if (available.length === 0) return;

    for (let i = 0; i < gridSize.total; i++) {
        if (i === gridSize.center) continue;
        
        const cell = document.querySelector(`[data-cell="${i}"]`);
        if (!cell) continue;
        
        const randomPerson = available[Math.floor(Math.random() * available.length)];
        updateCellContent(cell, randomPerson);
    }
}

function updateCellContent(cell, person) {
    if (!person || !cell) return;
    
    const codeEl = cell.querySelector('.employee-code-new');
    const nameEl = cell.querySelector('.employee-name-new');
    const deptEl = cell.querySelector('.employee-dept-new');
    
    if (displayColumns.length >= 3) {
        if (codeEl) codeEl.textContent = person[displayColumns[0]] || '---';
        if (nameEl) nameEl.textContent = person[displayColumns[1]] || '---';
        if (deptEl) deptEl.textContent = person[displayColumns[2]] || '---';
    }
}

function startRolling() {
    const available = getAvailableParticipants();
    if (available.length === 0) return;

    let speed = 30;
    const speedIncrement = 2;
    const maxSpeed = 200;
    let cellsArray = [];
    
    // Lấy danh sách các ô (không tính center)
    for (let i = 0; i < gridSize.total; i++) {
        if (i !== gridSize.center) cellsArray.push(i);
    }

    function rollFrame() {
        // Chỉ cập nhật người ngẫu nhiên - giữ phần rung, không thay màu
        for (let i = 0; i < cellsArray.length; i++) {
            const cellIndex = cellsArray[i];
            const cell = document.querySelector(`[data-cell="${cellIndex}"]`);
            if (!cell) continue;
            
            const randomPerson = available[Math.floor(Math.random() * available.length)];
            updateCellContent(cell, randomPerson);
            
            cell.classList.add('rolling');
        }

        speed += speedIncrement;
        
        if (speed <= maxSpeed) {
            rollInterval = setTimeout(rollFrame, speed);
        } else {
            // Kết thúc roll - clear rolling state
            for (let i = 0; i < cellsArray.length; i++) {
                const cellIndex = cellsArray[i];
                const cell = document.querySelector(`[data-cell="${cellIndex}"]`);
                if (cell) {
                    cell.classList.remove('rolling');
                }
            }
        }
    }

    rollFrame();
}

function stopRolling() {
    if (rollInterval) {
        clearTimeout(rollInterval);
        rollInterval = null;
    }
    
    // Clear all rolling states
    for (let i = 0; i < gridSize.total; i++) {
        if (i === gridSize.center) continue;
        const cell = document.querySelector(`[data-cell="${i}"]`);
        if (cell) {
            cell.classList.remove('rolling');
        }
    }
}

// ============= PRIZE FUNCTIONS =============

function handleAddPrize() {
    const name = document.getElementById('prizeName').value.trim();
    const quantity = parseInt(document.getElementById('prizeQuantity').value);
    const winnersPerPrize = parseInt(document.getElementById('winnersPerPrize').value);
    const musicFile = document.getElementById('prizeMusic').files[0];
    const musicUrl = document.getElementById('prizeMusic Url').value.trim();

    if (!name || quantity < 1 || winnersPerPrize < 1) {
        showNotification('Vui lòng nhập đầy đủ!', 'warning');
        return;
    }

    // Xử lý file nhạc
    if (musicFile) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const newPrize = {
                id: Date.now(),
                name: name,
                quantity: quantity,
                winnersPerPrize: winnersPerPrize,
                remaining: quantity,
                musicType: 'file',
                music: e.target.result,
                musicName: musicFile.name
            };

            prizes.push(newPrize);
            clearPrizeForm();
            updatePrizeList();
            updatePrizeSelect();
            
            if (isAdminPanelOpen) {
                updateAdminPrizeSelect();
            }
            
            showNotification(`Thêm giải "${name}" với nhạc!`, 'success');
        };
        reader.readAsArrayBuffer(musicFile);
    } else if (musicUrl) {
        // Validate URL
        try {
            new URL(musicUrl);
            const newPrize = {
                id: Date.now(),
                name: name,
                quantity: quantity,
                winnersPerPrize: winnersPerPrize,
                remaining: quantity,
                musicType: 'url',
                music: musicUrl,
                musicName: musicUrl
            };

            prizes.push(newPrize);
            clearPrizeForm();
            updatePrizeList();
            updatePrizeSelect();
            
            if (isAdminPanelOpen) {
                updateAdminPrizeSelect();
            }
            
            showNotification(`Thêm giải "${name}" với nhạc từ URL!`, 'success');
        } catch (e) {
            showNotification('URL nhạc không hợp lệ!', 'error');
        }
    } else {
        // Không có nhạc - vẫn thêm giải được
        const newPrize = {
            id: Date.now(),
            name: name,
            quantity: quantity,
            winnersPerPrize: winnersPerPrize,
            remaining: quantity,
            musicType: null,
            music: null,
            musicName: null
        };

        prizes.push(newPrize);
        clearPrizeForm();
        updatePrizeList();
        updatePrizeSelect();
        
        if (isAdminPanelOpen) {
            updateAdminPrizeSelect();
        }
        
        showNotification(`Thêm giải "${name}"!`, 'success');
    }
}

function clearPrizeForm() {
    document.getElementById('prizeName').value = '';
    document.getElementById('prizeQuantity').value = 1;
    document.getElementById('winnersPerPrize').value = 1;
    document.getElementById('prizeMusic').value = '';
    document.getElementById('prizeMusic Url').value = '';
    document.getElementById('musicFileName').textContent = 'Chọn nhạc';
}

function switchMusicTab(tab) {
    // Hide all tabs
    document.getElementById('musicFileTab').classList.remove('active');
    document.getElementById('musicUrlTab').classList.remove('active');
    
    // Remove active class from all buttons
    document.querySelectorAll('.music-tab-btn').forEach(btn => btn.classList.remove('active'));
    
    // Show selected tab and mark button as active
    if (tab === 'file') {
        document.getElementById('musicFileTab').classList.add('active');
        document.querySelectorAll('.music-tab-btn')[0].classList.add('active');
    } else if (tab === 'url') {
        document.getElementById('musicUrlTab').classList.add('active');
        document.querySelectorAll('.music-tab-btn')[1].classList.add('active');
    }
}

function handleDeletePrize(id) {
    if (confirm('Xóa giải này?')) {
        prizes = prizes.filter(p => p.id !== id);
        if (selectedPrizeId === id) {
            selectedPrizeId = null;
            document.getElementById('prizeSelect').value = '';
        }
        updatePrizeList();
        updatePrizeSelect();
        updateSpinButton();
        
        if (isAdminPanelOpen) {
            updateAdminPrizeSelect();
        }
    }
}

function getAvailableParticipants() {
    return participants.filter(p => 
        !winners.some(w => w.participantId === p.id) &&
        !absentPeople.some(a => a.id === p.id)
    );
}

// ============= SPIN FUNCTION - ENHANCED WITH PRESET =============

async function handleSpin() {
    if (isSpinning) return;

    const available = getAvailableParticipants();
    if (available.length === 0) {
        showNotification('Không còn người!', 'error');
        return;
    }

    const selectedPrize = prizes.find(p => p.id === selectedPrizeId);
    if (!selectedPrize || selectedPrize.remaining === 0) {
        showNotification('Chọn giải thưởng!', 'warning');
        return;
    }

    const winnersNeeded = Math.min(selectedPrize.winnersPerPrize, available.length);
    isSpinning = true;

    const spinBtn = document.getElementById('spinBtn');
    const spinBtnText = document.getElementById('spinBtnText');
    const centerSpinBtn = document.getElementById('centerSpinBtn');
    
    spinBtn.disabled = true;
    spinBtn.classList.add('spinning');
    spinBtnText.textContent = 'ĐANG QUAY...';
    
    if (centerSpinBtn) {
        centerSpinBtn.disabled = true;
        centerSpinBtn.classList.add('spinning');
    }

    startRolling();
    await new Promise(resolve => setTimeout(resolve, 9000));
    stopRolling();

    // ADMIN LOGIC - Kiểm tra preset list
    const selectedWinners = [];
    const tempAvailable = [...available];
    
    // Tìm preset phù hợp với giải đang quay
    const matchingPresets = adminPresetList.filter(preset => 
        preset.prize.id === selectedPrize.id &&
        tempAvailable.find(p => p.id === preset.person.id)
    );
    
    // Thêm người từ preset trước
    for (const preset of matchingPresets) {
        if (selectedWinners.length >= winnersNeeded) break;
        
        const winner = tempAvailable.find(p => p.id === preset.person.id);
        if (!winner) continue;
        
        const winnerData = {
            participantId: winner.id,
            prize: selectedPrize.name,
            timestamp: new Date().toLocaleString('vi-VN'),
        };
        excelColumns.forEach(col => {
            winnerData[col] = winner[col];
        });
        selectedWinners.push(winnerData);
        
        // Xóa khỏi available
        const removeIndex = tempAvailable.findIndex(p => p.id === winner.id);
        if (removeIndex !== -1) {
            tempAvailable.splice(removeIndex, 1);
        }
        
        // Xóa khỏi preset list
        const presetIndex = adminPresetList.findIndex(p => 
            p.person.id === preset.person.id && p.prize.id === preset.prize.id
        );
        if (presetIndex !== -1) {
            adminPresetList.splice(presetIndex, 1);
        }
    }

    // Chọn người còn lại ngẫu nhiên
    for (let i = selectedWinners.length; i < winnersNeeded; i++) {
        const randomIndex = Math.floor(Math.random() * tempAvailable.length);
        const winner = tempAvailable[randomIndex];
        tempAvailable.splice(randomIndex, 1);

        const winnerData = {
            participantId: winner.id,
            prize: selectedPrize.name,
            timestamp: new Date().toLocaleString('vi-VN'),
        };
        excelColumns.forEach(col => {
            winnerData[col] = winner[col];
        });
        selectedWinners.push(winnerData);
    }

    winners = [...winners, ...selectedWinners];
    currentWinners = selectedWinners;

    const prizeIndex = prizes.findIndex(p => p.id === selectedPrize.id);
    prizes[prizeIndex].remaining -= 1;

    await showWinnersInGrid(selectedWinners);

    displayWinners(selectedWinners);
    updateStats();
    updatePrizeList();
    updatePrizeSelect();
    updateAbsentCount();

    // Stop music when spin ends
    stopPrizeMusic();

    isSpinning = false;
    spinBtn.classList.remove('spinning');
    spinBtnText.textContent = 'QUAY NGAY!';
    
    if (centerSpinBtn) {
        centerSpinBtn.classList.remove('spinning');
    }
    
    updateSpinButton();
    
    if (isAdminPanelOpen) {
        updateAdminPrizeSelect();
        updateAdminParticipantList(document.getElementById('adminSearch').value.toLowerCase());
        updateAdminPresetList();
    }
}

async function showWinnersInGrid(winnersToShow) {
    const availableCells = [];
    for (let i = 0; i < gridSize.total; i++) {
        if (i !== gridSize.center) availableCells.push(i);
    }
    
    const shuffled = availableCells.sort(() => Math.random() - 0.5);
    const winnerCells = shuffled.slice(0, Math.min(winnersToShow.length, availableCells.length));

    for (let i = 0; i < winnerCells.length; i++) {
        const cellIndex = winnerCells[i];
        const cell = document.querySelector(`[data-cell="${cellIndex}"]`);
        const winner = winnersToShow[i];
        
        updateCellContent(cell, winner);
        cell.classList.add('winner-cell');
        
        await new Promise(resolve => setTimeout(resolve, 200));
    }

    await new Promise(resolve => setTimeout(resolve, 2000));

    winnerCells.forEach(cellIndex => {
        const cell = document.querySelector(`[data-cell="${cellIndex}"]`);
        cell.classList.remove('winner-cell');
    });

    initializeRollGrid();
}

function displayWinners(winnersToDisplay) {
    const winnerDisplay = document.getElementById('winnerDisplay');
    
    if (winnersToDisplay.length === 0) {
        winnerDisplay.className = 'winner-placeholder';
        winnerDisplay.innerHTML = '<p class="placeholder-text">🧧 Nhấn nút để quay! 🧧</p>';
        return;
    }

    winnerDisplay.className = 'winner-display';
    
    let html = `<h3 class="winner-title">🎉 ${winnersToDisplay[0].prize.toUpperCase()} 🎉</h3>`;
    html += '<div class="winner-list">';
    
    winnersToDisplay.forEach(winner => {
        html += '<div class="winner-card">';
        displayColumns.forEach(col => {
            if (winner[col]) {
                html += `
                    <div class="winner-info">
                        <span class="info-label">${col}:</span>
                        <span class="info-value">${winner[col]}</span>
                    </div>
                `;
            }
        });
        html += '</div>';
    });
    
    html += '</div>';
    winnerDisplay.innerHTML = html;
}

function updatePrizeList() {
    const prizeListDiv = document.getElementById('prizeList');
    
    if (prizes.length === 0) {
        prizeListDiv.innerHTML = '<p class="prize-empty">🎁 Chưa có giải</p>';
        return;
    }

    let html = '<div class="prize-list">';
    prizes.forEach(prize => {
        html += `
            <div class="prize-item">
                <div class="prize-header">
                    <h4 class="prize-name">🏆 ${prize.name}</h4>
                    <button onclick="handleDeletePrize(${prize.id})" class="btn-delete">
                        <svg class="icon-sm" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path>
                        </svg>
                    </button>
                </div>
                <p class="prize-details">👥 ${prize.winnersPerPrize} người × ${prize.quantity} giải</p>
                <div class="prize-remaining">Còn: ${prize.remaining}</div>
            </div>
        `;
    });
    html += '</div>';
    prizeListDiv.innerHTML = html;
}

function updatePrizeSelect() {
    const select = document.getElementById('prizeSelect');
    const availablePrizes = prizes.filter(p => p.remaining > 0);
    
    let html = '<option value="">Chọn giải...</option>';
    availablePrizes.forEach(prize => {
        const selected = selectedPrizeId === prize.id ? 'selected' : '';
        html += `<option value="${prize.id}" ${selected}>🎁 ${prize.name} (Còn ${prize.remaining})</option>`;
    });
    
    select.innerHTML = html;
    select.disabled = availablePrizes.length === 0;
}

function updateStats() {
    const total = participants.length - absentPeople.length;
    const winnersCount = winners.length;
    const remaining = total - winnersCount;

    document.getElementById('navStatTotal').textContent = total;
    document.getElementById('navStatWinners').textContent = winnersCount;
    document.getElementById('navStatRemaining').textContent = remaining;
}

function updateSpinButton() {
    const spinBtn = document.getElementById('spinBtn');
    const centerSpinBtn = document.getElementById('centerSpinBtn');
    const canSpin = !isSpinning && selectedPrizeId && participants.length > 0;
    
    spinBtn.disabled = !canSpin;
    if (centerSpinBtn) {
        centerSpinBtn.disabled = !canSpin;
    }
}

function exportWinners() {
    if (winners.length === 0) {
        showNotification('Chưa có người trúng!', 'info');
        return;
    }

    const data = winners.map((w, idx) => {
        const row = { STT: idx + 1 };
        excelColumns.forEach(col => {
            row[col] = w[col];
        });
        row['Giải Thưởng'] = w.prize;
        row['Thời Gian'] = w.timestamp;
        return row;
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Người Trúng');
    XLSX.writeFile(wb, `nguoi_trung_${Date.now()}.xlsx`);
}

function openAbsentModal() {
    if (winners.length === 0) {
        showNotification('Chưa có người trúng!', 'info');
        return;
    }
    document.getElementById('absentModal').classList.add('active');
    updateWinnerList();
}

function closeAbsentModal() {
    document.getElementById('absentModal').classList.remove('active');
    document.getElementById('searchAbsent').value = '';
}

function toggleAbsent(participantId) {
    const isAbsent = absentPeople.some(p => p.id === participantId);
    
    if (isAbsent) {
        absentPeople = absentPeople.filter(p => p.id !== participantId);
    } else {
        const person = participants.find(p => p.id === participantId);
        if (person) absentPeople.push(person);
    }
    
    updateWinnerList(document.getElementById('searchAbsent').value.toLowerCase());
    updateAbsentCount();
}

function filterWinners() {
    const searchTerm = document.getElementById('searchAbsent').value.toLowerCase();
    updateWinnerList(searchTerm);
}

function updateWinnerList(searchTerm = '') {
    const winnerListDiv = document.getElementById('winnerListModal');
    
    if (winners.length === 0) {
        winnerListDiv.innerHTML = '<p class="participant-empty">Chưa có người trúng</p>';
        updateModalCounts(0, 0);
        return;
    }
    
    const uniqueWinners = [];
    const seenIds = new Set();
    
    winners.forEach(w => {
        if (!seenIds.has(w.participantId)) {
            seenIds.add(w.participantId);
            uniqueWinners.push(w);
        }
    });
    
    let filteredWinners = uniqueWinners;
    
    if (searchTerm) {
        filteredWinners = uniqueWinners.filter(winner => {
            const searchableText = excelColumns.map(col => 
                String(winner[col] || '').toLowerCase()
            ).join(' ');
            return searchableText.includes(searchTerm) || winner.prize.toLowerCase().includes(searchTerm);
        });
    }
    
    if (filteredWinners.length === 0) {
        winnerListDiv.innerHTML = '<p class="participant-empty">Không tìm thấy</p>';
        updateModalCounts(uniqueWinners.length, absentPeople.length);
        return;
    }
    
    let html = '';
    let absentCount = 0;
    
    filteredWinners.forEach(winner => {
        const isAbsent = absentPeople.some(p => p.id === winner.participantId);
        if (isAbsent) absentCount++;
        
        const statusClass = isAbsent ? 'absent' : '';
        const statusBadgeClass = isAbsent ? 'status-badge-absent' : 'status-badge-available';
        const statusText = isAbsent ? 'Vắng' : 'Có mặt';
        
        html += `
            <div class="participant-card ${statusClass}" onclick="toggleAbsent(${winner.participantId})">
                <div class="participant-card-header">
                    <span class="participant-id">ID: ${winner.participantId}</span>
                    <span class="participant-status-badge ${statusBadgeClass}">${statusText}</span>
                </div>
                <div class="participant-details">
        `;
        
        displayColumns.forEach((col) => {
            if (winner[col]) {
                html += `
                    <div class="participant-field">
                        <span class="field-label">${col}:</span>
                        <span class="field-value">${winner[col]}</span>
                    </div>
                `;
            }
        });
        
        html += `
                </div>
                <div class="participant-prize">
                    <span class="prize-icon">🏆</span>
                    <span class="prize-text">${winner.prize}</span>
                </div>
            </div>
        `;
    });
    
    winnerListDiv.innerHTML = html;
    updateModalCounts(uniqueWinners.length, absentCount);
}

function updateModalCounts(total, absent) {
    const totalEl = document.getElementById('totalWinnersCount');
    const presentEl = document.getElementById('presentCount');
    const absentEl = document.getElementById('absentCountModal');
    
    if (totalEl) totalEl.textContent = total;
    if (presentEl) presentEl.textContent = total - absent;
    if (absentEl) absentEl.textContent = absent;
}

function updateAbsentCount() {
    document.getElementById('navAbsentCount').textContent = absentPeople.length;
}

function exportAbsent() {
    if (absentPeople.length === 0) {
        showNotification('Chưa có người vắng!', 'info');
        return;
    }

    const data = absentPeople.map((person, idx) => {
        const row = { STT: idx + 1 };
        excelColumns.forEach(col => {
            row[col] = person[col];
        });
        return row;
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Người Vắng');
    XLSX.writeFile(wb, `nguoi_vang_${Date.now()}.xlsx`);
}

function handleReset() {
    if (confirm('Reset toàn bộ?')) {
        participants = [];
        winners = [];
        prizes = [];
        absentPeople = [];
        excelColumns = [];
        displayColumns = [];
        selectedPrizeId = null;
        currentWinners = [];
        adminSelectedPerson = null;
        adminSelectedPrize = null;
        adminPresetList = [];
        
        stopRolling();
        
        document.getElementById('fileInput').value = '';
        document.getElementById('prizeSelect').value = '';
        document.getElementById('searchAbsent').value = '';
        
        for (let i = 0; i < gridSize.total; i++) {
            if (i === gridSize.center) continue;
            const cell = document.querySelector(`[data-cell="${i}"]`);
            if (!cell) continue;
            
            const codeEl = cell.querySelector('.employee-code-new');
            const nameEl = cell.querySelector('.employee-name-new');
            const deptEl = cell.querySelector('.employee-dept-new');
            
            if (codeEl) codeEl.textContent = '---';
            if (nameEl) nameEl.textContent = '---';
            if (deptEl) deptEl.textContent = '---';
            
            cell.classList.remove('winner-cell', 'rolling');
        }
        
        updateStats();
        updatePrizeList();
        updatePrizeSelect();
        updateAbsentCount();
        displayWinners([]);
        updateSpinButton();
        closeAbsentModal();
        closeAdminPanel();
        
        showNotification('Đã reset!', 'success');
    }
}