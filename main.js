// --- Firebase Configuration ---
const firebaseConfig = {
    databaseURL: "https://servicedesk-1dadb-default-rtdb.firebaseio.com/",
};

firebase.initializeApp(firebaseConfig);
const db = firebase.database();
document.documentElement.lang = "ko";

// --- Application State ---
let currentUserRole = null; // 'admin' or 'guest'
let currentDate = new Date();
let scheduleData = {};  // Format: { "YYYY-MM-DD_employeeId": "ShiftType" }
let employeesData = {}; // Format: { "empId": { name: "John", category: "Manager" } }
let customHolidaysData = {}; // Custom holiday name overrides
let checkedData = {}; // Format: { "cellKey": true/false } для 연차/대휴 체크상태

// Selections
let sidebarSelectedShift = null;
let selectedKeys = new Set();
let pivotKey = null;
let activeKey = null;
let allCellCoords = {};
let coordsToKey = {};
let isDragging = false;
let draggedEmpId = null;
let clipboardBuffer = null; // { baseRow, baseCol, shifts: { "rowOffset,colOffset": "Shift" } }

// Undo Stack
let undoStack = [];
const MAX_UNDO_STACK = 50;

// Limits
const CATEGORY_LIMITS = {
    "Manager": 2,
    "Weekday AR": 4,
    "Weekend AR": 7
};

const CATEGORY_LABELS = {
    "Manager": "매니저",
    "Weekday AR": "주중 AR",
    "Weekend AR": "주말 AR"
};

const DAYS_KR = ['일', '월', '화', '수', '목', '금', '토'];

// --- DOM References ---
const loginContainer = document.getElementById('login-container');
const mainApp = document.getElementById('main-app');
const loginIdInput = document.getElementById('login-id');
const loginPwInput = document.getElementById('login-pw');
const loginBtn = document.getElementById('login-btn');
const loginError = document.getElementById('login-error');
const logoutBtn = document.getElementById('logout-btn');
const userRoleDisplay = document.getElementById('user-role-display');

const currentMonthDisplay = document.getElementById('current-month-display');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const rosterGrid = document.getElementById('roster-grid');

const empNameInput = document.getElementById('emp-name');
const empCategorySelect = document.getElementById('emp-category');
const empErrorMsg = document.getElementById('emp-error-msg');
const saveEmpBtn = document.getElementById('save-emp-btn');
const copyPrevEmpBtn = document.getElementById('copy-prev-emp-btn');

const memoTextarea = document.getElementById('memo-textarea');
const memoStatus = document.getElementById('memo-status');

// --- Initialization ---
function init() {
    // Check session
    const savedRole = localStorage.getItem('sd_roster_role');
    if (savedRole) {
        showApp(savedRole);
    }

    setupLoginListeners();
    listenToFirebase();
    listenToEmployees(); // Initial load for current month
    setupEventListeners();
    setupGlobalKeyboard();

    // Global Mouse Up to stop dragging
    window.addEventListener('mouseup', () => {
        isDragging = false;
    });
}

function setupLoginListeners() {
    loginBtn.onclick = handleLogin;
    loginPwInput.onkeydown = (e) => { if (e.key === 'Enter') handleLogin(); };
    logoutBtn.onclick = () => {
        localStorage.removeItem('sd_roster_role');
        location.reload();
    };
}

function handleLogin() {
    const id = loginIdInput.value.trim();
    const pw = loginPwInput.value.trim();
    loginError.textContent = "";

    if (id === "admin" && pw === "0626") {
        showApp('admin');
    } else if (id === "ar" && pw === "2222") {
        showApp('guest');
    } else {
        loginError.textContent = "아이디 또는 비밀번호가 잘못되었습니다.";
    }
}

function showApp(role) {
    currentUserRole = role;
    localStorage.setItem('sd_roster_role', role);
    loginContainer.style.display = 'none';
    mainApp.style.display = 'flex';
    userRoleDisplay.textContent = role === 'admin' ? "관리자 모드" : "게스트 모드";

    // Hide administrative UI for guests
    if (role === 'guest') {
        document.querySelector('.employee-sidebar .sidebar-body').style.display = 'none';
        document.querySelector('.employee-sidebar .sidebar-header').style.display = 'none';

        if (memoTextarea) {
            memoTextarea.disabled = true;
            memoTextarea.placeholder = "관리자가 작성한 메모입니다.";
        }
    } else {
        if (memoTextarea) {
            memoTextarea.disabled = false;
            memoTextarea.placeholder = "공지사항이나 메모를 입력하세요...";
        }
    }

    renderRoster();
}

function hasPermission(action, targetEmpId = null) {
    if (currentUserRole === 'admin') return true;

    // Guest restrictions
    if (action === 'edit_shift') {
        if (!targetEmpId) return false;
        const emp = employeesData[targetEmpId];
        return emp && (emp.category === 'Weekday AR' || emp.category === 'Weekend AR');
    }

    // Guest cannot do these actions
    if (['add_employee', 'delete_employee', 'edit_holiday', 'reorder_employee', 'edit_checkbox', 'edit_memo'].includes(action)) {
        return false;
    }

    return false;
}

// --- Firebase Sync ---
function getMonthKey() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth() + 1;
    return {
        val: `${year}-${String(month).padStart(2, '0')}`,
        alt: `${year}-${month}`
    };
}

let activeEmpRef = null;
function listenToEmployees() {
    if (activeEmpRef) activeEmpRef.off();

    const keys = getMonthKey();
    console.log("Searching for employees in:", keys.val, "or", keys.alt);

    activeEmpRef = db.ref('employees/' + keys.val);
    activeEmpRef.on('value', (snapshot) => {
        const data = snapshot.val();
        if (data) {
            employeesData = data;
            console.log("Data found at primary key:", keys.val);
            renderRoster();
        } else {
            // If primary is empty, try alternative (unpadded) key
            db.ref('employees/' + keys.alt).once('value').then(snap => {
                const altData = snap.val();
                if (altData) {
                    employeesData = altData;
                    console.warn("Data found only at alternative key!", keys.alt);
                } else {
                    employeesData = {};
                    console.log("No data found in either key.");
                }
                renderRoster();
            });
        }
    }, (error) => {
        console.error("Firebase Employee Load Error:", error);
        alert("직원 정보를 불러오지 못했습니다. (오류: " + error.message + ")");
    });
}

function copyPreviousMonthEmployees() {
    if (!hasPermission('add_employee')) return;

    const prevDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
    const prevKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, '0')}`;
    const prevKeyAlt = `${prevDate.getFullYear()}-${prevDate.getMonth() + 1}`;
    const currKey = getMonthKey().val;

    if (Object.keys(employeesData).length > 0) {
        if (!confirm("현재 달의 직원 목록이 비어있지 않습니다. 무시하고 전월 목록을 복사하시겠습니까? (기존 목록 뒤에 추가됩니다)")) return;
    }

    // Try primary, then alt for previous month
    db.ref('employees/' + prevKey).once('value').then(snap => {
        let prevEmps = snap.val();

        if (!prevEmps) {
            return db.ref('employees/' + prevKeyAlt).once('value').then(altSnap => altSnap.val());
        }
        return prevEmps;
    }).then(prevEmps => {
        if (!prevEmps) {
            alert("가져올 전월 데이터가 없습니다.");
            return;
        }

        const updates = {};
        Object.values(prevEmps).forEach(emp => {
            const newId = db.ref('employees/' + currKey).push().key;
            updates[newId] = { ...emp };
        });

        db.ref('employees/' + currKey).update(updates)
            .then(() => alert("전월 직원 목록을 성공적으로 가져왔습니다."))
            .catch(err => console.error("Copy failed:", err));
    });
}

// --- Firebase Sync ---
function listenToFirebase() {
    const errorCallback = (error) => {
        console.error("Firebase Sync Error:", error);
        if (error.code === 'PERMISSION_DENIED') {
            alert("Firebase 접근 권한이 없습니다. 보안 규칙을 확인해주세요.");
        }
    };

    db.ref('schedule').on('value', (snapshot) => {
        scheduleData = snapshot.val() || {};
        renderRoster();
    }, errorCallback);

    db.ref('customHolidays').on('value', (snapshot) => {
        customHolidaysData = snapshot.val() || {};
        renderRoster();
    }, errorCallback);

    db.ref('checkedStatus').on('value', (snapshot) => {
        checkedData = snapshot.val() || {};
        renderRoster();
    }, errorCallback);

    db.ref('memo').on('value', (snapshot) => {
        const memoVal = snapshot.val() || "";
        if (memoTextarea && document.activeElement !== memoTextarea) {
            memoTextarea.value = memoVal;
        }
    }, errorCallback);
}

function saveEmployeeToFirebase(name, category) {
    if (!hasPermission('add_employee')) return;
    const key = getMonthKey().val;
    const newEmpRef = db.ref(`employees/${key}`).push();
    newEmpRef.set({ name, category, sortOrder: Date.now() })
        .catch((error) => console.error("Employee save failed:", error));
}

function deleteEmployee(empId) {
    if (!hasPermission('delete_employee')) return;
    if (confirm("정말 이 직원을 삭제하시겠습니까?")) {
        const keys = getMonthKey();
        
        // Try deleting from both potential keys to be sure
        db.ref(`employees/${keys.val}/${empId}`).remove();
        db.ref(`employees/${keys.alt}/${empId}`).remove();

        // Also clean up their schedule data
        Object.keys(scheduleData).forEach(key => {
            if (key.endsWith('_' + empId)) {
                db.ref('schedule/' + key).remove();
            }
        });
    }
}

function handleDrop(targetEmpId) {
    if (!hasPermission('reorder_employee')) return;
    const key = getMonthKey();
    if (!draggedEmpId || draggedEmpId === targetEmpId) return;

    const sourceEmp = employeesData[draggedEmpId];
    const targetEmp = employeesData[targetEmpId];
    if (!sourceEmp || !targetEmp) return;
    if (sourceEmp.category !== targetEmp.category) return;

    const tempOrder = sourceEmp.sortOrder || 0;
    const targetOrder = targetEmp.sortOrder || 0;

    let finalSourceOrder = targetOrder;
    let finalTargetOrder = tempOrder;
    if (finalSourceOrder === finalTargetOrder) {
        finalSourceOrder += 1;
    }

    db.ref(`employees/${key}/${draggedEmpId}/sortOrder`).set(finalSourceOrder);
    db.ref(`employees/${key}/${targetEmpId}/sortOrder`).set(finalTargetOrder);

    draggedEmpId = null;
}

function applyShiftChanges(updates, pushToUndo = true) {
    if (Object.keys(updates).length === 0) return;

    if (pushToUndo) {
        const previousState = {};
        for (const key in updates) {
            previousState[key] = scheduleData[key] || "";
        }
        undoStack.push(previousState);
        if (undoStack.length > MAX_UNDO_STACK) undoStack.shift();
    }

    for (const key in updates) {
        const firstUnderscore = key.indexOf('_');
        if (firstUnderscore !== -1) {
            const empId = key.substring(firstUnderscore + 1);
            if (hasPermission('edit_shift', empId)) {
                let val = updates[key];
                const value = (typeof val === 'string') ? val.toUpperCase() : val;
                db.ref('schedule/' + key).set(value === "" ? null : value).catch(console.error);
            }
        }
    }
}

function handleUndo() {
    if (undoStack.length === 0) return;
    const prevState = undoStack.pop();
    applyShiftChanges(prevState, false);
}

function saveCustomHolidayToFirebase(dateStr, holidayName) {
    if (!hasPermission('edit_holiday')) return;
    db.ref('customHolidays/' + dateStr).set(holidayName === "" ? null : holidayName)
        .catch((error) => console.error("Holiday save failed:", error));
}

// --- Holiday Logic ---
const PUBLIC_HOLIDAYS_FIXED = {
    "01-01": "신정", "03-01": "삼일절", "05-05": "어린이날", "06-06": "현충일",
    "08-15": "광복절", "10-03": "개천절", "10-09": "한글날", "12-25": "성탄절"
};

const HOLIDAYS_VAR_2026 = {
    "2026-02-16": "설날 연휴", "2026-02-17": "설날", "2026-02-18": "설날 연휴",
    "2026-03-02": "대체공휴일", "2026-05-24": "부처님오신날", "2026-05-25": "대체공휴일",
    "2026-08-17": "대체공휴일", "2026-09-24": "추석 연휴", "2026-09-25": "추석",
    "2026-09-26": "추석 연휴", "2026-10-05": "대체공휴일",
};

function getHolidayName(year, month, day) {
    const mmdd = String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    const yyyymmdd = year + '-' + mmdd;

    // Check custom overrides first
    if (customHolidaysData[yyyymmdd]) return customHolidaysData[yyyymmdd];

    if (year === 2026 && HOLIDAYS_VAR_2026[yyyymmdd]) return HOLIDAYS_VAR_2026[yyyymmdd];
    if (PUBLIC_HOLIDAYS_FIXED[mmdd]) return PUBLIC_HOLIDAYS_FIXED[mmdd];
    return null;
}

// --- Selection Helpers ---
function performRangeSelection(startKey, endKey) {
    const start = allCellCoords[startKey];
    const end = allCellCoords[endKey];
    if (!start || !end) return;

    const r1 = Math.min(start.row, end.row);
    const r2 = Math.max(start.row, end.row);
    const c1 = Math.min(start.col, end.col);
    const c2 = Math.max(start.col, end.col);

    selectedKeys.clear();
    Object.keys(allCellCoords).forEach(k => {
        const coord = allCellCoords[k];
        if (coord.row >= r1 && coord.row <= r2 && coord.col >= c1 && coord.col <= c2) {
            selectedKeys.add(k);
        }
    });
}

function handleMouseDown(e, key) {
    // If Painter mode is active, don't drag select, just apply
    if (sidebarSelectedShift !== null) {
        applyShiftToTarget(key, sidebarSelectedShift);
        return;
    }

    if (e.ctrlKey || e.metaKey) {
        if (selectedKeys.has(key)) selectedKeys.delete(key);
        else selectedKeys.add(key);
        pivotKey = key;
        activeKey = key;
    } else if (e.shiftKey && pivotKey) {
        activeKey = key;
        performRangeSelection(pivotKey, activeKey);
    } else {
        selectedKeys.clear();
        selectedKeys.add(key);
        pivotKey = key;
        activeKey = key;
        isDragging = true;
    }
    renderRoster();
}

function handleMouseOver(e, key) {
    if (isDragging && pivotKey) {
        performRangeSelection(pivotKey, key);
        renderRoster();
    }
}

function applyShiftToTarget(key, shift) {
    const info = allCellCoords[key];
    if (!info) return;

    if (selectedKeys.has(key)) {
        // Toggle logic for bulk: if first cell in selection already has this shift, we clear all.
        // Otherwise, we set all. (Standard on/off toggle behavior)
        const currentShift = scheduleData[key];
        const targetShift = (currentShift === shift) ? "" : shift;

        const updates = {};
        selectedKeys.forEach(k => {
            const ki = allCellCoords[k];
            if (ki) updates[`${ki.date}_${ki.eId}`] = targetShift;
        });
        applyShiftChanges(updates);
    } else {
        const currentShift = scheduleData[key];
        const targetShift = (currentShift === shift) ? "" : shift;
        applyShiftChanges({ [`${info.date}_${info.eId}`]: targetShift });
    }
}

// --- Display Logic ---
function renderRoster() {
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();

    currentMonthDisplay.textContent = `${year}년 ${month + 1}월`;
    rosterGrid.innerHTML = '';

    const daysInMonth = new Date(year, month + 1, 0).getDate();
    rosterGrid.style.gridTemplateColumns = `minmax(120px, auto) repeat(${daysInMonth}, minmax(32px, 1fr))`;

    if (Object.keys(employeesData).length === 0) {
        const emptyMsg = document.createElement('div');
        emptyMsg.style.cssText = 'grid-column: 1 / -1; padding: 3rem; text-align: center; color: #94a3b8; font-size: 0.9rem; background: #f8fafc; border-bottom: 2px solid #e2e8f0;';
        emptyMsg.innerHTML = `
            <div style="font-weight: 600; margin-bottom: 0.5rem;">현재 월(${getMonthKey().val})에 등록된 데이터가 없습니다.</div>
            <div>좌측 '직원 관리'에서 직원을 추가하거나 [전월 목록 가져오기] 버튼을 눌러주세요.</div>
        `;
        rosterGrid.appendChild(emptyMsg);
    }

    // 1. Header (일자)
    const cornerCell = document.createElement('div');
    cornerCell.className = 'r-cell r-corner';
    cornerCell.textContent = '일자';
    rosterGrid.appendChild(cornerCell);

    for (let day = 1; day <= daysInMonth; day++) {
        const dayOfWeek = new Date(year, month, day).getDay();
        const headerCell = document.createElement('div');
        headerCell.className = `r-cell r-header`;
        if (dayOfWeek === 0 || dayOfWeek === 6 || getHolidayName(year, month + 1, day)) {
            headerCell.classList.add('is-holiday');
            headerCell.classList.add('is-holiday-bg');
        }
        if (dayOfWeek === 0) {
            headerCell.classList.add('sun-border');
        }

        const numSpan = document.createElement('span');
        numSpan.className = 'date-num';
        numSpan.textContent = day;

        const dowSpan = document.createElement('span');
        dowSpan.textContent = DAYS_KR[dayOfWeek];

        headerCell.appendChild(numSpan);
        headerCell.appendChild(dowSpan);
        rosterGrid.appendChild(headerCell);
    }

    // Holiday Row
    const holidayRowLabel = document.createElement('div');
    holidayRowLabel.className = 'r-cell r-col-header holiday-row-label';
    holidayRowLabel.textContent = '공휴일';
    rosterGrid.appendChild(holidayRowLabel);

    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = new Date(year, month, day).getDay();
        const holidayName = getHolidayName(year, month + 1, day);
        const holiCell = document.createElement('div');
        holiCell.className = 'r-cell holiday-cell';
        if (dayOfWeek === 0 || dayOfWeek === 6 || holidayName) {
            // holiCell.classList.add('is-holiday-bg');
        }
        if (dayOfWeek === 0) {
            holiCell.classList.add('sun-border');
        }
        if (holidayName) {
            holiCell.textContent = holidayName;
        }

        // Make holidays editable for admin only
        holiCell.contentEditable = hasPermission('edit_holiday') ? "true" : "false";
        holiCell.onblur = () => {
            if (!hasPermission('edit_holiday')) return;
            const newValue = holiCell.innerText.trim();
            saveCustomHolidayToFirebase(dateStr, newValue);
        };
        // Also support Enter to save
        holiCell.onkeydown = (e) => {
            if (e.key === 'Enter') {
                if (e.altKey) {
                    e.preventDefault();
                    document.execCommand('insertLineBreak');
                    return;
                }
                e.preventDefault();
                holiCell.blur();
            }
        };

        rosterGrid.appendChild(holiCell);
    }

    // Workforce Count Row
    const countRowLabel = document.createElement('div');
    countRowLabel.className = 'r-cell r-col-header count-row-label';
    countRowLabel.textContent = '근무인원';
    rosterGrid.appendChild(countRowLabel);

    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = new Date(year, month, day).getDay();
        let count = 0;
        Object.keys(employeesData).forEach(empId => {
            const emp = employeesData[empId];
            if (emp.category === 'Weekday AR' || emp.category === 'Weekend AR') {
                const shiftVal = scheduleData[`${dateStr}_${empId}`];
                // Count A, A(연), B shifts for ARs only
                if (shiftVal && (shiftVal === 'A' || shiftVal === 'A(연)' || shiftVal === 'B')) {
                    count++;
                }
            }
        });

        const countCell = document.createElement('div');
        countCell.className = 'r-cell count-cell';
        if (dayOfWeek === 0 || dayOfWeek === 6 || getHolidayName(year, month + 1, day)) {
            // countCell.classList.add('is-holiday-bg');
        }
        if (dayOfWeek === 0) {
            countCell.classList.add('sun-border');
        }
        countCell.textContent = count > 0 ? count : '0';
        rosterGrid.appendChild(countCell);
    }

    // Employees
    const categories = ["Manager", "Weekday AR", "Weekend AR"];
    let currentRowIdx = 0;
    allCellCoords = {};
    coordsToKey = {};

    categories.forEach((cat) => {
        const catEmployees = Object.entries(employeesData)
            .filter(([id, emp]) => emp.category === cat)
            .sort((a, b) => (a[1].sortOrder || 0) - (b[1].sortOrder || 0));
        const catHeaderDiv = document.createElement('div');
        catHeaderDiv.className = `cat-row cat-${cat.split(' ')[0]}`;
        catHeaderDiv.textContent = CATEGORY_LABELS[cat];
        rosterGrid.appendChild(catHeaderDiv);

        catEmployees.forEach(([empId, emp]) => {
            try {
                const nameCell = document.createElement('div');
                nameCell.className = 'r-cell r-col-header employee-name-cell';
                nameCell.draggable = hasPermission('reorder_employee');
                nameCell.innerHTML = `
                    <div class="drag-handle" style="display: ${hasPermission('reorder_employee') ? 'block' : 'none'}">⠿</div>
                    <span class="emp-display-name">${emp.name}</span>
                    ${hasPermission('delete_employee') ? `<button class="danger-btn tiny" onmousedown="event.stopPropagation()" onclick="deleteEmployee('${empId}')">X</button>` : ''}
                `;

                nameCell.ondragstart = (e) => {
                    draggedEmpId = empId;
                    nameCell.classList.add('dragging');
                    e.dataTransfer.effectAllowed = 'move';
                };
                nameCell.ondragend = () => {
                    nameCell.classList.remove('dragging');
                };
                nameCell.ondragover = (e) => {
                    e.preventDefault();
                    if (draggedEmpId && draggedEmpId !== empId && employeesData[draggedEmpId].category === emp.category) {
                        nameCell.classList.add('drag-over');
                    }
                };
                nameCell.ondragleave = () => {
                    nameCell.classList.remove('drag-over');
                };
                nameCell.ondrop = (e) => {
                    e.preventDefault();
                    nameCell.classList.remove('drag-over');
                    handleDrop(empId);
                };

                rosterGrid.appendChild(nameCell);

                for (let day = 1; day <= daysInMonth; day++) {
                    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                    const cellKey = `${dateStr}_${empId}`;

                    allCellCoords[cellKey] = { row: currentRowIdx, col: day, date: dateStr, eId: empId };
                    coordsToKey[`${currentRowIdx},${day}`] = cellKey;

                    const cell = document.createElement('div');
                    cell.className = 'r-cell r-entry';
                    cell.dataset.key = cellKey;
                    cell.setAttribute('lang', 'ko');
                    cell.setAttribute('inputmode', 'text');
                    cell.setAttribute('spellcheck', 'false');
                    cell.tabIndex = 0;

                    const dayOfWeek = new Date(year, month, day).getDay();
                    if (dayOfWeek === 0 || dayOfWeek === 6 || getHolidayName(year, month + 1, day)) {
                        // cell.classList.add('is-holiday-bg');
                    }
                    if (dayOfWeek === 0) {
                        cell.classList.add('sun-border');
                    }
                    if (selectedKeys.has(cellKey)) cell.classList.add('selected-cell');
                    if (activeKey === cellKey) {
                        cell.classList.add('active-cell');
                        cell.classList.add('is-navigating');
                    }

                    const shiftVal = scheduleData[cellKey];
                    if (shiftVal) {
                        // Clean any accidental '✓' or 'M' that might have been saved to Firebase
                        const cleanShiftVal = shiftVal.toString().replace(/[✓M]/g, '').trim();
                        cell.textContent = cleanShiftVal;
                        const safeShift = cleanShiftVal.replace(/[\(\)\s\n]/g, '');
                        cell.classList.add(`shift-${safeShift}`);

                        if (cleanShiftVal === '연차' || cleanShiftVal === '대휴' || cleanShiftVal === '공휴') {
                            const chk = document.createElement('div');
                            chk.className = 'leave-checkbox-ui';
                            if (checkedData[cellKey]) chk.classList.add('is-checked');
                            chk.innerHTML = checkedData[cellKey] ? 'M' : '';

                            chk.onmousedown = (e) => { e.stopPropagation(); };
                            chk.onclick = (e) => {
                                e.stopPropagation();
                                if (!hasPermission('edit_checkbox')) return;
                                const currentChecked = !!checkedData[cellKey];
                                db.ref('checkedStatus/' + cellKey).set(currentChecked ? false : true)
                                    .catch((error) => console.error("Check toggle failed:", error));
                            };
                            cell.appendChild(chk);
                        }
                    }

                    cell.onmousedown = (e) => {
                        if (document.activeElement === cell) return;
                        if (selectedKeys.size === 1 && selectedKeys.has(cellKey)) return;
                        e.preventDefault();
                        handleMouseDown(e, cellKey);
                    };
                    cell.onmouseover = (e) => handleMouseOver(e, cellKey);
                    cell.ondblclick = () => {
                        if (hasPermission('edit_shift', empId)) cell.focus();
                    };
                    cell.contentEditable = hasPermission('edit_shift', empId) ? "true" : "false";
                    cell.onblur = () => {
                        // Strip '✓' and 'M' from innerText so it isn't accidentally saved alongside the text
                        const newValue = cell.innerText.replace(/[✓M]/g, '').trim();
                        if (newValue !== (scheduleData[cellKey] || "")) {
                            applyShiftChanges({ [cellKey]: newValue });
                        }
                    };
                    cell.onkeydown = (e) => {
                        if (e.key === 'Enter') {
                            if (e.altKey) { e.preventDefault(); document.execCommand('insertLineBreak'); return; }
                            e.preventDefault(); cell.blur(); moveSelection(1, 0);
                        }
                        if (e.key === 'Tab') { e.preventDefault(); cell.blur(); moveSelection(0, e.shiftKey ? -1 : 1); }
                        if (e.key === 'Escape') { cell.textContent = scheduleData[cellKey] || ""; cell.blur(); }
                        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && cell.classList.contains('is-navigating')) {
                            cell.classList.remove('is-navigating');
                            cell.innerText = "";
                        }
                    };
                    rosterGrid.appendChild(cell);
                }
                currentRowIdx++;
            } catch (err) {
                console.error("Error rendering employee row:", err);
            }
        });
    });
}

// --- Excel Navigation Logic ---
function moveSelection(dRow, dCol, shift = false, ctrl = false) {
    if (!activeKey || !allCellCoords[activeKey]) return;
    const current = allCellCoords[activeKey];

    let targetRow = current.row;
    let targetCol = current.col;

    if (ctrl) {
        // Find grid boundaries
        const rows = [...new Set(Object.values(allCellCoords).map(c => c.row))];
        const cols = [...new Set(Object.values(allCellCoords).map(c => c.col))];
        const minR = Math.min(...rows), maxR = Math.max(...rows);
        const minC = Math.min(...cols), maxC = Math.max(...cols);

        if (dRow < 0) targetRow = minR;
        if (dRow > 0) targetRow = maxR;
        if (dCol < 0) targetCol = minC;
        if (dCol > 0) targetCol = maxC;
    } else {
        targetRow += dRow;
        targetCol += dCol;
    }

    const nextKey = coordsToKey[`${targetRow},${targetCol}`];
    if (nextKey) {
        activeKey = nextKey;
        if (shift) {
            performRangeSelection(pivotKey, activeKey);
        } else {
            selectedKeys.clear();
            selectedKeys.add(nextKey);
            pivotKey = nextKey;
        }
        renderRoster();

        // Pre-emptive focus for IME support
        setTimeout(() => {
            const activeEl = document.querySelector('.active-cell');
            if (activeEl) {
                activeEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                activeEl.focus();
            }
        }, 0);
    }
}

function setupGlobalKeyboard() {
    window.onkeydown = (e) => {
        const isCtrl = e.ctrlKey || e.metaKey;
        const isEditing = document.activeElement.contentEditable === "true";

        // If user is typing in an input, textarea or focused in a cell, don't trigger global nav
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT' || e.target.tagName === 'TEXTAREA') return;

        // Allow shortcuts and navigation keys even when editing a cell
        const allowedKeys = ['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'Enter', 'Tab', 'Delete', 'Backspace'];
        if (isEditing && !isCtrl && !allowedKeys.includes(e.key)) {
            return;
        }

        if ((selectedKeys.size > 0 && pivotKey) || activeKey) {
            // Arrow Navigation
            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
                e.preventDefault();
                const dRow = e.key === 'ArrowUp' ? -1 : (e.key === 'ArrowDown' ? 1 : 0);
                const dCol = e.key === 'ArrowLeft' ? -1 : (e.key === 'ArrowRight' ? 1 : 0);
                moveSelection(dRow, dCol, e.shiftKey, e.ctrlKey);
                return;
            }

            // Delete / Backspace: Clear entire cell content immediately
            if (e.key === 'Delete' || e.key === 'Backspace') {
                e.preventDefault();
                const updates = {};
                selectedKeys.forEach(k => {
                    const info = allCellCoords[k];
                    if (info) updates[`${info.date}_${info.eId}`] = "";

                    // Instant UI feedback
                    const cellEl = document.querySelector(`.r-entry[data-key="${k}"]`);
                    if (cellEl) {
                        cellEl.innerText = "";
                        cellEl.classList.remove('is-navigating');
                        // Clean shift classes
                        const classes = Array.from(cellEl.classList).filter(c => c.startsWith('shift-'));
                        classes.forEach(c => cellEl.classList.remove(c));
                    }
                });
                applyShiftChanges(updates);
                return;
            }

            // Instant Typing: Handled locally in cell.onkeydown for focused cells,
            // but for safety if window still catches it:
            if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !document.activeElement.classList.contains('r-entry')) {
                const targetCell = document.querySelector(`.r-entry.active-cell`);
                if (targetCell) {
                    targetCell.focus();
                    targetCell.innerText = "";
                    targetCell.classList.remove('is-navigating');
                }
            }

            // Copy/Paste/Undo
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
                e.preventDefault();
                handleCopy();
            }
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
                // We let the native 'paste' event handle this to support external data
            }
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
                e.preventDefault();
                handleUndo();
            }
        }
    };

    // Global Paste Listener for external (Excel) data
    window.addEventListener('paste', (e) => {
        // If focused on an input, let default behavior happen
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
        // If editing a cell, let default handle it (usually)
        if (document.activeElement.contentEditable === "true") return;

        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text');
        if (!text) return;

        // Parse Excel/TSV data
        const rows = text.split(/\r?\n/).filter(line => line.trim() !== "").map(row => row.split('\t'));
        if (rows.length > 0) {
            handleExternalPaste(rows);
        }
    });
}

function handleCopy() {
    if (selectedKeys.size === 0) return;

    // Find top-left to determine offsets
    let minR = Infinity, minC = Infinity;
    selectedKeys.forEach(key => {
        const coord = allCellCoords[key];
        if (coord.row < minR) minR = coord.row;
        if (coord.col < minC) minC = coord.col;
    });

    const shifts = {};
    selectedKeys.forEach(key => {
        const coord = allCellCoords[key];
        const val = scheduleData[key] || "";
        shifts[`${coord.row - minR},${coord.col - minC}`] = val;
    });

    clipboardBuffer = { shifts };
    console.log("Copied", Object.keys(shifts).length, "cells");
}

function handlePaste() {
    if (clipboardBuffer) {
        handleInternalPaste();
    }
}

function handleInternalPaste() {
    if (!clipboardBuffer || !activeKey) return;
    const targetBase = allCellCoords[activeKey];
    if (!targetBase) return;

    const updates = {};
    Object.entries(clipboardBuffer.shifts).forEach(([offset, value]) => {
        const [dr, dc] = offset.split(',').map(Number);
        const tr = targetBase.row + dr;
        const tc = targetBase.col + dc;

        const targetKey = coordsToKey[`${tr},${tc}`];
        if (targetKey) {
            const targetInfo = allCellCoords[targetKey];
            if (hasPermission('edit_shift', targetInfo.eId)) {
                updates[`${targetInfo.date}_${targetInfo.eId}`] = value;
            }
        }
    });
    applyShiftChanges(updates);
}

function handleExternalPaste(dataMatrix) {
    if (!activeKey) return;
    const targetBase = allCellCoords[activeKey];
    if (!targetBase) return;

    const updates = {};
    dataMatrix.forEach((row, ri) => {
        row.forEach((cellValue, ci) => {
            const tr = targetBase.row + ri;
            const tc = targetBase.col + ci;
            const targetKey = coordsToKey[`${tr},${tc}`];
            if (targetKey) {
                const targetInfo = allCellCoords[targetKey];
                if (hasPermission('edit_shift', targetInfo.eId)) {
                    updates[`${targetInfo.date}_${targetInfo.eId}`] = cellValue.trim().toUpperCase();
                }
            }
        });
    });
    applyShiftChanges(updates);
}

// --- Interaction ---
function setupEventListeners() {
    prevBtn.onclick = () => { currentDate.setMonth(currentDate.getMonth() - 1); listenToEmployees(); };
    nextBtn.onclick = () => { currentDate.setMonth(currentDate.getMonth() + 1); listenToEmployees(); };

    copyPrevEmpBtn.onclick = copyPreviousMonthEmployees;

    saveEmpBtn.onclick = () => {
        const name = empNameInput.value.trim();
        const cat = empCategorySelect.value;
        empErrorMsg.textContent = '';

        if (!name) {
            empErrorMsg.textContent = '이름을 입력해주세요.';
            return;
        }

        const currentCount = Object.values(employeesData).filter(e => e.category === cat).length;
        if (currentCount >= CATEGORY_LIMITS[cat]) {
            empErrorMsg.textContent = `해당 범주는 최대 ${CATEGORY_LIMITS[cat]}명까지만 추가할 수 있습니다.`;
            return;
        }

        saveEmployeeToFirebase(name, cat);
        empNameInput.value = '';
    };

    empNameInput.onkeydown = (e) => {
        if (e.key === 'Enter') {
            saveEmpBtn.onclick();
        }
    };

    document.querySelectorAll('.bulk-shift-btn').forEach(btn => {
        btn.onclick = () => {
            const shift = btn.dataset.shift;

            if (selectedKeys.size > 0) {
                // Bulk Toggle: If already have this shift, clear it.
                // We'll check the first selected key for the current state.
                const firstKey = Array.from(selectedKeys)[0];
                const currentShift = scheduleData[firstKey];
                const targetShift = (currentShift === shift) ? "" : shift;

                const updates = {};
                selectedKeys.forEach(key => {
                    const info = allCellCoords[key];
                    if (info) updates[`${info.date}_${info.eId}`] = targetShift;
                });
                applyShiftChanges(updates);

                selectedKeys.clear();
                sidebarSelectedShift = null;
                document.querySelectorAll('.bulk-shift-btn').forEach(b => b.classList.remove('active'));
                renderRoster();
                return;
            }

            if (sidebarSelectedShift === shift) {
                sidebarSelectedShift = null;
                btn.classList.remove('active');
            } else {
                document.querySelectorAll('.bulk-shift-btn').forEach(b => b.classList.remove('active'));
                sidebarSelectedShift = shift;
                btn.classList.add('active');
            }
        };
    });

    if (memoTextarea) {
        let saveTimeout = null;
        memoTextarea.oninput = () => {
            if (!hasPermission('edit_memo')) return;

            if (memoStatus) memoStatus.textContent = "저장 중...";

            clearTimeout(saveTimeout);
            saveTimeout = setTimeout(() => {
                db.ref('memo').set(memoTextarea.value)
                    .then(() => { if (memoStatus) memoStatus.textContent = "저장 완료"; })
                    .catch(e => { if (memoStatus) memoStatus.textContent = "저장 실패"; console.error(e); });
            }, 1000);
        };
    }
}

window.deleteEmployee = deleteEmployee;
init();
