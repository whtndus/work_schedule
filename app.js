// ===== State =====
const state = {
    year: 2026,
    month: 2, // 1-indexed
    names: ['직원A', '직원B', '직원C', '직원D'],
    schedule: null,
};

const SHIFTS = [
    { id: 7, label: '7시', time: '07:00~15:00' },
    { id: 9, label: '9시', time: '09:00~17:00' },
    { id: 13, label: '1시', time: '13:00~21:00' },
];

const DAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

// Off-day role templates (0=Sun ... 6=Sat)
// Each role defines which days-of-week are off (consecutive pairs + 1 solo)
// Role 0: Sun+Mon (consecutive)
// Role 1: Tue+Wed (consecutive)
// Role 2: Thu+Fri (consecutive)
// Role 3: Sat only → but next week transitions to Role 0 (Sun+Mon)
//         so cross-week gives: Sat + Sun + Mon = 3 consecutive off days!
const OFF_PATTERNS = [
    [0, 1], // Role 0: Sun, Mon off
    [2, 3], // Role 1: Tue, Wed off
    [4, 5], // Role 2: Thu, Fri off
    [6],    // Role 3: Sat off (connects to next week's Role 0)
];

// ===== DOM Elements =====
const monthDisplay = document.getElementById('monthDisplay');
const prevMonthBtn = document.getElementById('prevMonth');
const nextMonthBtn = document.getElementById('nextMonth');
const generateBtn = document.getElementById('generateBtn');
const printBtn = document.getElementById('printBtn');
const excelBtn = document.getElementById('excelBtn');
const calendarGrid = document.getElementById('calendarGrid');
const calendarTitle = document.getElementById('calendarTitle');
const legendSection = document.getElementById('legendSection');
const calendarSection = document.getElementById('calendarSection');
const summarySection = document.getElementById('summarySection');
const summaryGrid = document.getElementById('summaryGrid');
const tableSection = document.getElementById('tableSection');
const scheduleTable = document.getElementById('scheduleTable');

// ===== Initialization =====
updateMonthDisplay();

prevMonthBtn.addEventListener('click', () => changeMonth(-1));
nextMonthBtn.addEventListener('click', () => changeMonth(1));
generateBtn.addEventListener('click', generate);
printBtn.addEventListener('click', () => window.print());
excelBtn.addEventListener('click', exportToExcel);

function changeMonth(delta) {
    state.month += delta;
    if (state.month < 1) {
        state.month = 12;
        state.year--;
    } else if (state.month > 12) {
        state.month = 1;
        state.year++;
    }
    updateMonthDisplay();
}

function updateMonthDisplay() {
    monthDisplay.textContent = `${state.year}년 ${state.month}월`;
}

function readNames() {
    state.names = [
        document.getElementById('name1').value || '직원A',
        document.getElementById('name2').value || '직원B',
        document.getElementById('name3').value || '직원C',
        document.getElementById('name4').value || '직원D',
    ];
}

// ===== Schedule Generation Algorithm =====
// Consecutive off-day rotation system:
//
// 4-week cycle, each worker rotates through 4 roles:
//   Role 0 → off Sun+Mon (2 consecutive)
//   Role 1 → off Tue+Wed (2 consecutive)
//   Role 2 → off Thu+Fri (2 consecutive)
//   Role 3 → off Sat (1 day, but next week → Role 0 = Sat+Sun+Mon 3 consecutive!)
//
// Worker's role each week = (workerIndex + weekIndex) % 4
// This guarantees D→A transition: whoever has Sat-only off gets Sun+Mon next week.
//
// Per 4-week cycle each worker:
//   Work days: 5+5+5+6 = 21 days
//   Off days: 2+2+2+1 = 7 days
//   Hours: 168h (avg 42h/week) — one week at 48h is offset by 3-day rest next week

function generateSchedule(year, month) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const schedule = {};

    // Build day info
    const days = [];
    for (let d = 1; d <= daysInMonth; d++) {
        days.push({ date: d, dayOfWeek: new Date(year, month - 1, d).getDay() });
    }

    // Group into weeks (Sun=0 starts a new week)
    const weeks = [];
    let currentWeek = [];
    for (const day of days) {
        if (day.dayOfWeek === 0 && currentWeek.length > 0) {
            weeks.push(currentWeek);
            currentWeek = [];
        }
        currentWeek.push(day);
    }
    if (currentWeek.length > 0) weeks.push(currentWeek);

    // Track shift counts per worker for fair rotation: [worker][shiftIndex]
    const shiftCounts = [
        [0, 0, 0],
        [0, 0, 0],
        [0, 0, 0],
        [0, 0, 0],
    ];

    for (let wi = 0; wi < weeks.length; wi++) {
        const week = weeks[wi];

        for (const day of week) {
            // Determine who is off based on role rotation
            // Every day of the week maps to exactly one role via OFF_PATTERNS,
            // so there is always exactly 1 person off.
            let offWorker = -1;
            for (let w = 0; w < 4; w++) {
                const role = (w + wi) % 4;
                if (OFF_PATTERNS[role].includes(day.dayOfWeek)) {
                    offWorker = w;
                    break;
                }
            }

            const workingWorkers = [0, 1, 2, 3].filter(w => w !== offWorker);

            // Assign shifts using greedy balancing:
            // For each shift slot, pick the worker with the fewest assignments of that shift.
            const shifts = {};
            const availableWorkers = [...workingWorkers];

            // Try all 6 possible worker→shift permutations, pick the one with best balance
            const perms = permutations(availableWorkers);
            let bestPerm = perms[0];
            let bestScore = Infinity;

            for (const perm of perms) {
                // Score: sum of (count + 1) for the assigned shift per worker → lower is more balanced
                let score = 0;
                for (let s = 0; s < 3; s++) {
                    score += shiftCounts[perm[s]][s];
                }
                // Tiebreak: prefer variety (penalize if worker always gets same shift)
                for (let s = 0; s < 3; s++) {
                    const maxCount = Math.max(...shiftCounts[perm[s]]);
                    const minCount = Math.min(...shiftCounts[perm[s]]);
                    score += (maxCount - minCount) * 0.1;
                }
                if (score < bestScore) {
                    bestScore = score;
                    bestPerm = perm;
                }
            }

            for (let s = 0; s < 3; s++) {
                shifts[bestPerm[s]] = SHIFTS[s].id;
                shiftCounts[bestPerm[s]][s]++;
            }

            schedule[day.date] = {
                off: offWorker,
                shifts: shifts,
            };
        }
    }

    return schedule;
}

// Generate all permutations of an array (for 3 elements = 6 permutations)
function permutations(arr) {
    if (arr.length <= 1) return [arr];
    const result = [];
    for (let i = 0; i < arr.length; i++) {
        const rest = [...arr.slice(0, i), ...arr.slice(i + 1)];
        for (const perm of permutations(rest)) {
            result.push([arr[i], ...perm]);
        }
    }
    return result;
}

// ===== Render Calendar =====
function renderCalendar(year, month, schedule) {
    calendarGrid.innerHTML = '';

    // Day headers
    DAY_NAMES.forEach((name, idx) => {
        const header = document.createElement('div');
        header.className = 'day-header' + (idx === 0 ? ' sunday' : '') + (idx === 6 ? ' saturday' : '');
        header.textContent = name;
        calendarGrid.appendChild(header);
    });

    const daysInMonth = new Date(year, month, 0).getDate();
    const firstDay = new Date(year, month - 1, 1).getDay();

    // Today
    const today = new Date();
    const isCurrentMonth = today.getFullYear() === year && today.getMonth() + 1 === month;

    // Empty cells
    for (let i = 0; i < firstDay; i++) {
        const empty = document.createElement('div');
        empty.className = 'calendar-cell empty';
        calendarGrid.appendChild(empty);
    }

    // Day cells
    for (let d = 1; d <= daysInMonth; d++) {
        const dow = new Date(year, month - 1, d).getDay();
        const cell = document.createElement('div');
        let classes = 'calendar-cell';
        if (dow === 0) classes += ' sunday';
        if (dow === 6) classes += ' saturday';
        if (isCurrentMonth && today.getDate() === d) classes += ' today';
        cell.className = classes;

        // Date number
        const dateNum = document.createElement('div');
        dateNum.className = 'date-number';
        dateNum.textContent = d;
        cell.appendChild(dateNum);

        // Schedule entries
        const daySchedule = schedule[d];
        if (daySchedule) {
            const entries = [];
            for (let w = 0; w < 4; w++) {
                if (w === daySchedule.off) {
                    entries.push({ worker: w, shift: null, sortKey: 999 });
                } else {
                    entries.push({ worker: w, shift: daySchedule.shifts[w], sortKey: daySchedule.shifts[w] });
                }
            }
            entries.sort((a, b) => a.sortKey - b.sortKey);

            for (const entry of entries) {
                const div = document.createElement('div');
                if (entry.shift === null) {
                    div.className = 'shift-entry shift-off';
                    div.innerHTML = `<span class="shift-time">휴무</span><span class="shift-name">${state.names[entry.worker]}</span>`;
                } else {
                    const shiftInfo = SHIFTS.find(s => s.id === entry.shift);
                    div.className = `shift-entry shift-${entry.shift}`;
                    div.innerHTML = `<span class="shift-time">${shiftInfo.label}</span><span class="shift-name">${state.names[entry.worker]}</span>`;
                }
                cell.appendChild(div);
            }
        }

        calendarGrid.appendChild(cell);
    }
}

// ===== Render Summary =====
function renderSummary(year, month, schedule) {
    summaryGrid.innerHTML = '';

    const daysInMonth = new Date(year, month, 0).getDate();

    // Build day info for week counting
    const days = [];
    for (let d = 1; d <= daysInMonth; d++) {
        days.push(new Date(year, month - 1, d).getDay());
    }

    // Count weeks
    let weekCount = 0;
    for (const dow of days) {
        if (dow === 0) weekCount++;
    }
    if (days[0] !== 0) weekCount++; // partial first week

    // Compute stats per worker
    const stats = state.names.map((name, idx) => {
        let workDays = 0;
        let offDays = 0;
        let shift7 = 0, shift9 = 0, shift13 = 0;

        // Track consecutive off-day streaks
        let currentStreak = 0;
        let maxStreak = 0;
        let totalStreaks = 0; // number of off-day groups
        let consecutivePairs = 0;

        for (let d = 1; d <= daysInMonth; d++) {
            const ds = schedule[d];
            if (!ds) continue;
            if (ds.off === idx) {
                offDays++;
                currentStreak++;
                if (currentStreak > maxStreak) maxStreak = currentStreak;
            } else {
                if (currentStreak >= 2) consecutivePairs++;
                if (currentStreak > 0) totalStreaks++;
                currentStreak = 0;
                workDays++;
                const shiftId = ds.shifts[idx];
                if (shiftId === 7) shift7++;
                else if (shiftId === 9) shift9++;
                else if (shiftId === 13) shift13++;
            }
        }
        if (currentStreak >= 2) consecutivePairs++;
        if (currentStreak > 0) totalStreaks++;

        return { name, workDays, offDays, shift7, shift9, shift13, hours: workDays * 8, maxStreak, consecutivePairs };
    });

    for (const s of stats) {
        const total = s.shift7 + s.shift9 + s.shift13;
        const pct7 = total > 0 ? (s.shift7 / total * 100) : 0;
        const pct9 = total > 0 ? (s.shift9 / total * 100) : 0;
        const pct13 = total > 0 ? (s.shift13 / total * 100) : 0;

        const weeklyAvg = (s.hours / weekCount).toFixed(0);

        const item = document.createElement('div');
        item.className = 'summary-item';
        item.innerHTML = `
            <div class="summary-name">${s.name}</div>
            <div class="summary-stats">
                <div class="summary-stat">
                    <span>근무일</span>
                    <span class="summary-stat-value stat-work">${s.workDays}일</span>
                </div>
                <div class="summary-stat">
                    <span>휴무일</span>
                    <span class="summary-stat-value stat-off">${s.offDays}일</span>
                </div>
                <div class="summary-stat">
                    <span>총 근무시간</span>
                    <span class="summary-stat-value stat-hours">${s.hours}시간</span>
                </div>
                <div class="summary-stat">
                    <span>주당 평균</span>
                    <span class="summary-stat-value stat-hours">~${weeklyAvg}시간</span>
                </div>
                <div class="summary-stat">
                    <span>연속 휴무 최대</span>
                    <span class="summary-stat-value" style="color: var(--accent-purple)">${s.maxStreak}일</span>
                </div>
                <div class="summary-stat">
                    <span>2일+ 연속 휴무</span>
                    <span class="summary-stat-value" style="color: var(--accent-purple)">${s.consecutivePairs}회</span>
                </div>
                <div class="summary-stat">
                    <span>7시 / 9시 / 1시</span>
                    <span class="summary-stat-value">${s.shift7} / ${s.shift9} / ${s.shift13}</span>
                </div>
            </div>
            <div class="summary-bar">
                <div class="bar-segment-7" style="width:${pct7}%"></div>
                <div class="bar-segment-9" style="width:${pct9}%"></div>
                <div class="bar-segment-13" style="width:${pct13}%"></div>
            </div>
        `;
        summaryGrid.appendChild(item);
    }
}

// ===== Generate =====
function generate() {
    readNames();

    const schedule = generateSchedule(state.year, state.month);
    state.schedule = schedule;

    legendSection.style.display = 'block';
    calendarSection.style.display = 'block';
    tableSection.style.display = 'block';
    summarySection.style.display = 'block';

    calendarTitle.textContent = `${state.year}년 ${state.month}월 근무 일정표`;

    renderCalendar(state.year, state.month, schedule);
    renderTable(state.year, state.month, schedule);
    renderSummary(state.year, state.month, schedule);

    calendarSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== Render Table (Daily List) =====
function renderTable(year, month, schedule) {
    const daysInMonth = new Date(year, month, 0).getDate();

    let html = `
        <thead>
            <tr>
                <th rowspan="2">일자</th>
                <th rowspan="2">요일</th>
                <th colspan="1" class="shift-header-a">A조 (07:00~15:00)</th>
                <th colspan="1" class="shift-header-b">B조 (09:00~17:00)</th>
                <th colspan="1" class="shift-header-c">C조 (13:00~21:00)</th>
                <th rowspan="2" class="shift-header-off">휴무</th>
            </tr>
            <tr>
                <th class="shift-header-a">근무자</th>
                <th class="shift-header-b">근무자</th>
                <th class="shift-header-c">근무자</th>
            </tr>
        </thead>
        <tbody>
    `;

    for (let d = 1; d <= daysInMonth; d++) {
        const dow = new Date(year, month - 1, d).getDay();
        const ds = schedule[d];
        let w7 = '', w9 = '', w13 = '', wOff = '';

        for (let w = 0; w < 4; w++) {
            if (w === ds.off) {
                wOff = state.names[w];
            } else {
                const shiftId = ds.shifts[w];
                if (shiftId === 7) w7 = state.names[w];
                else if (shiftId === 9) w9 = state.names[w];
                else if (shiftId === 13) w13 = state.names[w];
            }
        }

        const rowClass = dow === 0 ? 'row-sunday' : dow === 6 ? 'row-saturday' : '';
        const dayClass = dow === 0 ? 'sunday' : dow === 6 ? 'saturday' : '';

        html += `
            <tr class="${rowClass}">
                <td class="td-date">${d}</td>
                <td class="td-day ${dayClass}">${DAY_NAMES[dow]}</td>
                <td class="td-shift-a">${w7}</td>
                <td class="td-shift-b">${w9}</td>
                <td class="td-shift-c">${w13}</td>
                <td class="td-off">${wOff}</td>
            </tr>
        `;
    }

    html += '</tbody>';
    scheduleTable.innerHTML = html;
}

// ===== Excel Export =====
function exportToExcel() {
    if (!state.schedule) {
        alert('먼저 일정을 생성해주세요.');
        return;
    }

    const year = state.year;
    const month = state.month;
    const schedule = state.schedule;
    const daysInMonth = new Date(year, month, 0).getDate();

    // ----- Sheet 1: 근무명령부 (vertical daily format) -----
    const rows = [];

    // Row 0: Title
    rows.push([`${year}년 ${month}월 근무 일정표`, '', '', '', '', '', `${month}월 1일 ~ ${month}월 ${daysInMonth}일`]);

    // Row 1: Empty spacer
    rows.push([]);

    // Row 2: Shift time headers
    rows.push([
        '', '',
        'A조 (07:00~15:00)', '',
        'B조 (09:00~17:00)', '',
        'C조 (13:00~21:00)', '',
        '휴무', '비고'
    ]);

    // Row 3: Sub-headers
    rows.push([
        '일자', '요일',
        '근무자', '서명',
        '근무자', '서명',
        '근무자', '서명',
        '', ''
    ]);

    // Row 4+: One row per day
    for (let d = 1; d <= daysInMonth; d++) {
        const dow = new Date(year, month - 1, d).getDay();
        const ds = schedule[d];
        let w7 = '', w9 = '', w13 = '', wOff = '';

        for (let w = 0; w < 4; w++) {
            if (w === ds.off) {
                wOff = state.names[w];
            } else {
                const shiftId = ds.shifts[w];
                if (shiftId === 7) w7 = state.names[w];
                else if (shiftId === 9) w9 = state.names[w];
                else if (shiftId === 13) w13 = state.names[w];
            }
        }

        rows.push([
            d,
            DAY_NAMES[dow],
            w7, '',       // A조 근무자, 서명
            w9, '',       // B조 근무자, 서명
            w13, '',      // C조 근무자, 서명
            wOff, ''      // 휴무, 비고
        ]);
    }

    // ----- Sheet 2: 근무 통계 -----
    const summaryRows = [];
    summaryRows.push(['직원', '근무일', '휴무일', '총 근무시간', '7시 출근', '9시 출근', '1시 출근']);

    for (let w = 0; w < 4; w++) {
        let workDays = 0, offDays = 0, s7 = 0, s9 = 0, s13 = 0;
        for (let d = 1; d <= daysInMonth; d++) {
            const ds = schedule[d];
            if (ds.off === w) {
                offDays++;
            } else {
                workDays++;
                const sid = ds.shifts[w];
                if (sid === 7) s7++;
                else if (sid === 9) s9++;
                else if (sid === 13) s13++;
            }
        }
        summaryRows.push([state.names[w], workDays, offDays, workDays * 8, s7, s9, s13]);
    }

    // ----- Build Workbook -----
    const wb = XLSX.utils.book_new();

    // Sheet 1: 근무명령부
    const ws1 = XLSX.utils.aoa_to_sheet(rows);

    // Column widths
    ws1['!cols'] = [
        { wch: 6 },   // 일자
        { wch: 5 },   // 요일
        { wch: 10 },  // A조 근무자
        { wch: 8 },   // A조 서명
        { wch: 10 },  // B조 근무자
        { wch: 8 },   // B조 서명
        { wch: 10 },  // C조 근무자
        { wch: 8 },   // C조 서명
        { wch: 10 },  // 휴무
        { wch: 12 },  // 비고
    ];

    // Merge cells for title and shift headers
    ws1['!merges'] = [
        // Title: merge A1:F1
        { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
        // Date range: merge G1:J1
        { s: { r: 0, c: 6 }, e: { r: 0, c: 9 } },
        // A조 header: merge C3:D3
        { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } },
        // B조 header: merge E3:F3
        { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } },
        // C조 header: merge G3:H3
        { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } },
        // 휴무 header: merge row 3-4 for I column
        { s: { r: 2, c: 8 }, e: { r: 3, c: 8 } },
        // 비고 header: merge row 3-4 for J column
        { s: { r: 2, c: 9 }, e: { r: 3, c: 9 } },
        // 일자 header: merge row 3-4
        { s: { r: 2, c: 0 }, e: { r: 3, c: 0 } },
        // 요일 header: merge row 3-4
        { s: { r: 2, c: 1 }, e: { r: 3, c: 1 } },
    ];

    XLSX.utils.book_append_sheet(wb, ws1, '근무명령부');

    // Sheet 2: Summary
    const ws2 = XLSX.utils.aoa_to_sheet(summaryRows);
    ws2['!cols'] = [
        { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    ];
    XLSX.utils.book_append_sheet(wb, ws2, '근무 통계');

    // Download using Blob (reliable for file:// protocol)
    const filename = `근무일정표_${year}년${month}월.xlsx`;
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }, 100);
}
