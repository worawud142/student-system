async function exportExcel() {
    if (!activeSubj || !students.length) return alert('ไม่พบข้อมูลนักเรียน');

    try {
        const btn = document.querySelector('button[onclick="exportExcel()"]');
        const oldText = btn ? btn.innerText : '📥 Export Excel';
        if (btn) btn.innerText = '⏳ กำลังบันทึกข้อมูล...';

        const pendingIds = typeof pendingScores !== 'undefined'
            ? Object.keys(pendingScores || {}) : [];
        if (pendingIds.length > 0)
            await Promise.all(pendingIds.map(id => saveScores(id)));
        await loadData();

        if (btn) btn.innerText = '⏳ กำลังสร้างไฟล์...';

        const binaryString = atob(TEMPLATE_B64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) bytes[i] = binaryString.charCodeAt(i);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(bytes.buffer);

        // ── Helpers ─────────────────────────────────────────────
        function breakFormula(cell) {
            // ทำลาย shared/cross formula โดยไม่ทำให้ _value เป็น null
            if (cell._value && typeof cell._value === 'object') {
                const model = cell._value.model || {};
                delete model.formula;
                delete model.sharedFormula;
                delete model.result;
                delete model.shareType;
                delete model.ref;
                delete model.si;
            }
            // ล้าง model ระดับ cell ด้วย
            if (cell.model) {
                delete cell.model.formula;
                delete cell.model.sharedFormula;
                delete cell.model.result;
            }
        }
        function safeSet(ws, ref, val) {
            const cell = typeof ref === 'string' ? ws.getCell(ref) : ws.getCell(ref[0], ref[1]);
            breakFormula(cell);
            cell.value = val;
        }
        function safeSetRC(ws, r, c, val) {
            const cell = ws.getCell(r, c);
            breakFormula(cell);
            cell.value = val;
        }
        function setFormulaCell(ws, ref, formula, result = null) {
            const cell = typeof ref === 'string' ? ws.getCell(ref) : ws.getCell(ref[0], ref[1]);
            breakFormula(cell);
            cell.value = { formula, result };
        }
        function clearRect(ws, r1, r2, c1, c2) {
            for (let r = r1; r <= r2; r++)
                for (let c = c1; c <= c2; c++) safeSetRC(ws, r, c, null);
        }
        function getSData(s) {
            const rec = s.scores.find(x => x.subject_name === activeSubj.subject_name) || {};
            let d = rec.score_data || {};
            if (typeof d === 'string') try { d = JSON.parse(d); } catch { d = {}; }
            if (typeof pendingScores !== 'undefined' && pendingScores[s.id])
                d = Object.assign({}, d, pendingScores[s.id]);
            return d;
        }
        function modeScore(values) {
            if (!values || values.length === 0) return null;
            const freq = new Map();
            values.forEach((v) => freq.set(v, (freq.get(v) || 0) + 1));
            let bestVal = null;
            let bestCount = -1;
            for (const [val, count] of freq.entries()) {
                if (count > bestCount || (count === bestCount && val > bestVal)) {
                    bestVal = val;
                    bestCount = count;
                }
            }
            return bestVal;
        }

        const n    = students.length;
        const conf = typeof coverConfig !== 'undefined'
            ? coverConfig
            : { term: '2', year: '2568', teacher: '', advisors: [], director: '' };
        const hpw  = conf.hoursPerWeek || 2;

        const MONTHS_TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                           'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

        // ── 1. ปก (1) ────────────────────────────────────────────
        const ws1 = wb.getWorksheet('ปก (1)');
        if (ws1) {
            safeSet(ws1, 'G9',  activeSubj.grade_level || '');
            safeSet(ws1, 'J9',  `ภาคเรียนที่  ${conf.term}`);
            safeSet(ws1, 'O9',  parseInt(conf.year) || conf.year);
            safeSet(ws1, 'E10', activeSubj.subject_name || '');
            safeSet(ws1, 'L10', activeSubj.subject_code || '');
            safeSet(ws1, 'E11', hpw);
            safeSet(ws1, 'N11', conf.credits || 1.5);
            if (conf.teacher)   safeSet(ws1, 'E12', conf.teacher);
            if (conf.advisors && conf.advisors.length > 0) {
                safeSet(ws1, 'E13', conf.advisors[0] || '');
                safeSet(ws1, 'E14', conf.advisors[1] || '');
            }
            safeSet(ws1, 'E34', '☐');
            safeSet(ws1, 'I34', '☐');
        }

        // ── วันเปิดภาค → หา Sunday ต้นสัปดาห์ ───────────────────
        // template: col G = อาทิตย์, col M = เสาร์ (7 cols/week)
        const semStartStr = conf.semesterStart || '2025-11-03';
        const semStartDate = new Date(semStartStr + 'T00:00:00');
        const weekStart = new Date(semStartDate);
        weekStart.setDate(semStartDate.getDate() - semStartDate.getDay()); // ถอยไปวันอาทิตย์

        function makeWeekDates(startSunday, numWeeks) {
            const dates = [];
            for (let w = 0; w < numWeeks; w++) {
                for (let d = 0; d < 7; d++) {
                    const dt = new Date(startSunday);
                    dt.setDate(dt.getDate() + w * 7 + d);
                    dates.push(
                        `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`
                    );
                }
            }
            return dates;
        }

        // ชีต 2: สัปดาห์ 1-7 (49 วัน), ชีต 3: 8-14, ชีต 4: 15-20
        const datesW2 = makeWeekDates(weekStart, 7);
        const sun8  = new Date(weekStart); sun8.setDate(weekStart.getDate() + 49);
        const datesW3 = makeWeekDates(sun8, 7);
        const sun15 = new Date(weekStart); sun15.setDate(weekStart.getDate() + 98);
        const datesW4 = makeWeekDates(sun15, 6);
        const allSemesterDates = datesW2.concat(datesW3, datesW4);
        const lastGradeRow = 6 + n;
        const lastTraitRow = 4 + n;

        // วันที่สอนจริง (จากข้อมูลเช็คชื่อ)
        const daySet = new Set();
        students.forEach(s => {
            if (s.attendance)
                s.attendance
                    .filter(a => a.subject_name === activeSubj.subject_name)
                    .forEach(a => daySet.add(new Date(a.check_date + 'T00:00:00').getDay()));
        });
        const meetingDays = [...daySet].sort();
        if (meetingDays.length === 0) meetingDays.push(1, 2, 3, 4, 5);

        // ── 2 & 3. เวลาเรียน (2) และ (3) ─────────────────────
        // col G(7)–BC(55) = 49 col, 7 cols/week [อา,จ,อ,พ,พฤ,ศ,ส]
        // สรุปไม่มี (template ดึงจาก ชีต 4 ผ่าน cross-sheet formula)
        // mark: มา→"", ขาด→"ข", ลา→"ล", สาย→"ป", ลากิจ→"ด"
        function fillAttSheet23(wsName, datesSpan) {
            const ws = wb.getWorksheet(wsName);
            if (!ws) return;

            // header วัน/เดือน
            for (let c = 0; c < 49; c++) {
                const col = 7 + c;
                const dStr = datesSpan[c];
                if (dStr) {
                    const dObj = new Date(dStr + 'T00:00:00');
                    if (c % 7 === 0) safeSetRC(ws, 2, col, MONTHS_TH[dObj.getMonth()]);
                    safeSetRC(ws, 4, col, dObj.getDate());
                } else {
                    if (c % 7 === 0) safeSetRC(ws, 2, col, '');
                    safeSetRC(ws, 4, col, '');
                }
            }

            // ข้อมูลนักเรียน
            students.forEach((s, i) => {
                const r = 6 + i;
                safeSetRC(ws, r, 1, s.roll_number || (i + 1));
                safeSetRC(ws, r, 2, s.std_id || '');
                safeSetRC(ws, r, 3, s.name || '');

                const attMap = {};
                if (s.attendance)
                    s.attendance
                        .filter(a => a.subject_name === activeSubj.subject_name)
                        .forEach(a => { attMap[a.check_date] = a.status; });

                for (let c = 0; c < 49; c++) {
                    const dayInWeek = c % 7; // 0=อาทิตย์, 6=เสาร์
                    const dStr = datesSpan[c];
                    let mark = '';
                    if (dStr && meetingDays.includes(dayInWeek)) {
                        const st = attMap[dStr] || '';
                        if      (st === 'P') mark = '';
                        else if (st === 'A') mark = 'ข';
                        else if (st === 'L') mark = 'ล';
                        else if (st === 'T') mark = 'ป';
                        else if (st === 'D') mark = 'ด';
                    }
                    safeSetRC(ws, r, 7 + c, mark);
                }
            });

            // ล้างแถวเกิน
            for (let i = n; i < 50; i++) {
                const r = 6 + i;
                for (let c = 1; c <= 55; c++) safeSetRC(ws, r, c, null);
            }
        }

        // ── 4. เวลาเรียน (4) ─────────────────────────────────
        // col H(8)–AW(49) = 42 col, 7 cols/week [อา,จ,อ,พ,พฤ,ศ,ส]
        // col AX(50) = สรุปเวลาเรียน (total - ขาด/ลา * ชั่วโมงต่อครั้ง)
        function fillAttSheet4(wsName, datesSpan) {
            const ws = wb.getWorksheet(wsName);
            if (!ws) return;

            for (let c = 0; c < 42; c++) {
                const col = 8 + c;
                const dStr = datesSpan[c];
                if (dStr) {
                    const dObj = new Date(dStr + 'T00:00:00');
                    if (c % 7 === 0) safeSetRC(ws, 2, col, MONTHS_TH[dObj.getMonth()]);
                    safeSetRC(ws, 4, col, dObj.getDate());
                } else {
                    if (c % 7 === 0) safeSetRC(ws, 2, col, '');
                    safeSetRC(ws, 4, col, '');
                }
            }

            students.forEach((s, i) => {
                const r = 6 + i;
                const attMap = {};
                if (s.attendance)
                    s.attendance
                        .filter(a => a.subject_name === activeSubj.subject_name)
                        .forEach(a => { attMap[a.check_date] = a.status; });

                for (let c = 0; c < 42; c++) {
                    const dayInWeek = c % 7;
                    const dStr = datesSpan[c];
                    let mark = '';
                    if (dStr && meetingDays.includes(dayInWeek)) {
                        const st = attMap[dStr] || '';
                        if      (st === 'P') mark = '';
                        else if (st === 'A') mark = 'ข';
                        else if (st === 'L') mark = 'ล';
                        else if (st === 'T') mark = 'ป';
                        else if (st === 'D') mark = 'ด';
                    }
                    safeSetRC(ws, r, 8 + c, mark);
                }

                let kh = 0, la = 0, pa = 0, da = 0;
                allSemesterDates.forEach((date) => {
                    if (!date) return;
                    const dayInWeek = new Date(date + 'T00:00:00').getDay();
                    if (!meetingDays.includes(dayInWeek)) return;
                    const st = attMap[date] || '';
                    if (st === 'A') kh += hpw;
                    else if (st === 'L') la += hpw;
                    else if (st === 'T') pa += hpw;
                    else if (st === 'D') da += hpw;
                });

                // เวลาเรียน = เวลาเต็ม - เวลาที่ขาด/ลา
                const totalSlots = allSemesterDates.reduce((sum, date) => {
                    if (!date) return sum;
                    const dayInWeek = new Date(date + 'T00:00:00').getDay();
                    return meetingDays.includes(dayInWeek) ? sum + hpw : sum;
                }, 0);
                const totalAtt = totalSlots - kh - la;
                safeSetRC(ws, r, 50, totalAtt);
            });

            for (let i = n; i < 35; i++) {
                const r = 6 + i;
                for (let c = 1; c <= 50; c++) safeSetRC(ws, r, c, null);
            }
        }

        fillAttSheet23('เวลาเรียน (2)', datesW2);
        fillAttSheet23('เวลาเรียน (3)', datesW3);
        fillAttSheet4('เวลาเรียน (4)', datesW4);

        // ── Score categories ─────────────────────────────────────
        const uc12 = new Array(12).fill(null);
        scoreConfig.forEach(c => {
            if (c.type && c.type.startsWith('u')) {
                const idx = parseInt(c.type.substring(1)) - 1;
                if (idx >= 0 && idx < 12) uc12[idx] = c;
            }
        });
        let midCat   = scoreConfig.find(c => c.type === 'mid'   || c.name === 'กลางภาค') || null;
        let finalCat = scoreConfig.find(c => c.type === 'final' || c.name === 'ปลายภาค') || null;
        scoreConfig
            .filter(c => (c.type === 'none' || !c.type) && c !== midCat && c !== finalCat)
            .forEach(c => {
                let idx = 0;
                while (uc12[idx] !== null) idx++;
                if (idx < 12) uc12[idx] = c;
            });

        function distributeSubScores(cat, totalScore) {
            const total = parseFloat(totalScore) || 0;
            const blank = Array(5).fill('');
            if (!cat || total <= 0) return blank;
            const subs = (cat.subScores || '').split(',')
                .map(s => parseFloat(String(s).trim()) || 0).filter(n => n > 0).slice(0, 5);
            if (!subs.length) { blank[0] = total; return blank; }
            const maxSum = subs.reduce((a, v) => a + v, 0);
            if (!maxSum) { blank[0] = total; return blank; }
            const ints = subs.map(v => Math.floor((v / maxSum) * total));
            let diff = Math.round(total - ints.reduce((a, v) => a + v, 0));
            subs.map((v, i2) => ({ i2, frac: ((v / maxSum) * total) - ints[i2] }))
                .sort((a, b) => b.frac - a.frac)
                .forEach((item, i2) => { if (i2 < diff) ints[item.i2]++; });
            ints.forEach((v, i2) => { blank[i2] = v; });
            return blank;
        }

        // ── 5,6,7. ชีตหน่วย ─────────────────────────────────
        const UNIT_COLS = [
            { start: 3,  sum: 8  },
            { start: 10, sum: 15 },
            { start: 17, sum: 22 },
            { start: 24, sum: 29 },
        ];

        const unitScores = students.map(() => new Array(12).fill(0));

        function fillUnitSheet(wsName, unitIndices) {
            const ws = wb.getWorksheet(wsName);
            if (!ws) return;
            clearRect(ws, 5, 55, 1, 30);

            unitIndices.forEach((uIdx, slot) => {
                const cat = uc12[uIdx];
                safeSetRC(ws, 5, UNIT_COLS[slot].start, cat ? parseInt(cat.maxScore) : 0);
            });

            students.forEach((s, i) => {
                const r = 6 + i;
                const d = getSData(s);
                safeSetRC(ws, r, 1, s.roll_number || (i + 1));
                safeSetRC(ws, r, 2, s.name || '');

                unitIndices.forEach((uIdx, slot) => {
                    const cat   = uc12[uIdx];
                    const score = cat ? (parseFloat(d[cat.id]) || 0) : 0;
                    const dist  = distributeSubScores(cat, score);
                    const uc    = UNIT_COLS[slot];
                    for (let k = 0; k < 5; k++) safeSetRC(ws, r, uc.start + k, dist[k]);
                    safeSetRC(ws, r, uc.sum, score);
                    safeSetRC(ws, r, uc.sum + 1, null);
                    unitScores[i][uIdx] = score;
                });
            });
        }

        fillUnitSheet('หน่วย 1,4 (5)',  [0, 1, 2, 3]);
        fillUnitSheet('หน่วย 5,8 (6)',  [4, 5, 6, 7]);
        fillUnitSheet('หน่วย 9,12 (7)', [8, 9, 10, 11]);

        // ── 8. สรุปผลรวม (8) ─────────────────────────────────
        const gradeValues = [];
        const gradeCounts = { '4': 0, '3.5': 0, '3': 0, '2.5': 0, '2': 0, '1.5': 0, '1': 0, '0': 0 };
        const traitCounts = { '3': 0, '2': 0, '1': 0, '0': 0 };
        const readCounts = { '3': 0, '2': 0, '1': 0, '0': 0 };
        const compCounts = { '3': 0, '2': 0, '1': 0, '0': 0 };
        const ws8 = wb.getWorksheet('สรุปผลรวม (8)');
        if (ws8) {
            clearRect(ws8, 6, 56, 1, 19);

            const midMax   = midCat   ? parseInt(midCat.maxScore)   : 0;
            const finalMax = finalCat ? parseInt(finalCat.maxScore) : 0;
            let totalUnitMax = 0;
            uc12.forEach(c => { if (c) totalUnitMax += parseInt(c.maxScore); });
            const totalMaxOut = totalUnitMax + midMax + finalMax;

            for (let j = 0; j < 12; j++)
                safeSetRC(ws8, 6, 3 + j, uc12[j] ? parseInt(uc12[j].maxScore) : 0);
            safeSetRC(ws8, 6, 15, midMax);
            safeSetRC(ws8, 6, 16, finalMax);
            safeSetRC(ws8, 6, 17, totalMaxOut);

            const calcGrade = (t, m) => {
                if (!m) return 0;
                const p = (t / m) * 100;
                if (p >= 80) return 4; if (p >= 75) return 3.5; if (p >= 70) return 3;
                if (p >= 65) return 2.5; if (p >= 60) return 2; if (p >= 55) return 1.5;
                if (p >= 50) return 1; return 0;
            };

            students.forEach((s, i) => {
                const r = 7 + i;
                const d = getSData(s);
                safeSetRC(ws8, r, 1, s.roll_number || (i + 1));
                safeSetRC(ws8, r, 2, s.name || '');

                let unitTotal = 0;
                for (let j = 0; j < 12; j++) {
                    safeSetRC(ws8, r, 3 + j, unitScores[i][j]);
                    unitTotal += unitScores[i][j];
                }

                const midV = midCat   ? (parseFloat(d[midCat.id])   || 0) : 0;
                const finV = finalCat ? (parseFloat(d[finalCat.id]) || 0) : 0;
                safeSetRC(ws8, r, 15, midV);
                safeSetRC(ws8, r, 16, finV);
                const total = unitTotal + midV + finV;
                safeSetRC(ws8, r, 17, total);
                const grade = calcGrade(total, totalMaxOut);
                safeSetRC(ws8, r, 18, grade);
                gradeValues.push(grade);
                if (gradeCounts[String(grade)] !== undefined) gradeCounts[String(grade)]++;
            });
        }

        // ── 9. คุณลักษณะ อ่าน สมรรถนะ (9) ──────────────────
        // template row5+ มี shared/cross formula — ใช้ cell._value = null ก่อนเสมอ
        const ws9 = wb.getWorksheet('คุณลักษณะ อ่าน สมรรถนะ (9)');
        if (ws9) {
            const CHAR_TRAITS = ['t1','t2','t3','t4','t5','t6','t7','t8'];
            const READ_SKILLS = ['r1','r2','r3','r4','r5'];
            const COMP_SKILLS = ['c1','c2','c3','c4','c5'];

            // ล้างทุก cell ในช่วงข้อมูลก่อน (รวมถึง shared formula chain)
            for (let r = 5; r <= 5 + n + 5; r++)
                for (let c = 1; c <= 27; c++) safeSetRC(ws9, r, c, null);

            students.forEach((s, i) => {
                const r = 5 + i;
                const d = getSData(s);
                safeSetRC(ws9, r, 1, s.roll_number || (i + 1));
                safeSetRC(ws9, r, 2, s.name || '');

                // คุณลักษณะ 8 ด้าน: col C-J (3-10), K=สรุป (11)
                const traitVals = [];
                CHAR_TRAITS.forEach((key, j) => {
                    const val = (d[key] !== undefined && d[key] !== '' && d[key] !== null)
                        ? parseFloat(d[key]) : null;
                    safeSetRC(ws9, r, 3 + j, val);
                    if (val !== null) traitVals.push(val);
                });
                const traitSummary = modeScore(traitVals);
                safeSetRC(ws9, r, 11, traitSummary);
                if (traitSummary !== null && traitCounts[String(traitSummary)] !== undefined) {
                    traitCounts[String(traitSummary)]++;
                }

                // การอ่าน 5 ด้าน: col M-Q (13-17), R=สรุป (18)
                const readVals = [];
                READ_SKILLS.forEach((key, j) => {
                    const val = (d[key] !== undefined && d[key] !== '' && d[key] !== null)
                        ? parseFloat(d[key]) : null;
                    safeSetRC(ws9, r, 13 + j, val);
                    if (val !== null) readVals.push(val);
                });
                const readSummary = modeScore(readVals);
                safeSetRC(ws9, r, 18, readSummary);
                if (readSummary !== null && readCounts[String(readSummary)] !== undefined) {
                    readCounts[String(readSummary)]++;
                }

                // สมรรถนะ 5 ด้าน: col T-X (20-24), Y=สรุป (25)
                const compVals = [];
                COMP_SKILLS.forEach((key, j) => {
                    const val = (d[key] !== undefined && d[key] !== '' && d[key] !== null)
                        ? parseFloat(d[key]) : null;
                    safeSetRC(ws9, r, 20 + j, val);
                    if (val !== null) compVals.push(val);
                });
                const compSummary = modeScore(compVals);
                safeSetRC(ws9, r, 25, compSummary);
                if (compSummary !== null && compCounts[String(compSummary)] !== undefined) {
                    compCounts[String(compSummary)]++;
                }
            });
        }

        // ── x-bar / SD → ปก (1) ──────────────────────────────
        // เก็บสูตรเดิมไว้ให้ Excel คำนวณเอง และให้ workbook บังคับ recalculation ตอนเปิดไฟล์
        const gradeRange = `'สรุปผลรวม (8)'!$R$7:$R$${lastGradeRow}`;
        const traitRange = `'คุณลักษณะ อ่าน สมรรถนะ (9)'!$K$5:$K$${lastTraitRow}`;
        const readRange  = `'คุณลักษณะ อ่าน สมรรถนะ (9)'!$R$5:$R$${lastTraitRow}`;
        const compRange  = `'คุณลักษณะ อ่าน สมรรถนะ (9)'!$Y$5:$Y$${lastTraitRow}`;

        if (ws1) {
            setFormulaCell(ws1, 'A18', `SUM(C18:M18)`, n);
            setFormulaCell(ws1, 'C18', `COUNTIF(${gradeRange},'ปก (1)'!C17)`, gradeCounts['4']);
            setFormulaCell(ws1, 'D18', `COUNTIF(${gradeRange},'ปก (1)'!D17)`, gradeCounts['3.5']);
            setFormulaCell(ws1, 'E18', `COUNTIF(${gradeRange},'ปก (1)'!E17)`, gradeCounts['3']);
            setFormulaCell(ws1, 'F18', `COUNTIF(${gradeRange},'ปก (1)'!F17)`, gradeCounts['2.5']);
            setFormulaCell(ws1, 'G18', `COUNTIF(${gradeRange},'ปก (1)'!G17)`, gradeCounts['2']);
            setFormulaCell(ws1, 'H18', `COUNTIF(${gradeRange},'ปก (1)'!H17)`, gradeCounts['1.5']);
            setFormulaCell(ws1, 'I18', `COUNTIF(${gradeRange},'ปก (1)'!I17)`, gradeCounts['1']);
            setFormulaCell(ws1, 'J18', `COUNTIF(${gradeRange},'ปก (1)'!J17)`, gradeCounts['0']);
            setFormulaCell(ws1, 'K18', `COUNTIF(${gradeRange},'ปก (1)'!K17)`, 0);
            setFormulaCell(ws1, 'L18', `COUNTIF(${gradeRange},'ปก (1)'!L17)`, 0);
            setFormulaCell(ws1, 'M18', `COUNTIF(${gradeRange},'ปก (1)'!M17)`, 0);

            if (gradeValues.length > 0) {
                const xbar = gradeValues.reduce((a, v) => a + v, 0) / gradeValues.length;
                const sd = Math.sqrt(gradeValues.reduce((a, v) => a + (v - xbar) ** 2, 0) / gradeValues.length);
                setFormulaCell(ws1, 'N18', `AVERAGE(${gradeRange})`, Math.round(xbar * 100) / 100);
                setFormulaCell(ws1, 'P18', `STDEV(${gradeRange})`, Math.round(sd * 100) / 100);
            }

            setFormulaCell(ws1, 'A22', `A18`, n);
            setFormulaCell(ws1, 'D22', `COUNTIF(${traitRange},'ปก (1)'!D21)`, traitCounts['3']);
            setFormulaCell(ws1, 'E22', `COUNTIF(${traitRange},'ปก (1)'!E21)`, traitCounts['2']);
            setFormulaCell(ws1, 'F22', `COUNTIF(${traitRange},'ปก (1)'!F21)`, traitCounts['1']);
            setFormulaCell(ws1, 'G22', `COUNTIF(${traitRange},'ปก (1)'!G21)`, traitCounts['0']);
            setFormulaCell(ws1, 'H22', `COUNTIF(${readRange},'ปก (1)'!H21)`, readCounts['3']);
            setFormulaCell(ws1, 'I22', `COUNTIF(${readRange},'ปก (1)'!I21)`, readCounts['2']);
            setFormulaCell(ws1, 'J22', `COUNTIF(${readRange},'ปก (1)'!J21)`, readCounts['1']);
            setFormulaCell(ws1, 'K22', `COUNTIF(${readRange},'ปก (1)'!K21)`, readCounts['0']);
            setFormulaCell(ws1, 'L22', `COUNTIF(${compRange},'ปก (1)'!L21)`, compCounts['3']);
            setFormulaCell(ws1, 'M22', `COUNTIF(${compRange},'ปก (1)'!M21)`, compCounts['2']);
            setFormulaCell(ws1, 'N22', `COUNTIF(${compRange},'ปก (1)'!N21)`, compCounts['1']);
            setFormulaCell(ws1, 'O22', `COUNTIF(${compRange},'ปก (1)'!O21)`, compCounts['0']);
        }

        if (wb.calcProperties) {
            wb.calcProperties.fullCalcOnLoad = true;
            wb.calcProperties.forceFullCalc = true;
            wb.calcProperties.calcMode = 'auto';
        }

        // ── ผลการเรียน ────────────────────────────────────────
        const wsResult = wb.getWorksheet('ผลการเรียน');
        if (wsResult) {
            const boyCount  = students.filter(s => s.gender === 'M' || s.gender === 'ชาย').length;
            const girlCount = students.filter(s => s.gender === 'F' || s.gender === 'หญิง').length;
            safeSetRC(wsResult, 9, 11, boyCount  || 0);
            safeSetRC(wsResult, 9, 12, girlCount || 0);
            safeSetRC(wsResult, 9, 13, n);
        }

        // ── Export ────────────────────────────────────────────
        const levelObj = (activeSubj.grade_level || 'รายวิชา').replace(/\//g, '-').replace(/ /g, '_');
        const fname = `ปพ5_${activeSubj.subject_code || 'รายวิชา'}_${levelObj}.xlsx`;
        const outBuffer = await wb.xlsx.writeBuffer();
        saveAs(new Blob([outBuffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8'
        }), fname);

        if (btn) btn.innerText = oldText;
        alert(`✅ ส่งออกไฟล์ ${fname} สำเร็จ!`);

    } catch (e) {
        console.error(e);
        alert('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: ' + e.message);
        const btn = document.querySelector('button[onclick="exportExcel()"]');
        if (btn) btn.innerText = '📥 Export Excel';
    }
}
