async function exportExcel() {
    if (!activeSubj || !students.length) return alert('ไม่พบข้อมูลนักเรียน');

    try {
        const btn = document.querySelector('button[onclick="exportExcel()"]');
        const oldText = btn ? btn.innerText : '📥 Export Excel';
        if (btn) btn.innerText = '⏳ กำลังสร้างไฟล์...';

        const binaryString = atob(TEMPLATE_B64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        const buffer = bytes.buffer;

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);

        // Helper: safely set cell value by clearing any shared formula first
        function clearCellFormulas(cell) {
            // Clear all possible formula-related properties
            if (cell._value && cell._value.model) {
                delete cell._value.model.formula;
                delete cell._value.model.sharedFormula;
                delete cell._value.model.result;
                delete cell._value.model.shareType;
                delete cell._value.model.ref;
                delete cell._value.model.formulaType;
                delete cell._value.model.calculateCell;
                delete cell._value.model.si;
            }
            if (cell.model) {
                delete cell.model.formula;
                delete cell.model.sharedFormula;
                delete cell.model.result;
                delete cell.model.shareType;
                delete cell.model.ref;
                delete cell.model.formulaType;
                delete cell.model.calculateCell;
                delete cell.model.si;
            }
            // Clear the cell's internal formula reference
            if (cell._formula) {
                cell._formula = null;
            }
            if (cell._sharedFormula) {
                cell._sharedFormula = null;
            }
            // Reset the cell type to force recalculation
            cell.type = undefined;
        }

        // Important: do NOT clear formulas across the whole workbook.
        // We only clear formula metadata on cells that we explicitly overwrite via safeSet/safeSetRC.
        function safeSet(ws, cellRef, val) {
            let cell;
            if (typeof cellRef === 'string') {
                cell = ws.getCell(cellRef);
            } else if (typeof cellRef === 'number') {
                // called as safeSet(ws, row, col) — but that's 4 args; handle below
                cell = ws.getCell(cellRef);
            } else {
                cell = cellRef; // already a cell object
            }
            clearCellFormulas(cell);
            cell.value = null;
            cell.value = val;
        }
        // Overload: safeSetRC(ws, row, col, val)
        function safeSetRC(ws, row, col, val) {
            const cell = ws.getCell(row, col);
            clearCellFormulas(cell);
            cell.value = null;
            cell.value = val;
        }
        function clearRect(ws, rowStart, rowEnd, colStart, colEnd) {
            for (let r = rowStart; r <= rowEnd; r++) {
                for (let c = colStart; c <= colEnd; c++) {
                    safeSetRC(ws, r, c, null);
                }
            }
        }

        const n = students.length;

        // 1. ปก (1)
        const ws1 = wb.getWorksheet('ปก (1)');
        const conf = typeof coverConfig !== 'undefined' ? coverConfig : { term: '2', year: '2568', teacher: '', advisors: [], director: '' };
        if (ws1) {
            ws1.getCell('C8').value = `ระดับชั้น${activeSubj.grade_level} `;
            ws1.getCell('J8').value = `ภาคเรียนที่ ${conf.term}`;
            ws1.getCell('M8').value = `ปีการศึกษา    ${conf.year}`;
            ws1.getCell('E9').value = activeSubj.subject_name;
            ws1.getCell('L9').value = activeSubj.subject_code || '';
            ws1.getCell('A17').value = n;

            if (conf.teacher) ws1.getCell('F11').value = conf.teacher;
            if (conf.advisors && conf.advisors.length > 0) {
                ws1.getCell('E12').value = conf.advisors[0] || '';
                ws1.getCell('E13').value = conf.advisors[1] || '';
            }
            if (conf.director) {
                ws1.getCell('G36').value = `ลงชื่อ ${conf.director}`;

                // Add checkmark (✓) for approval status in E34 (Approved) and I34 (Disapproved)
                if (conf.approvalStatus === 'approved') {
                    ws1.getCell('E34').value = '✓';
                    ws1.getCell('I34').value = '';
                } else if (conf.approvalStatus === 'disapproved') {
                    ws1.getCell('E34').value = '';
                    ws1.getCell('I34').value = '✓';
                }
            }
        }

        // 2. เวลาเรียน (2) & 3. เวลาเรียน (3)
        // สร้างวันที่ 20 สัปดาห์ (100 วัน) เริ่มตั้งแต่ 3 พฤศจิกายน 2568
        const START_DATE = new Date('2025-11-03T00:00:00');
        const allDates = [];
        let currDate = new Date(START_DATE);
        while (allDates.length < 100) {
            const day = currDate.getDay();
            if (day >= 1 && day <= 5) { // จันทร์ - ศุกร์
                const y = currDate.getFullYear();
                const m = String(currDate.getMonth() + 1).padStart(2, '0');
                const d = String(currDate.getDate()).padStart(2, '0');
                allDates.push(`${y}-${m}-${d}`);
            }
            currDate.setDate(currDate.getDate() + 1);
        }

        const datesW1_10 = allDates.slice(0, 50);
        const datesW11_20 = allDates.slice(50, 100);

        const MONTHS_TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

        const hpw = conf.hoursPerWeek || 2;
        const dayOfWeekSet = new Set();
        students.forEach(s => {
            if (s.attendance) s.attendance.filter(a => a.subject_name === activeSubj.subject_name).forEach(a => dayOfWeekSet.add(new Date(a.check_date).getDay()));
        });
        const meetingDays = [...dayOfWeekSet];
        if (meetingDays.length === 0) meetingDays.push(1, 2, 3, 4, 5);

        let globalDaysCount = 0;
        allDates.forEach(d => { if (meetingDays.includes(new Date(d).getDay())) globalDaysCount++; });
        const globalTotalSlots = globalDaysCount * hpw;

        function fillAttendanceSheet(wsName, datesSpan) {
            const ws = wb.getWorksheet(wsName);
            if (!ws) return;

            let actualDaysCount = 0;
            datesSpan.forEach(d => {
                if (d && meetingDays.includes(new Date(d).getDay())) {
                    actualDaysCount++;
                }
            });
            const totalSlots = actualDaysCount * hpw;

            if (wsName === 'เวลาเรียน (3)') {
                ws.getCell('BJ4').value = globalTotalSlots;
            } else {
                ws.getCell('BJ4').value = totalSlots;
            }

            // Write date headers into the template (Row 2 = Month, Row 4 = Day)
            // Leave Row 3 (Week) and Row 5 (Hour) untouched to preserve Excel template formatting
            for (let c = 0; c < 50; c++) {
                let col = 8 + c; // Start at Column H (index 8)
                if (c < datesSpan.length) {
                    let dObj = new Date(datesSpan[c]);
                    // Only write Month on the first column of the week to avoid breaking merged cells
                    if (c % 5 === 0) {
                        safeSetRC(ws, 2, col, MONTHS_TH[dObj.getMonth()]);
                    }
                    safeSetRC(ws, 4, col, dObj.getDate());
                } else {
                    if (c % 5 === 0) safeSetRC(ws, 2, col, '');
                    safeSetRC(ws, 4, col, '');
                }
            }

            students.forEach((s, i) => {
                let r = 6 + i;
                ws.getCell(r, 1).value = s.roll_number || (i + 1); // A
                ws.getCell(r, 2).value = s.std_id || ''; // B
                ws.getCell(r, 3).value = s.name || '';   // C

                const attMap = {};
                if (s.attendance) {
                    s.attendance.filter(a => a.subject_name === activeSubj.subject_name).forEach(a => {
                        attMap[a.check_date] = a.status;
                    });
                }

                let kh = 0, la = 0, pa = 0, da = 0;
                datesSpan.forEach((date, dIdx) => {
                    let col = 8 + dIdx; // D is 4, H is 8. Dates start at H.
                    let st = attMap[date] || '';
                    let mark = '';
                    if (meetingDays.includes(new Date(date).getDay())) {
                        if (st === 'A') { mark = 'ข'; kh += hpw; }
                        else if (st === 'L') { mark = 'ล'; la += hpw; }
                        else if (st === 'T') { mark = 'ป'; pa += hpw; }
                        else if (st === 'D') { mark = 'ด'; da += hpw; }
                    }
                    safeSetRC(ws, r, col, mark);
                });

                // Clear remaining un-checked date columns
                for (let dIdx = datesSpan.length; dIdx < 50; dIdx++) {
                    safeSetRC(ws, r, 8 + dIdx, '');
                }

                // Write Excel calculations for summary retaining dynamic formula updates with hpw
                // Make formulas accumulate from previous sheet if it is the second sheet
                if (wsName === 'เวลาเรียน (3)') {
                    ws.getCell(r, 58).value = { formula: `COUNTIF(H${r}:BE${r},"ข")*${hpw}+'เวลาเรียน (2)'!BF${r}` };
                    ws.getCell(r, 59).value = { formula: `COUNTIF(H${r}:BE${r},"ล")*${hpw}+'เวลาเรียน (2)'!BG${r}` };
                    ws.getCell(r, 60).value = { formula: `COUNTIF(H${r}:BE${r},"ป")*${hpw}+'เวลาเรียน (2)'!BH${r}` };
                    ws.getCell(r, 61).value = { formula: `COUNTIF(H${r}:BE${r},"ด")*${hpw}+'เวลาเรียน (2)'!BI${r}` };
                    ws.getCell(r, 62).value = { formula: `BJ4-SUM(BF${r}:BI${r})` };
                } else {
                    ws.getCell(r, 58).value = { formula: `COUNTIF(H${r}:BE${r},"ข")*${hpw}`, result: kh };
                    ws.getCell(r, 59).value = { formula: `COUNTIF(H${r}:BE${r},"ล")*${hpw}`, result: la };
                    ws.getCell(r, 60).value = { formula: `COUNTIF(H${r}:BE${r},"ป")*${hpw}`, result: pa };
                    ws.getCell(r, 61).value = { formula: `COUNTIF(H${r}:BE${r},"ด")*${hpw}`, result: da };
                    ws.getCell(r, 62).value = { formula: `BJ4-SUM(BF${r}:BI${r})`, result: totalSlots - kh - la - pa - da };
                }
            });

            // clear remaining rows for student details and dates, including summary formulas
            for (let i = n; i < 50; i++) {
                let r = 6 + i;
                for (let c = 1; c <= 62; c++) safeSetRC(ws, r, c, null);
            }
        }

        fillAttendanceSheet('เวลาเรียน (2)', datesW1_10);
        fillAttendanceSheet('เวลาเรียน (3)', datesW11_20);

        function getSData(s) {
            const rec = s.scores.find(x => x.subject_name === activeSubj.subject_name) || {};
            let d = rec.score_data || {};
            if (typeof d === 'string') try { d = JSON.parse(d); } catch (e) { d = {}; }
            return d;
        }

        const uc8 = [];
        scoreConfig.forEach(c => {
            if (c.type && c.type.startsWith('u')) {
                const uIdx = parseInt(c.type.substring(1)) - 1;
                if (uIdx >= 0 && uIdx < 8) uc8[uIdx] = c;
            }
        });
        let midCat = scoreConfig.find(c => c.type === 'mid' || c.name === 'กลางภาค') || null;
        let finalCat = scoreConfig.find(c => c.type === 'final' || c.name === 'ปลายภาค') || null;

        scoreConfig.filter(c => (c.type === 'none' || !c.type) && c !== midCat && c !== finalCat).forEach(c => {
            let openIdx = 0;
            while (uc8[openIdx] !== undefined) openIdx++;
            uc8[openIdx] = c;
        });

        function distributeSubScores(cat, totalScore) {
            const total = parseFloat(totalScore) || 0;
            const blank = Array(7).fill('');
            if (!cat || total <= 0) return blank;

            const subs = (cat.subScores || '')
                .split(',')
                .map((s) => parseFloat(String(s).trim()) || 0)
                .filter((n) => n > 0)
                .slice(0, 7);

            if (!subs.length) {
                blank[0] = total;
                return blank;
            }

            const totalMaxSub = subs.reduce((sum, value) => sum + value, 0);
            if (!totalMaxSub) {
                blank[0] = total;
                return blank;
            }

            const ints = subs.map((maxVal) => Math.floor((maxVal / totalMaxSub) * total));
            let diff = Math.round(total - ints.reduce((sum, value) => sum + value, 0));
            const fracs = subs
                .map((maxVal, idx) => ({ idx, frac: ((maxVal / totalMaxSub) * total) - ints[idx] }))
                .sort((a, b) => b.frac - a.frac);

            for (let i = 0; i < diff && i < fracs.length; i++) {
                ints[fracs[i].idx]++;
            }

            ints.forEach((value, idx) => {
                blank[idx] = value;
            });
            return blank;
        }

        const setUnit = (wsName, key1Idx, key2Idx) => {
            const ws = wb.getWorksheet(wsName);
            if (!ws) return;
            const c1 = uc8[key1Idx];
            const c2 = uc8[key2Idx];

            // Clear student data block to remove any shared-formula clones left in template rows
            clearRect(ws, 6, 55, 1, 40); // A:AN, rows 6-55

            const max1 = c1 ? parseInt(c1.maxScore) : 0;
            const max2 = c2 ? parseInt(c2.maxScore) : 0;

            safeSet(ws, 'C5', max1);
            safeSet(ws, 'J5', max1);
            safeSet(ws, 'K5', max1 ? Math.round(max1 * 0.6) : 0);

            safeSet(ws, 'N5', max2);
            safeSet(ws, 'U5', max2);
            safeSet(ws, 'V5', max2 ? Math.round(max2 * 0.6) : 0);

            students.forEach((s, i) => {
                const d = getSData(s);
                const s1 = c1 ? (parseFloat(d[c1.id]) || 0) : 0;
                const s2 = c2 ? (parseFloat(d[c2.id]) || 0) : 0;
                const dist1 = distributeSubScores(c1, s1);
                const dist2 = distributeSubScores(c2, s2);
                let r = 6 + i;

                safeSet(ws, `A${r}`, s.roll_number || (i + 1));
                safeSet(ws, `B${r}`, s.name || '');
                ['C', 'D', 'E', 'F', 'G', 'H', 'I'].forEach((col, idx) => safeSet(ws, `${col}${r}`, dist1[idx]));
                safeSet(ws, `J${r}`, s1);
                safeSet(ws, `K${r}`, null);

                safeSet(ws, `L${r}`, s.roll_number || (i + 1));
                safeSet(ws, `M${r}`, s.name || '');
                ['N', 'O', 'P', 'Q', 'R', 'S', 'T'].forEach((col, idx) => safeSet(ws, `${col}${r}`, dist2[idx]));
                safeSet(ws, `U${r}`, s2);
                safeSet(ws, `V${r}`, null);
            });
        };

        setUnit('หน่วย 1,2 (4)', 0, 1);
        setUnit('หน่วย 3,4 (5)', 2, 3);
        setUnit('หน่วย 5,6 (6)', 4, 5);
        setUnit('หน่วย 7,8 (7)', 6, 7);

        // 8. สรุปผลรวม (8)
        const ws8 = wb.getWorksheet('สรุปผลรวม (8)');
        if (ws8) {
            // Clear student rows block first to avoid dangling shared-formula clones
            clearRect(ws8, 7, 56, 1, 40); // A:AN, rows 7-56

            const cols = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'];
            for (let i = 0; i < 8; i++) {
                safeSet(ws8, `${cols[i]}6`, uc8[i] ? parseInt(uc8[i].maxScore) : 0);
            }
            safeSet(ws8, 'L6', midCat ? parseInt(midCat.maxScore) : 0);
            safeSet(ws8, 'M6', finalCat ? parseInt(finalCat.maxScore) : 0);

            let totalUnitMax = 0;
            for (let i = 0; i < 8; i++) if (uc8[i]) totalUnitMax += parseInt(uc8[i].maxScore);
            safeSet(ws8, 'P6', totalUnitMax + (midCat ? parseInt(midCat.maxScore) : 0) + (finalCat ? parseInt(finalCat.maxScore) : 0));

            const totalMaxOut = totalUnitMax + (midCat ? parseInt(midCat.maxScore) : 0) + (finalCat ? parseInt(finalCat.maxScore) : 0);

            const calcGrade = (t, m) => {
                if (!m) return '0';
                const p = (t / m) * 100;
                if (p >= 80) return '4';
                if (p >= 75) return '3.5';
                if (p >= 70) return '3';
                if (p >= 65) return '2.5';
                if (p >= 60) return '2';
                if (p >= 55) return '1.5';
                if (p >= 50) return '1';
                return '0';
            };

            students.forEach((s, i) => {
                let r = 7 + i;
                safeSet(ws8, `A${r}`, s.roll_number || (i + 1));
                safeSet(ws8, `B${r}`, s.name || '');

                let total = 0;
                const d = getSData(s);
                const scoreCols = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'];

                for (let j = 0; j < 8; j++) {
                    const v = uc8[j] ? (parseFloat(d[uc8[j].id]) || 0) : 0;
                    safeSet(ws8, `${scoreCols[j]}${r}`, v);
                    total += v;
                }
                const midV = midCat ? (parseFloat(d[midCat.id]) || 0) : 0;
                const finV = finalCat ? (parseFloat(d[finalCat.id]) || 0) : 0;
                safeSet(ws8, `L${r}`, midV);
                safeSet(ws8, `M${r}`, finV);
                total += midV + finV;

                safeSet(ws8, `P${r}`, total);
                safeSet(ws8, `Q${r}`, calcGrade(total, totalMaxOut));
            });
        }

        // 9. คุณลักษณะ (9)
        const ws9 = wb.getWorksheet('คุณลักษณะ (9)');
        if (ws9) {
            // Clear student rows block first to avoid dangling shared-formula clones
            clearRect(ws9, 7, 56, 1, 40); // A:AN, rows 7-56

            const CHAR_TRAITS = ['t1', 't2', 't3', 't4', 't5', 't6', 't7', 't8'];
            students.forEach((s, i) => {
                const d = getSData(s);
                let b = [];
                for (let j = 0; j < 8; j++) b.push(d[CHAR_TRAITS[j]] !== undefined ? d[CHAR_TRAITS[j]] : 3);

                let r = 7 + i;
                ws9.getCell(`A${r}`).value = s.roll_number || (i + 1);
                ws9.getCell(`B${r}`).value = s.name || '';

                ['C', 'D', 'E', 'F'].forEach(c => safeSet(ws9, `${c}${r}`, b[0]));
                safeSet(ws9, `G${r}`, b[0]);

                ['H', 'I'].forEach(c => safeSet(ws9, `${c}${r}`, b[1]));
                safeSet(ws9, `J${r}`, b[1]);

                safeSet(ws9, `K${r}`, b[2]);

                ['L', 'M'].forEach(c => safeSet(ws9, `${c}${r}`, b[3]));
                safeSet(ws9, `N${r}`, b[3]);

                ['O', 'P'].forEach(c => safeSet(ws9, `${c}${r}`, b[4]));
                safeSet(ws9, `Q${r}`, b[4]);

                ['R', 'S'].forEach(c => safeSet(ws9, `${c}${r}`, b[5]));
                safeSet(ws9, `T${r}`, b[5]);

                ['U', 'V', 'W'].forEach(c => safeSet(ws9, `${c}${r}`, b[6]));
                safeSet(ws9, `X${r}`, b[6]);

                ['Y', 'Z'].forEach(c => safeSet(ws9, `${c}${r}`, b[7]));
                safeSet(ws9, `AA${r}`, b[7]);

                const sumB = b.reduce((a, x) => parseFloat(a) + parseFloat(x), 0);
                safeSet(ws9, `AB${r}`, Math.round((sumB / 8) * 100) / 100);
            });
        }

        // 10. อ่าน-คิด-เขียน (10)
        const ws10 = wb.getWorksheet('อ่าน-คิด-เขียน (10)');
        if (ws10) {
            // Clear student rows block first to avoid dangling shared-formula clones
            clearRect(ws10, 7, 56, 1, 40); // A:AN, rows 7-56

            const READ_SKILLS = ['r1', 'r2', 'r3', 'r4', 'r5'];
            students.forEach((s, i) => {
                const d = getSData(s);
                let rv = [];
                for (let j = 0; j < 5; j++) rv.push(d[READ_SKILLS[j]] !== undefined ? d[READ_SKILLS[j]] : 3);

                let r = 7 + i;
                ws10.getCell(`A${r}`).value = s.roll_number || (i + 1);
                ws10.getCell(`B${r}`).value = s.name || '';

                ['C', 'D', 'E'].forEach(c => safeSet(ws10, `${c}${r}`, rv[0]));
                safeSet(ws10, `F${r}`, rv[0]);

                ['G', 'H', 'I'].forEach(c => safeSet(ws10, `${c}${r}`, rv[1]));
                safeSet(ws10, `J${r}`, rv[1]);

                ['K', 'L', 'M'].forEach(c => safeSet(ws10, `${c}${r}`, rv[2]));
                safeSet(ws10, `N${r}`, rv[2]);

                ['O', 'P', 'Q'].forEach(c => safeSet(ws10, `${c}${r}`, rv[3]));
                safeSet(ws10, `R${r}`, rv[3]);

                ['S', 'T', 'U'].forEach(c => safeSet(ws10, `${c}${r}`, rv[4]));
                safeSet(ws10, `V${r}`, rv[4]);

                const sumR = rv.reduce((a, x) => parseFloat(a) + parseFloat(x), 0);
                safeSet(ws10, `W${r}`, Math.round((sumR / 5) * 100) / 100);
            });
        }

        const levelObj = (activeSubj.grade_level || "รายวิชา").replace(/\//g, '-').replace(/ /g, '_');
        const fname = `ปพ5_${activeSubj.subject_code || 'รายวิชา'}_${levelObj}.xlsx`;

        const outBuffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8' });
        saveAs(blob, fname);

        if (btn) btn.innerText = oldText;
        alert(`✅ ส่งออกไฟล์ ${fname} สำเร็จ!`);
    } catch (e) {
        console.error(e);
        alert('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: ' + e.message);
        const btn = document.querySelector('button[onclick="exportExcel()"]');
        if (btn) btn.innerText = '📥 Export Excel';
    }
}
