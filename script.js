// Global Data Store
let workbookData = {
    schoolName: '',
    examName: '',
    className: '',
    subjects: [],
    students: [],
    sections: []
};

let rawWorkbook = null;
let currentStudent = null;
let currentMode = 'single';

// Teacher Signatures Mapping
const teacherSignatures = {
    "মোহাম্মদ আনিসুজ্জামান": "images/anis.jpeg",
    "মো.হুমায়ুন কবির": "images/humayun.jpeg",
    "মো.আব্দুল সাত্তার সরকার": "images/romana.jpeg", //signature nai
    "মুহম্মদ ছাইফুল ইসলাম": "images/saiful.jpeg",
    "আমিনুল ইসলাম খাঁন": "images/aminul.jpeg",
    "মো.আকরাম হোসেন": "images/akram.jpeg",
    "মো:আসাদুজ্জামান": "images/asad.jpeg", // signature nai
    "রোমানা আক্তার": "images/romana.jpeg",
    "মোহাম্মদ শাহীন মিঞা": "images/shahin.jpeg",
    "ফলাফল প্রস্তুতকারী": "images/nasrin.jpeg",
    "সোহরাব উদ্দিন": "images/shorab.jpeg" //signature nai
};

function getSignatureSrc(teacherName) {
    return teacherSignatures[teacherName] || 'images/romana.jpeg';
}

function toBn(num) {
    if (num === null || num === undefined || num === '') return '০';
    const bengaliDigits = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
    return String(num).replace(/[0-9]/g, d => bengaliDigits[d]);
}

function getBengaliRankStr(rank) {
    if (!rank || rank <= 0) return '—';
    const numBn = toBn(rank);
    if (rank > 10) return `${numBn}তম`;
    const lastDigit = rank % 10;
    let suffix = 'ম';
    if (lastDigit === 1) suffix = 'ম';
    else if (lastDigit === 2 || lastDigit === 3) suffix = 'য়';
    else if (lastDigit === 4) suffix = 'র্থ';
    else if (lastDigit === 6) suffix = 'ষ্ঠ';
    return `${numBn}${suffix}`;
}

function getCleanClassNumber() {
    let rawClass = workbookData.className || '৬ষ্ঠ';
    let match = rawClass.match(/\d+/);
    if (match) return match[0];
    const bnToEn = { '০':'0','১':'1','২':'2','৩':'3','৪':'4','৫':'5','৬':'6','৭':'7','৮':'8','৯':'9' };
    let enStr = rawClass.replace(/[০-৯]/g, d => bnToEn[d]);
    let enMatch = enStr.match(/\d+/);
    return enMatch ? enMatch[0] : '6';
}

function getPDFFileName(student) {
    let classNum = getCleanClassNumber();
    let rollClean = student && student.roll ? String(student.roll).trim() : '1';
    let sectionClean = student && student.section ? String(student.section).trim() : 'A';
    return `${rollClean}_${classNum}-${sectionClean}.pdf`;
}

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-file-input');
    const uploadZone = document.getElementById('upload-zone');

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('drag-over');
    });

    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('drag-over');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelect();
        }
    });

    fileInput.addEventListener('change', handleFileSelect);
});

function handleFileSelect() {
    const fileInput = document.getElementById('excel-file-input');
    const display = document.getElementById('file-name-display');
    const processBtn = document.getElementById('process-btn');

    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        display.textContent = `📄 নির্বাচিত ফাইল: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        display.classList.remove('hidden');
        processBtn.disabled = false;
    } else {
        display.classList.add('hidden');
        processBtn.disabled = true;
    }
}

function processExcel() {
    const fileInput = document.getElementById('excel-file-input');
    if (!fileInput.files.length) return;

    const file = fileInput.files[0];
    const processBtn = document.getElementById('process-btn');
    processBtn.disabled = true;
    processBtn.innerHTML = `<span class="spinner"></span> প্রসেস করা হচ্ছে...`;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            rawWorkbook = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });

            const firstSheetName = rawWorkbook.SheetNames[0];
            const worksheet = rawWorkbook.Sheets[firstSheetName];

            parseWorksheet(worksheet);

            showProcessingStatus(true);
            populateSectionDropdowns();

            document.getElementById('search-section').classList.remove('hidden');
            document.getElementById('search-section').scrollIntoView({ behavior: 'smooth' });
        } catch (err) {
            console.error(err);
            showProcessingStatus(false, err.message);
        } finally {
            processBtn.disabled = false;
            processBtn.innerHTML = `⚙️ Excel প্রসেস করুন`;
        }
    };
    reader.readAsArrayBuffer(file);
}

function getCellValue(sheet, r, c) {
    const cellAddress = XLSX.utils.encode_cell({ r: r, c: c });
    const cell = sheet[cellAddress];
    if (!cell || cell.v === undefined || cell.v === null) return null;
    return cell.v;
}

function setCell(sheet, r, c, val, type = 's') {
    const addr = XLSX.utils.encode_cell({ r: r, c: c });
    sheet[addr] = { v: String(val), t: type };
}

function getSubjectPassMark(subjName, hyMax) {
    const name = String(subjName || '').trim();
    if (name.includes('বাংলা ২য়') || name.includes('বাংলা ২') || (name.includes('বাংলা') && name.includes('২'))) return 17;
    if (name.includes('ইংরেজি ২য়') || name.includes('ইংরেজি ২') || (name.includes('ইংরেজি') && name.includes('২')) || name.toLowerCase().includes('english 2')) return 17;
    if (name.includes('গার্হস্থ') || name.includes('কৃষি') || name.toLowerCase().includes('agriculture')) return 17;
    if (name.includes('আইসিটি') || name.toLowerCase().includes('ict')) return 8;
    if (hyMax === 100) return 33;
    if (hyMax === 50) return 17;
    if (hyMax === 25) return 8;
    return Math.ceil(hyMax * 0.33);
}

function parseWorksheet(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:Z500');

    workbookData.schoolName = String(getCellValue(sheet, 1, 0) || 'সেরাজনগর মুনছর আলী পাইলট মডেল সরকারি উচ্চ বিদ্যালয়').trim();
    workbookData.schoolName = workbookData.schoolName.replace(/\u09b8\u09b0\u0995\u09be\u09b0\u09c0/g, '\u09b8\u09b0\u0995\u09be\u09b0\u09bf');
    workbookData.examName = String(getCellValue(sheet, 3, 2) || 'অর্ধবাষিক মূল্যায়ন প্রতিবেদন-২০২৬').trim();
    let rawClass = String(getCellValue(sheet, 4, 2) || 'শ্রেণি-৬ষ্ঠ').trim();
    workbookData.className = rawClass.replace(/^শ্রেণি-?/i, '').replace(/^শ্রেণী-?/i, '').trim() || '৬ষ্ঠ';

    const mergedRanges = sheet['!merges'] || [];
    function getMergedValue(r, c) {
        for (let m of mergedRanges) {
            if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) return getCellValue(sheet, m.s.r, m.s.c);
        }
        return getCellValue(sheet, r, c);
    }

    let subjectsMap = [];
    for (let c = 4; c <= range.e.c; c++) {
        let subjRaw = getMergedValue(7, c) || getCellValue(sheet, 7, c);
        let subjName = subjRaw ? String(subjRaw).trim() : null;

        if (subjName && ['সর্বমোট', 'অকৃতকার্য বিষয়', 'মেধাক্রম', 'ফলাফলা', 'ফলাফল'].some(k => subjName.includes(k))) continue;

        if (subjName) {
            let assessName = String(getCellValue(sheet, 8, c) || '').trim();
            let maxMark = parseFloat(getCellValue(sheet, 9, c)) || 0;

            let existing = subjectsMap.find(s => s.name === subjName);
            if (!existing) {
                existing = {
                    name: subjName,
                    classTest: { col: null, max: 0, highestObtained: 0 },
                    halfYearly: { col: null, max: 0, passMark: 0, highestObtained: 0 },
                    totalMax: 0
                };
                subjectsMap.push(existing);
            }

            if (assessName.includes('ক্লাস') || assessName.includes('CT') || c % 2 === 4 % 2) {
                if (assessName.includes('অর্ধবার্ষিক')) {
                    existing.halfYearly.col = c;
                    existing.halfYearly.max = maxMark || 100;
                    existing.halfYearly.passMark = getSubjectPassMark(existing.name, existing.halfYearly.max);
                } else {
                    existing.classTest.col = c;
                    existing.classTest.max = maxMark || 50;
                }
            } else if (assessName.includes('অর্ধবার্ষিক') || assessName.includes('HY')) {
                existing.halfYearly.col = c;
                existing.halfYearly.max = maxMark || 100;
                existing.halfYearly.passMark = getSubjectPassMark(existing.name, existing.halfYearly.max);
            } else {
                if (maxMark <= 50) {
                    existing.classTest.col = c;
                    existing.classTest.max = maxMark;
                } else {
                    existing.halfYearly.col = c;
                    existing.halfYearly.max = maxMark;
                    existing.halfYearly.passMark = getSubjectPassMark(existing.name, existing.halfYearly.max);
                }
            }
        }
    }

    subjectsMap.forEach(s => { s.totalMax = s.halfYearly.max || 100; });
    workbookData.subjects = subjectsMap;
    workbookData.sectionHighest = {};

    let students = [];
    let currentSection = 'A';
    let sectionSet = new Set();

    for (let r = 10; r <= range.e.r; r++) {
        let rollVal = getCellValue(sheet, r, 1);
        let nameVal = getCellValue(sheet, r, 2);
        let secVal  = getCellValue(sheet, r, 3);

        if (secVal && String(secVal).trim().length === 1) currentSection = String(secVal).trim().toUpperCase();

        if (nameVal || rollVal !== null) {
            let nameStr = nameVal ? String(nameVal).trim() : '';
            let rollStr = rollVal !== null ? String(rollVal).trim() : '';

            if (nameStr.includes('ছাত্র') || nameStr.includes('নাম') || rollStr.includes('রোল')) continue;
            if (!nameStr && !rollStr) continue;

            sectionSet.add(currentSection);

            let studentObj = {
                roll: rollStr,
                name: nameStr || 'নামবিহীন',
                section: currentSection,
                subjects: {},
                totalObtained: 0,
                totalMax: 0,
                percentage: 0,
                failedCount: 0,
                rowIndex: r
            };

            subjectsMap.forEach(subj => {
                let ctObt = subj.classTest.col !== null ? getCellValue(sheet, r, subj.classTest.col) : null;
                let hyObt = subj.halfYearly.col !== null ? getCellValue(sheet, r, subj.halfYearly.col) : null;

                let ctVal = (ctObt !== null && ctObt !== undefined && ctObt !== '') ? parseFloat(ctObt) : 0;
                let hyVal = (hyObt !== null && hyObt !== undefined && hyObt !== '') ? parseFloat(hyObt) : 0;

                let passMark = subj.halfYearly.passMark;
                let isFailed = hyVal < passMark;

                if (isFailed) studentObj.failedCount++;

                if (subj.classTest.col !== null && ctObt !== null) {
                    setCell(sheet, r, subj.classTest.col, toBn(ctVal), 's');
                }
                if (subj.halfYearly.col !== null && hyObt !== null) {
                    let textVal = toBn(hyVal);
                    if (isFailed) textVal += ' (F)';
                    setCell(sheet, r, subj.halfYearly.col, textVal, 's');
                }

                let subjTotal = hyVal;

                studentObj.subjects[subj.name] = {
                    classTest: { obtained: (ctObt !== null && ctObt !== '') ? ctVal : null, max: subj.classTest.max, hasCol: subj.classTest.col !== null },
                    halfYearly: { obtained: (hyObt !== null && hyObt !== '') ? hyVal : null, max: subj.halfYearly.max, passMark: passMark, hasCol: subj.halfYearly.col !== null },
                    totalObtained: subjTotal,
                    totalMax: subj.halfYearly.max,
                    isFailed: isFailed
                };

                studentObj.totalObtained += subjTotal;
                studentObj.totalMax += subj.halfYearly.max;
            });

            studentObj.percentage = studentObj.totalMax > 0 ? (studentObj.totalObtained / studentObj.totalMax) * 100 : 0;
            students.push(studentObj);
        }
    }

    workbookData.students = students;
    workbookData.sections = Array.from(sectionSet).sort();

    let sortedOverall = [...students].sort((a, b) => b.totalObtained - a.totalObtained);
    sortedOverall.forEach((st, idx) => { st.rankOverall = idx + 1; });

    workbookData.sections.forEach(sec => {
        let secStudents = students.filter(s => s.section === sec);
        secStudents.sort((a, b) => b.totalObtained - a.totalObtained);
        secStudents.forEach((st, idx) => { st.rankSection = idx + 1; });

        let secHighest = {};
        subjectsMap.forEach(subj => {
            let ctMax = 0, hyMax = 0;
            let hasCTData = false, hasHYData = false;
            secStudents.forEach(st => {
                let stSubj = st.subjects[subj.name];
                if (!stSubj) return;
                if (stSubj.classTest.obtained !== null) {
                    hasCTData = true;
                    if (stSubj.classTest.obtained > ctMax) ctMax = stSubj.classTest.obtained;
                }
                if (stSubj.halfYearly.obtained !== null) {
                    hasHYData = true;
                    if (stSubj.halfYearly.obtained > hyMax) hyMax = stSubj.halfYearly.obtained;
                }
            });
            secHighest[subj.name] = { ct: ctMax, hy: hyMax, hasCT: hasCTData, hasHY: hasHYData };
        });
        workbookData.sectionHighest[sec] = secHighest;
    });

    setCell(sheet, 7, 24, 'সর্বমোট', 's');
    setCell(sheet, 7, 25, 'অকৃতকার্য বিষয়', 's');
    setCell(sheet, 7, 26, 'মেধাক্রম', 's');
    setCell(sheet, 7, 27, 'ফলাফল', 's');

    students.forEach(st => {
        let r = st.rowIndex;
        setCell(sheet, r, 24, toBn(st.totalObtained), 's');
        setCell(sheet, r, 25, toBn(st.failedCount), 's');
        setCell(sheet, r, 26, getBengaliRankStr(st.rankSection), 's');
        setCell(sheet, r, 27, st.failedCount === 0 ? 'P' : 'F', 's');
    });

    if (range.e.c < 27) {
        range.e.c = 27;
        sheet['!ref'] = XLSX.utils.encode_range(range);
    }
}

function downloadUpdatedExcel() {
    if (!rawWorkbook) {
        alert('কোনো Excel ফাইল লোড করা হয়নি!');
        return;
    }
    try {
        XLSX.writeFile(rawWorkbook, 'updated_result.xlsx');
    } catch (err) {
        console.error('Excel download error:', err);
        alert('Excel ফাইল ডাউনলোড করতে সমস্যা হয়েছে: ' + err.message);
    }
}

function showProcessingStatus(success, errorMsg = '') {
    const statusPanel = document.getElementById('status-panel');
    const statusContent = document.getElementById('status-content');
    const diagPanel = document.getElementById('diagnostic-panel');
    const diagContent = document.getElementById('diagnostic-content');

    statusPanel.classList.remove('hidden');

    if (!success) {
        statusContent.innerHTML = `
            <div class="status-item error">
                <span class="status-icon">❌</span>
                <div><strong>ত্রুটি!</strong> Excel ফাইল প্রসেস করতে ব্যর্থ হয়েছে।<br><small>${errorMsg}</small></div>
            </div>
        `;
        return;
    }

    let subjChips = workbookData.subjects.map(s => `<span class="subject-chip">✓ ${s.name}</span>`).join('');

    statusContent.innerHTML = `
        <div class="status-item success">
            <span class="status-icon">✅</span>
            <div><strong>সফলভাবে লোড হয়েছে!</strong> ${workbookData.students.length} জন ছাত্র এবং ${workbookData.subjects.length} টি বিষয় সনাক্ত হয়েছে।</div>
        </div>
        <div class="status-item info">
            <span class="status-icon">📌</span>
            <div><strong>সনাক্ত হওয়া শাখা:</strong> ${workbookData.sections.join(', ')}</div>
        </div>
        <div style="margin-top: 8px;">
            <div style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 4px;"><strong>সনাক্ত হওয়া বিষয়সমূহ:</strong></div>
            <div class="subjects-grid">${subjChips}</div>
        </div>
    `;

    diagPanel.classList.remove('hidden');
    let diagLines = [
        `বিদ্যালয়: ${workbookData.schoolName}`,
        `পরীক্ষা: ${workbookData.examName}`,
        `শ্রেণি: ${workbookData.className}`,
        `মোট শিক্ষার্থী: ${workbookData.students.length}`,
        `মোট বিষয়: ${workbookData.subjects.length}`
    ];
    diagContent.textContent = diagLines.join('\n');
}

function toggleDiagnostic() {
    const content = document.getElementById('diagnostic-content');
    const icon = document.getElementById('diag-toggle-icon');
    content.classList.toggle('hidden');
    icon.textContent = content.classList.contains('hidden') ? '▼' : '▲';
}

function populateSectionDropdowns() {
    const datalist = document.getElementById('section-list');
    const batchSelect = document.getElementById('batch-section-select');

    datalist.innerHTML = workbookData.sections.map(sec => `<option value="${sec}"></option>`).join('');
    batchSelect.innerHTML = workbookData.sections.map(sec => `<option value="${sec}">শাখা ${sec}</option>`).join('');

    if (workbookData.sections.length > 0) {
        document.getElementById('section-input').value = workbookData.sections[0];
    }
}

function searchStudent() {
    const rollInput = document.getElementById('roll-input').value.trim();
    const sectionInput = document.getElementById('section-input').value.trim().toUpperCase();
    const nameInput = document.getElementById('name-input').value.trim();
    const t1Name = document.getElementById('teacher1-name-select').value.trim();
    const asstHeadName = document.getElementById('asst-head-name-select').value.trim();
    const examCommitteeName = document.getElementById('exam-committee-name-select').value.trim();

    const errorDiv = document.getElementById('search-error');
    const marksheetSec = document.getElementById('marksheet-section');
    const sectionTitle = document.getElementById('marksheet-section-title');

    errorDiv.classList.add('hidden');

    if (!workbookData.students.length) {
        errorDiv.textContent = 'দয়া করে প্রথমে একটি Excel ফাইল আপলোড ও প্রসেস করুন।';
        errorDiv.classList.remove('hidden');
        return;
    }

    let found = null;

    if (rollInput && sectionInput) {
        found = workbookData.students.find(s =>
            s.roll.toString().toLowerCase() === rollInput.toLowerCase() &&
            s.section.toUpperCase() === sectionInput
        );
    }
    if (!found && rollInput) {
        found = workbookData.students.find(s => s.roll.toString().toLowerCase() === rollInput.toLowerCase());
    }
    if (!found && nameInput) {
        found = workbookData.students.find(s => s.name.toLowerCase().includes(nameInput.toLowerCase()));
    }

    if (!found) {
        errorDiv.innerHTML = `⚠️ কোনো শিক্ষার্থী পাওয়া যায়নি!<br>রোল: <strong>${rollInput || '—'}</strong>, শাখা: <strong>${sectionInput || '—'}</strong>`;
        errorDiv.classList.remove('hidden');
        marksheetSec.classList.add('hidden');
        return;
    }

    currentStudent = found;
    currentMode = 'single';
    sectionTitle.textContent = `মার্কশিট — ${found.name} (রোল: ${found.roll}, শাখা: ${found.section})`;

    const display = document.getElementById('marksheet-display');
    display.innerHTML = generateSingleMarksheetHTML(found, t1Name, asstHeadName, examCommitteeName);

    marksheetSec.classList.remove('hidden');
    marksheetSec.scrollIntoView({ behavior: 'smooth' });
}

// Generate Marksheet HTML with Original Styling + Logo Header & 2-Row Signature Layout
//function generateSingleMarksheetHTML(student, t1Name, asstHeadName, examCommitteeName) {
//    const secHighest = workbookData.sectionHighest[student.section] || {};
//
//    let tableRowsHTML = '';
//    workbookData.subjects.forEach((subj, idx) => {
//        let stSubj = student.subjects[subj.name] || {
//            classTest: { obtained: null },
//            halfYearly: { obtained: null, passMark: 33 },
//            totalObtained: 0,
//            isFailed: false
//        };
//
//        let sh = secHighest[subj.name] || { ct: 0, hy: 0 };
//
//        let ctObtStr = stSubj.classTest.obtained !== null ? toBn(stSubj.classTest.obtained) : '—';
//        let hyObtStr = stSubj.halfYearly.obtained !== null ? toBn(stSubj.halfYearly.obtained) : '—';
//        let hyPassStr = toBn(stSubj.halfYearly.passMark);
//        let hyHighestStr = sh.hasHY ? toBn(sh.hy) : '—';
//
//        let failClass = stSubj.isFailed ? 'style="color: red; font-weight: bold;"' : '';
//
//        tableRowsHTML += `
//            <tr>
//                <td style="text-align: center;">${toBn(idx + 1)}</td>
//                <td style="text-align: left; font-weight: 600;">${subj.name}</td>
//                <td style="text-align: center;">${toBn(subj.halfYearly.max)}</td>
//                <td style="text-align: center;">${hyPassStr}</td>
//                <td style="text-align: center;">${ctObtStr}</td>
//                <td style="text-align: center;" ${failClass}>${hyObtStr}</td>
//                <td style="text-align: center;">${hyHighestStr}</td>
//            </tr>
//        `;
//    });
//
//    const sig1Src = getSignatureSrc("শামীমা নাসরিন");
//    const sig2Src = getSignatureSrc("আমিনুল ইসলাম খান");
//    const sigClassTeacherSrc = getSignatureSrc(t1Name);
//    const sigAsstHeadSrc = getSignatureSrc(asstHeadName);
//    const sigExamCommitteeSrc = getSignatureSrc(examCommitteeName);
//
//    return `
//        <div class="paper-marksheet-container">
//            <!-- Header Section with Top-Left Logo -->
//            <div class="pm-header">
//                <div class="pm-header-grid">
//                    <div class="pm-school-logo-wrap">
//                        <img src="images/logo.webp" alt="School Logo" class="pm-school-logo">
//                    </div>
//                    <div class="pm-header-center">
//                        <div class="pm-school-name">${workbookData.schoolName}</div>
//                        <div class="pm-school-sub1">ডাকঘর: সেরাজনগর, উপজেলা: রায়পুরা, জেলা: নরসিংদী।</div>
//                        <div class="pm-school-sub2">স্থাপিত: ১৯১১ খ্রিঃ | EIIN: 112839 | বিদ্যালয় কোড: ৩8৫১</div>
//                        <div class="pm-school-contact">ইমেইল: smampg.hs112839@gmail.com</div>
//                    </div>
//                </div>
//            </div>
//
//            <div style="text-align: center;">
//                <span class="pm-title-box">${workbookData.examName}</span>
//            </div>
//
//            <!-- Student Metadata -->
//            <div class="pm-student-meta">
//                <table style="width: 100%; border-collapse: collapse; margin-bottom: 6px;">
//                    <tr>
//                        <td style="padding: 3px 0;"><strong>শিক্ষার্থীর নাম:</strong> ${student.name}</td>
//                        <td style="padding: 3px 0; text-align: right;"><strong>শ্রেণি:</strong> ${workbookData.className}</td>
//                    </tr>
//                    <tr>
//                        <td style="padding: 3px 0;"><strong>রোল নম্বর:</strong> ${toBn(student.roll)}</td>
//                        <td style="padding: 3px 0; text-align: right;"><strong>শাখা:</strong> ${student.section}</td>
//                    </tr>
//                    <tr>
//                        <td style="padding: 3px 0;"><strong>মেধাক্রম (শাখা):</strong> ${getBengaliRankStr(student.rankSection)}</td>
//                        <td style="padding: 3px 0; text-align: right;"><strong>ফলাফল:</strong> <span style="font-weight: bold; color: ${student.failedCount === 0 ? 'green' : 'red'};">${student.failedCount === 0 ? 'কৃতকার্য' : 'অকৃতকার্য (' + toBn(student.failedCount) + ' বিষয়)'}</span></td>
//                    </tr>
//                </table>
//            </div>
//
//            <!-- Marks Table -->
//            <table class="pm-table" style="width: 100%; border: 1px solid #000; text-align: center; margin-bottom: 8px;">
//                <thead>
//                    <tr style="background-color: #f2f2f2;">
//                        <th style="border: 1px solid #000; width: 6%;">ক্র:</th>
//                        <th style="border: 1px solid #000; width: 34%; text-align: left;">বিষয়</th>
//                        <th style="border: 1px solid #000; width: 12%;">পূর্ণমান</th>
//                        <th style="border: 1px solid #000; width: 12%;">পাশ নম্বর</th>
//                        <th style="border: 1px solid #000; width: 12%;">ক্লাস টেস্ট</th>
//                        <th style="border: 1px solid #000; width: 12%;">প্রাপ্ত নম্বর</th>
//                        <th style="border: 1px solid #000; width: 12%;">সর্বোচ্চ নম্বর</th>
//                    </tr>
//                </thead>
//                <tbody>
//                    ${tableRowsHTML}
//                    <tr style="font-weight: bold; background-color: #fafafa;">
//                        <td colspan="2" style="border: 1px solid #000; text-align: right; padding-right: 8px;" class="total-row-label">সর্বমোট নম্বর:</td>
//                        <td colspan="5" style="border: 1px solid #000; text-align: left; padding-left: 8px;" class="total-row-value">${toBn(student.totalObtained)} / ${toBn(student.totalMax)}</td>
//                    </tr>
//                </tbody>
//            </table>
//
//            <!-- Signatures Section (2 Rows Layout with Line Spacing for Bengali Matras) -->
//            <div class="pm-signatures-wrapper">
//                <!-- Row 1: 3 Signatures -->
//                <div class="pm-sig-row">
//                    <!-- 1. Result Preparation (Dual Signatures) -->
//                    <div class="pm-sig-box">
//                        <div class="pm-sig-img-container">
//                            <img src="${sig1Src}" alt="Signature" class="pm-sig-img" style="margin-right: 4px;">
//                            <img src="${sig2Src}" alt="Signature" class="pm-sig-img">
//                        </div>
//                        <div class="pm-sig-line"></div>
//                        <div class="pm-sig-name">শামীমা নাসরিন / আমিনুল ইসলাম খান</div>
//                        <div class="pm-sig-title">ফলাফল প্রস্তুতকারী</div>
//                    </div>
//
//                    <!-- 2. Class Teacher -->
//                    <div class="pm-sig-box">
//                        <div class="pm-sig-img-container">
//                            <img src="${sigClassTeacherSrc}" alt="Signature" class="pm-sig-img">
//                        </div>
//                        <div class="pm-sig-line"></div>
//                        <div class="pm-sig-name">${t1Name}</div>
//                        <div class="pm-sig-title">শ্রেণি শিক্ষক</div>
//                    </div>
//
//                    <!-- 3. Internal Exam Committee -->
//                    <div class="pm-sig-box">
//                        <div class="pm-sig-img-container">
//                            <img src="${sigExamCommitteeSrc}" alt="Signature" class="pm-sig-img">
//                        </div>
//                        <div class="pm-sig-line"></div>
//                        <div class="pm-sig-name">${examCommitteeName}</div>
//                        <div class="pm-sig-title">অভ্যন্তরীন পরীক্ষা কমিটি</div>
//                    </div>
//                </div>
//
//                <!-- Row 2: 2 Signatures -->
//                <div class="pm-sig-row" style="justify-content: space-evenly;">
//                    <!-- 4. Assistant Headteacher -->
//                    <div class="pm-sig-box">
//                        <div class="pm-sig-img-container">
//                            <img src="${sigAsstHeadSrc}" alt="Signature" class="pm-sig-img">
//                        </div>
//                        <div class="pm-sig-line"></div>
//                        <div class="pm-sig-name">${asstHeadName}</div>
//                        <div class="pm-sig-title">সহকারী প্রধান শিক্ষক</div>
//                    </div>
//
//                    <!-- 5. Headteacher -->
//                    <div class="pm-sig-box">
//                        <div class="pm-sig-img-container">
//                            <img src="images/romana.jpeg" alt="Signature" class="pm-sig-img">
//                        </div>
//                        <div class="pm-sig-line"></div>
//                        <div class="pm-sig-name">প্রধান শিক্ষক</div>
//                        <div class="pm-sig-title">প্রধান শিক্ষক</div>
//                    </div>
//                </div>
//            </div>
//        </div>
//    `;
//}

// Generate Marksheet HTML matching exact table format from reference image
// Generate Marksheet HTML with updated teacher signatures & layout specs
function generateSingleMarksheetHTML(student, t1Name, asstHeadName, examCommitteeName) {
    const secHighest = workbookData.sectionHighest[student.section] || {};
    const totalSubjects = workbookData.subjects.length;

    // Shift Assignment: D, E, F -> প্রভাতি, Else -> দিবা
    const sectionUpper = (student.section || '').toString().trim().toUpperCase();
    const shiftName = ['D', 'E', 'F'].includes(sectionUpper) ? 'প্রভাতি' : 'দিবা';

    let tableRowsHTML = '';

    workbookData.subjects.forEach((subj, idx) => {
        let stSubj = student.subjects[subj.name] || {
            classTest: { obtained: null },
            halfYearly: { obtained: null, passMark: 33 },
            totalObtained: 0,
            isFailed: false
        };

        let sh = secHighest[subj.name] || { ct: 0, hy: 0 };

        // Class Test Highest & Half Yearly Highest
        let ctHighestStr = sh.hasCT ? toBn(sh.ct) : '—';
        let hyHighestStr = sh.hasHY ? toBn(sh.hy) : '—';

        // Class Test Obtained & Half Yearly Obtained
        let ctObtStr = (stSubj.classTest.obtained !== null && stSubj.classTest.obtained !== undefined) ? toBn(stSubj.classTest.obtained) : '—';
        let hyObtStr = (stSubj.halfYearly.obtained !== null && stSubj.halfYearly.obtained !== undefined) ? toBn(stSubj.halfYearly.obtained) : '—';
        let subjTotalStr = toBn(stSubj.totalObtained);

        // Format total marks column with red border box and (F) tag if failed
        let totalCellHTML = subjTotalStr;
        if (stSubj.isFailed) {
            totalCellHTML = `<span style="display: inline-block; border: 1.5px solid red; color: red; padding: 1px 4px; border-radius: 3px; font-weight: bold;">${subjTotalStr} (F)</span>`;
        }

        let failStyle = stSubj.isFailed ? 'color: red; font-weight: bold;' : '';

        tableRowsHTML += `
            <tr>
                <td style="text-align: center; border: 1px solid #000; padding: 4px;">${toBn(idx + 1)}</td>
                <td style="text-align: left; font-weight: 600; border: 1px solid #000; padding: 4px 8px;">${subj.name}</td>
                <td style="text-align: center; border: 1px solid #000; padding: 4px;">${ctHighestStr}</td>
                <td style="text-align: center; border: 1px solid #000; padding: 4px;">${hyHighestStr}</td>
                <td style="text-align: center; border: 1px solid #000; padding: 4px;">${ctObtStr}</td>
                <td style="text-align: center; border: 1px solid #000; padding: 4px; ${failStyle}">${hyObtStr}</td>
                <td style="text-align: center; border: 1px solid #000; padding: 4px;">${totalCellHTML}</td>
        `;

        // Attach Rowspan Columns (Failed Count & Merit Position) on the First Row
        if (idx === 0) {
            tableRowsHTML += `
                <td rowspan="${totalSubjects}" style="text-align: center; vertical-align: middle; font-weight: bold; border: 1px solid #000; padding: 4px; font-size: 1.1rem;">
                    ${toBn(student.failedCount)}
                </td>
                <td rowspan="${totalSubjects}" style="text-align: center; vertical-align: middle; font-weight: bold; border: 1px solid #000; padding: 4px; font-size: 1.1rem;">
                    ${getBengaliRankStr(student.rankSection)}
                </td>
            `;
        }

        tableRowsHTML += `</tr>`;
    });

    const sigResultPrepSrc = getSignatureSrc("ফলাফল প্রস্তুতকারী");
    const sigClassTeacherSrc = getSignatureSrc(t1Name);
    const sigAsstHeadSrc = getSignatureSrc(asstHeadName);
    const sigExamCommitteeSrc = getSignatureSrc(examCommitteeName);
    const sigHeadTeacherSrc = getSignatureSrc("সোহরাব উদ্দিন");

    return `
        <div class="paper-marksheet-container" style="background: #fff; padding: 15px; color: #000; font-family: 'Hind Siliguri', 'Anek Bangla', sans-serif; position: relative; overflow: hidden; margin-top: 25px;">
            <!-- Background Watermark Logo -->
            <img src="images/logo.webp" alt="" class="pm-watermark-logo" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 65%; max-width: 500px; opacity: 0.07; z-index: 0; pointer-events: none; user-select: none;">
            <div class="pm-content-wrap" style="position: relative; z-index: 1;">
            <!-- School Header & Logo -->
            <div class="pm-header" style="margin-bottom: 10px;">
                <div class="pm-header-grid" style="display: flex; align-items: center; justify-content: space-between;">
                    <div class="pm-school-logo-wrap" style="flex: 0 0 80px;">
                        <img src="images/logo.webp" alt="School Logo" class="pm-school-logo" style="width: 80px; height: 80px; object-fit: contain;">
                    </div>
                    <div class="pm-header-center" style="flex: 1; text-align: center;">
                        <div class="pm-school-name" style="font-size: 1.4rem; font-weight: 700; color: #000;">${workbookData.schoolName}</div>
                        <div class="pm-school-sub1" style="font-size: 0.85rem;">ডাকঘর: সেরাজনগর, উপজেলা: রায়পুরা, জেলা: নরসিংদী।</div>
                        <div class="pm-school-sub2" style="font-size: 0.8rem;">স্থাপিত: ১৯১১ খ্রিঃ | EIIN: 112839 | বিদ্যালয় কোড: ৩8৫১</div>
                        <div class="pm-school-contact" style="font-size: 0.8rem;">ই-মেইল: smampg.hs112839@gmail.com</div>
                    </div>
                </div>
            </div>

            <div style="text-align: center; margin-bottom: 12px;">
                <span class="pm-title-box" style="display: inline-block; border: 1.5px solid #000; padding: 3px 18px; font-size: 1.1rem; font-weight: 700; border-radius: 4px;">${workbookData.examName}</span>
            </div>

            <!-- Student Info Metadata Header -->
            <div class="pm-student-meta" style="margin-bottom: 8px; font-size: 0.95rem; font-weight: 600;">
                <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <td style="padding: 2px 0; width: 60%;"><strong>শিক্ষার্থীর নাম :</strong> ${student.name}</td>
                        <td style="padding: 2px 0; width: 40%; text-align: right;"><strong>শ্রেণি :</strong> ${workbookData.className}</td>
                    </tr>
                    <tr>
                        <td style="padding: 2px 0;"><strong>শিফট :</strong> ${shiftName}</td>
                        <td style="padding: 2px 0; text-align: right;"><strong>রোল :</strong> ${toBn(student.roll)}</td>
                    </tr>
                </table>
            </div>

            <!-- Subjects & Results Table -->
            <table class="pm-table" style="width: 100%; border-collapse: collapse; border: 1.5px solid #000; text-align: center; font-size: 0.88rem; margin-bottom: 12px;">
                <thead>
                    <tr>
                        <th rowspan="2" style="border: 1px solid #000; width: 6%; padding: 4px;">ক্রমিক নং</th>
                        <th rowspan="2" style="border: 1px solid #000; width: 28%; padding: 4px; text-align: center;">বিষয় এর নাম</th>
                        <th colspan="2" style="border: 1px solid #000; padding: 4px;">শ্রেণিতে বিষয় ভিত্তিক<br>সর্বোচ্চ প্রাপ্ত নম্বর</th>
                        <th colspan="3" style="border: 1px solid #000; padding: 4px;">প্রাপ্ত নম্বর</th>
                        <th rowspan="2" style="border: 1px solid #000; width: 11%; padding: 4px;">অকৃতকার্য বিষয়<br>(অর্ধবার্ষিক)</th>
                        <th rowspan="2" style="border: 1px solid #000; width: 11%; padding: 4px;">মেধাক্রম</th>
                    </tr>
                    <tr>
                        <th style="border: 1px solid #000; width: 10%; padding: 4px;">১ম ক্লাস</th>
                        <th style="border: 1px solid #000; width: 10%; padding: 4px;">অর্ধবার্ষিক</th>
                        <th style="border: 1px solid #000; width: 9%; padding: 4px;">১ম ক্লাস</th>
                        <th style="border: 1px solid #000; width: 9%; padding: 4px;">অর্ধবার্ষিক</th>
                        <th style="border: 1px solid #000; width: 9%; padding: 4px;">মোট নম্বর</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableRowsHTML}
                    <tr>
                        <td colspan="6" style="border: 1px solid #000; text-align: center; font-weight: 700; padding: 6px;">সর্বমোট প্রাপ্ত নম্বর</td>
                        <td style="border: 1px solid #000; text-align: center; font-weight: 800; padding: 6px; font-size: 1rem;">${toBn(student.totalObtained)}</td>
                        <td colspan="2" style="border: 1px solid #000; background-color: #ffffff;"></td>
                    </tr>
                </tbody>
            </table>

            <!-- Signatures Section -->
            <div class="pm-signatures-wrapper" style="margin-top: 25px; display: flex; flex-direction: column; gap: 24px;">
                <!-- Row 1: 3 Signatures -->
                <div class="pm-sig-row" style="display: flex; justify-content: space-around; align-items: flex-end; gap: 15px;">
                    <div class="pm-sig-box" style="flex: 1; text-align: center; max-width: 220px; display: flex; flex-direction: column; align-items: center;">
                        <div class="pm-sig-img-container" style="min-height: 38px; display: flex; align-items: flex-end; justify-content: center; margin-bottom: 4px;">
                            <img src="${sigResultPrepSrc}" alt="Signature" class="pm-sig-img" style="max-height: 36px; max-width: 120px; object-fit: contain;">
                        </div>
                        <div class="pm-sig-line" style="width: 100%; border-top: 1px dashed #000; margin: 4px 0;"></div>
                        <div class="pm-sig-name" style="font-size: 0.85rem; font-weight: 600; line-height: 1.5; color: #1e293b;">শামিমা নাসরিন & আমিনুল ইসলাম খাঁন</div>
                        <div class="pm-sig-title" style="font-size: 0.78rem; font-weight: 500; line-height: 1.4; color: #475569;">ফলাফল প্রস্তুতকারী</div>
                    </div>

                    <div class="pm-sig-box" style="flex: 1; text-align: center; max-width: 220px; display: flex; flex-direction: column; align-items: center;">
                        <div class="pm-sig-img-container" style="min-height: 38px; display: flex; align-items: flex-end; justify-content: center; margin-bottom: 4px;">
                            <img src="${sigClassTeacherSrc}" alt="Signature" class="pm-sig-img" style="max-height: 36px; max-width: 120px; object-fit: contain;">
                        </div>
                        <div class="pm-sig-line" style="width: 100%; border-top: 1px dashed #000; margin: 4px 0;"></div>
                        <div class="pm-sig-name" style="font-size: 0.85rem; font-weight: 600; line-height: 1.5; color: #1e293b;">${t1Name}</div>
                        <div class="pm-sig-title" style="font-size: 0.78rem; font-weight: 500; line-height: 1.4; color: #475569;">শ্রেণি শিক্ষক</div>
                    </div>

                    <div class="pm-sig-box" style="flex: 1; text-align: center; max-width: 220px; display: flex; flex-direction: column; align-items: center;">
                        <div class="pm-sig-img-container" style="min-height: 38px; display: flex; align-items: flex-end; justify-content: center; margin-bottom: 4px;">
                            <img src="${sigExamCommitteeSrc}" alt="Signature" class="pm-sig-img" style="max-height: 36px; max-width: 120px; object-fit: contain;">
                        </div>
                        <div class="pm-sig-line" style="width: 100%; border-top: 1px dashed #000; margin: 4px 0;"></div>
                        <div class="pm-sig-name" style="font-size: 0.85rem; font-weight: 600; line-height: 1.5; color: #1e293b;">${examCommitteeName}</div>
                        <div class="pm-sig-title" style="font-size: 0.78rem; font-weight: 500; line-height: 1.4; color: #475569;">অভ্যন্তরীন পরীক্ষা কমিটি</div>
                    </div>
                </div>

                <!-- Row 2: 2 Signatures -->
                <div class="pm-sig-row" style="display: flex; justify-content: space-evenly; align-items: flex-end; gap: 15px;">
                    <div class="pm-sig-box" style="flex: 1; text-align: center; max-width: 220px; display: flex; flex-direction: column; align-items: center;">
                        <div class="pm-sig-img-container" style="min-height: 38px; display: flex; align-items: flex-end; justify-content: center; margin-bottom: 4px;">
                            <img src="${sigAsstHeadSrc}" alt="Signature" class="pm-sig-img" style="max-height: 36px; max-width: 120px; object-fit: contain;">
                        </div>
                        <div class="pm-sig-line" style="width: 100%; border-top: 1px dashed #000; margin: 4px 0;"></div>
                        <div class="pm-sig-name" style="font-size: 0.85rem; font-weight: 600; line-height: 1.5; color: #1e293b;">${asstHeadName}</div>
                        <div class="pm-sig-title" style="font-size: 0.78rem; font-weight: 500; line-height: 1.4; color: #475569;">সহকারী প্রধান শিক্ষক</div>
                    </div>

                    <div class="pm-sig-box" style="flex: 1; text-align: center; max-width: 220px; display: flex; flex-direction: column; align-items: center;">
                        <div class="pm-sig-img-container" style="min-height: 38px; display: flex; align-items: flex-end; justify-content: center; margin-bottom: 4px;">
                            <img src="${sigHeadTeacherSrc}" alt="Signature" class="pm-sig-img" style="max-height: 36px; max-width: 120px; object-fit: contain;">
                        </div>
                        <div class="pm-sig-line" style="width: 100%; border-top: 1px dashed #000; margin: 4px 0;"></div>
                        <div class="pm-sig-name" style="font-size: 0.85rem; font-weight: 600; line-height: 1.5; color: #1e293b;">সোহরাব উদ্দিন</div>
                        <div class="pm-sig-title" style="font-size: 0.78rem; font-weight: 500; line-height: 1.4; color: #475569;">প্রধান শিক্ষক</div>
                    </div>
                </div>
            </div>
            </div>
        </div>
    `;
}



function viewAllSectionStudents() {
    const sec = document.getElementById('batch-section-select').value;
    const t1Name = document.getElementById('teacher1-name-select').value.trim();
    const asstHeadName = document.getElementById('asst-head-name-select').value.trim();
    const examCommitteeName = document.getElementById('exam-committee-name-select').value.trim();

    const marksheetSec = document.getElementById('marksheet-section');
    const sectionTitle = document.getElementById('marksheet-section-title');
    const display = document.getElementById('marksheet-display');

    let secStudents = workbookData.students.filter(s => s.section === sec);
    if (!secStudents.length) {
        alert('এই শাখায় কোনো শিক্ষার্থী পাওয়া যায়নি!');
        return;
    }

    currentMode = 'batch';
    sectionTitle.textContent = `শাখা ${sec}-এর সকল শিক্ষার্থীর মার্কশিট (${secStudents.length} জন)`;

    display.innerHTML = secStudents.map(st => generateSingleMarksheetHTML(st, t1Name, asstHeadName, examCommitteeName)).join('<div style="page-break-after: always; height: 1px;"></div>');

    marksheetSec.classList.remove('hidden');
    marksheetSec.scrollIntoView({ behavior: 'smooth' });
}

async function downloadPDF() {
    const display = document.getElementById('marksheet-display');
    const element = display.querySelector('.paper-marksheet-container');
    if (!element) return;

    try {
        const canvas = await html2canvas(element, { scale: 2, useCORS: true });
        const imgData = canvas.toDataURL('image/jpeg', 1.0);

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF('p', 'mm', 'a4');
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

        pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
        pdf.save(getPDFFileName(currentStudent));
    } catch (err) {
        console.error('PDF Generation Error:', err);
        alert('PDF তৈরি করতে সমস্যা হয়েছে!');
    }
}

async function downloadAllSectionPDF() {
    const sec = document.getElementById('batch-section-select').value;
    const t1Name = document.getElementById('teacher1-name-select').value.trim();
    const asstHeadName = document.getElementById('asst-head-name-select').value.trim();
    const examCommitteeName = document.getElementById('exam-committee-name-select').value.trim();

    let secStudents = workbookData.students.filter(s => s.section === sec);
    if (!secStudents.length) {
        alert('এই শাখায় কোনো শিক্ষার্থী পাওয়া যায়নি!');
        return;
    }

    const progressContainer = document.getElementById('batch-progress-container');
    const progressText = document.getElementById('batch-progress-text');
    const progressPercent = document.getElementById('batch-progress-percent');
    const progressBarFill = document.getElementById('batch-progress-bar-fill');

    progressContainer.classList.remove('hidden');

    const tempDiv = document.createElement('div');
    tempDiv.style.position = 'absolute';
    tempDiv.style.left = '-9999px';
    tempDiv.style.top = '-9999px';
    document.body.appendChild(tempDiv);

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p', 'mm', 'a4');

    for (let i = 0; i < secStudents.length; i++) {
        const student = secStudents[i];
        const percent = Math.round(((i + 1) / secStudents.length) * 100);

        progressText.textContent = `প্রসেস করা হচ্ছে: ${student.name} (${i + 1}/${secStudents.length})`;
        progressPercent.textContent = `${toBn(percent)}%`;
        progressBarFill.style.width = `${percent}%`;

        tempDiv.innerHTML = generateSingleMarksheetHTML(student, t1Name, asstHeadName, examCommitteeName);
        const element = tempDiv.querySelector('.paper-marksheet-container');

        const canvas = await html2canvas(element, { scale: 2, useCORS: true });
        const imgData = canvas.toDataURL('image/jpeg', 1.0);

        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

        if (i > 0) pdf.addPage();
        pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
    }

    document.body.removeChild(tempDiv);
    progressContainer.classList.add('hidden');

    let classNum = getCleanClassNumber();
    pdf.save(`Section_${sec}_Class_${classNum}_Marksheets.pdf`);
}