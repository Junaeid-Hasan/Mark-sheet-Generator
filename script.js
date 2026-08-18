// Global Data Store
let workbookData = {
    schoolName: '',
    examName: '',
    className: '',
    subjects: [], // { name, classTest: { col, max, highestObtained }, halfYearly: { col, max, passMark, highestObtained }, totalMax }
    students: [], 
    sections: []
};

let rawWorkbook = null;
let currentStudent = null;
let currentMode = 'single'; // 'single' or 'batch'

// Convert digits to Bengali numerals
function toBn(num) {
    if (num === null || num === undefined || num === '') return '০';
    const bengaliDigits = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
    return String(num).replace(/[0-9]/g, d => bengaliDigits[d]);
}

// Convert rank to Bengali ordinal rank string (১ম, ২য়, ৩য়, ৪র্থ, ৫ম, ৬ষ্ঠ, ইত্যাদি)
function getBengaliRankStr(rank) {
    if (!rank || rank <= 0) return '—';
    const numBn = toBn(rank);
    if (rank > 10) {
        return `${numBn}তম`;
    }
    const lastDigit = rank % 10;
    let suffix = 'ম';
    
    if (lastDigit === 1) {
        suffix = 'ম';
    } else if (lastDigit === 2 || lastDigit === 3) {
        suffix = 'য়';
    } else if (lastDigit === 4) {
        suffix = 'র্থ';
    } else if (lastDigit === 6) {
        suffix = 'ষ্ঠ';
    }
    return `${numBn}${suffix}`;
}

// Clean Class Number for PDF filename (e.g. "6" from "শ্রেণি-৬ষ্ঠ")
function getCleanClassNumber() {
    let rawClass = workbookData.className || '৬ষ্ঠ';
    let match = rawClass.match(/\d+/);
    if (match) return match[0];
    const bnToEn = { '০':'0','১':'1','২':'2','৩':'3','৪':'4','৫':'5','৬':'6','৭':'7','৮':'8','৯':'9' };
    let enStr = rawClass.replace(/[০-৯]/g, d => bnToEn[d]);
    let enMatch = enStr.match(/\d+/);
    return enMatch ? enMatch[0] : '6';
}

// Format PDF filename: rollnum_class-section.pdf (e.g. 01_6-A.pdf)
function getPDFFileName(student) {
    let classNum = getCleanClassNumber();
    let rollClean = student && student.roll ? String(student.roll).trim() : '1';
    let sectionClean = student && student.section ? String(student.section).trim() : 'A';
    return `${rollClean}_${classNum}-${sectionClean}.pdf`;
}

// Event Listeners
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

// Process Excel File
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

// Safe Cell Value getter
function getCellValue(sheet, r, c) {
    const cellAddress = XLSX.utils.encode_cell({ r: r, c: c });
    const cell = sheet[cellAddress];
    if (!cell || cell.v === undefined || cell.v === null) return null;
    return cell.v;
}

// Function to write value to worksheet cell
function setCell(sheet, r, c, val, type = 's') {
    const addr = XLSX.utils.encode_cell({ r: r, c: c });
    sheet[addr] = { v: String(val), t: type };
}

// Pass Marks calculation logic (Half-Yearly Only):
// - Bangla 2nd: 17
// - English 2nd: 17
// - Agriculture / Home Science: 17
// - ICT: 8
// - All other subjects: 33
function getSubjectPassMark(subjName, hyMax) {
    const name = String(subjName || '').trim();
    if (name.includes('বাংলা ২য়') || name.includes('বাংলা ২') || (name.includes('বাংলা') && name.includes('২'))) {
        return 17;
    }
    if (name.includes('ইংরেজি ২য়') || name.includes('ইংরেজি ২') || (name.includes('ইংরেজি') && name.includes('২')) || name.toLowerCase().includes('english 2')) {
        return 17;
    }
    if (name.includes('গার্হস্থ') || name.includes('কৃষি') || name.toLowerCase().includes('agriculture')) {
        return 17;
    }
    if (name.includes('আইসিটি') || name.toLowerCase().includes('ict')) {
        return 8;
    }
    if (hyMax === 100) return 33;
    if (hyMax === 50) return 17;
    if (hyMax === 25) return 8;
    return Math.ceil(hyMax * 0.33);
}

// Parse Worksheet
function parseWorksheet(sheet) {
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:Z500');
    
    // Metadata
    workbookData.schoolName = String(getCellValue(sheet, 1, 0) || 'সেরাজনগর মুনছর আলী পাইলট মডেল সরকারি উচ্চ বিদ্যালয়').trim();
    workbookData.examName = String(getCellValue(sheet, 3, 2) || 'অর্ধবাষিক মূল্যায়ন প্রতিবেদন-২০২৬').trim();
    let rawClass = String(getCellValue(sheet, 4, 2) || 'শ্রেণি-৬ষ্ঠ').trim();
    // Strip either spelling variant ("শ্রেণি-" or the older "শ্রেণী-") that might appear
    // as a prefix in the source Excel file, since we can't control how the sheet was authored.
    let cleanClass = rawClass.replace(/^শ্রেণি-?/i, '').replace(/^শ্রেণী-?/i, '').trim() || '৬ষ্ঠ';
    workbookData.className = cleanClass;

    // teacher class inputs are populated later once sections are detected

    // Merged Cells Map
    const mergedRanges = sheet['!merges'] || [];
    function getMergedValue(r, c) {
        for (let m of mergedRanges) {
            if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) {
                return getCellValue(sheet, m.s.r, m.s.c);
            }
        }
        return getCellValue(sheet, r, c);
    }

    // Detect Subjects (Row 7 = 8th row 1-indexed)
    let subjectsMap = [];

    for (let c = 4; c <= range.e.c; c++) {
        let subjRaw = getMergedValue(7, c) || getCellValue(sheet, 7, c);
        let subjName = subjRaw ? String(subjRaw).trim() : null;

        if (subjName && ['সর্বমোট', 'অকৃতকার্য বিষয়', 'মেধাক্রম', 'ফলাফলা', 'ফলাফল'].some(k => subjName.includes(k))) {
            continue;
        }

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

    subjectsMap.forEach(s => {
        s.totalMax = s.halfYearly.max || 100; // Only Half-Yearly counts towards total
    });

    workbookData.subjects = subjectsMap;
    // sectionHighest is computed after all students are parsed
    workbookData.sectionHighest = {}; // { section: { subjName: { ct, hy, hasCT, hasHY } } }

    // Student Rows Parsing (Row index 10 to max)
    let students = [];
    let currentSection = 'A';
    let sectionSet = new Set();

    for (let r = 10; r <= range.e.r; r++) {
        let rollVal = getCellValue(sheet, r, 1); // Col B
        let nameVal = getCellValue(sheet, r, 2); // Col C
        let secVal  = getCellValue(sheet, r, 3); // Col D

        if (secVal && String(secVal).trim().length === 1) {
            currentSection = String(secVal).trim().toUpperCase();
        }

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

                // (highest marks are computed in a second pass after all students are collected)

                // Pass/Fail is determined ONLY by Half-Yearly score vs Subject Pass Mark
                // Class Test mark does NOT affect pass/fail status and is NOT added to total
                let passMark = subj.halfYearly.passMark;
                let isFailed = hyVal < passMark;

                if (isFailed) {
                    studentObj.failedCount++;
                }

                // Write Bengali numerals to Excel sheet cells
                if (subj.classTest.col !== null && ctObt !== null) {
                    setCell(sheet, r, subj.classTest.col, toBn(ctVal), 's');
                }
                if (subj.halfYearly.col !== null && hyObt !== null) {
                    let textVal = toBn(hyVal);
                    if (isFailed) textVal += ' (F)';
                    setCell(sheet, r, subj.halfYearly.col, textVal, 's');
                }

                // Subject Total is ONLY Half-Yearly mark
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

    // Ranks Calculation (based on sum of Half-Yearly marks)
    let sortedOverall = [...students].sort((a, b) => b.totalObtained - a.totalObtained);
    sortedOverall.forEach((st, idx) => {
        st.rankOverall = idx + 1;
    });

    workbookData.sections.forEach(sec => {
        let secStudents = students.filter(s => s.section === sec);
        secStudents.sort((a, b) => b.totalObtained - a.totalObtained);
        secStudents.forEach((st, idx) => {
            st.rankSection = idx + 1;
        });

        // Compute section-wise highest marks per subject
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

    // Write calculated data into the 4 last columns (Col Y=24, Z=25, AA=26, AB=27) of the raw worksheet
    // Header cells at row index 7 (Row 8 1-indexed)
    setCell(sheet, 7, 24, 'সর্বমোট', 's');
    setCell(sheet, 7, 25, 'অকৃতকার্য বিষয়', 's');
    setCell(sheet, 7, 26, 'মেধাক্রম', 's');
    setCell(sheet, 7, 27, 'ফলাফল', 's');

    students.forEach(st => {
        let r = st.rowIndex;
        // 1. সর্বমোট in Bangla numbers (e.g. ৪২১)
        setCell(sheet, r, 24, toBn(st.totalObtained), 's');
        // 2. অকৃতকার্য বিষয় in Bangla numbers (e.g. ৪ or ০)
        setCell(sheet, r, 25, toBn(st.failedCount), 's');
        // 3. মেধাক্রম in ordinal format (e.g. ১ম, ২য়, ৩য়, ৪র্থ)
        setCell(sheet, r, 26, getBengaliRankStr(st.rankSection), 's');
        // 4. ফলাফল: 'P' for Pass, 'F' for Fail
        setCell(sheet, r, 27, st.failedCount === 0 ? 'P' : 'F', 's');
    });

    // Ensure range bounds include columns Y, Z, AA, AB
    if (range.e.c < 27) {
        range.e.c = 27;
        sheet['!ref'] = XLSX.utils.encode_range(range);
    }
}

// Download Updated Excel file with populated columns
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

// Processing Status UI
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
            <div style="font-size: 0.85rem; color: var(--text-secondary); margin-bottom: 4px;"><strong>সনাক্ত হওয়া বিষয়সমূহ (সর্বোচ্চ প্রাপ্ত নম্বর সহ):</strong></div>
            <div class="subjects-grid">${subjChips}</div>
        </div>
    `;

    diagPanel.classList.remove('hidden');
    let diagLines = [
        `বিদ্যালয়: ${workbookData.schoolName}`,
        `পরীক্ষা: ${workbookData.examName}`,
        `শ্রেণি: ${workbookData.className}`,
        `মোট শিক্ষার্থী: ${workbookData.students.length}`,
        `মোট বিষয়: ${workbookData.subjects.length}`,
        `----------------------------------------`
    ];
    workbookData.subjects.forEach(s => {
        diagLines.push(`বিষয়: ${s.name} (Pass Mark=${s.halfYearly.passMark})`);
        workbookData.sections.forEach(sec => {
            const sh = (workbookData.sectionHighest[sec] || {})[s.name] || {};
            diagLines.push(`  শাখা ${sec}: Highest CT=${sh.hasCT ? sh.ct : '—'} | Highest HY=${sh.hasHY ? sh.hy : '—'}`);
        });
    });
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

// Single Student Search
function searchStudent() {
    const rollInput = document.getElementById('roll-input').value.trim();
    const sectionInput = document.getElementById('section-input').value.trim().toUpperCase();
    const nameInput = document.getElementById('name-input').value.trim();
    const t1Name = document.getElementById('teacher1-name-input').value.trim();
    const t2Name = document.getElementById('teacher2-name-input').value.trim();
    const asstHeadName = document.getElementById('asst-head-name-input').value.trim();
    const examCommitteeName = document.getElementById('exam-committee-name-input').value.trim();

    const errorDiv = document.getElementById('search-error');
    const marksheetSec = document.getElementById('marksheet-section');
    const sectionTitle = document.getElementById('marksheet-section-title');

    errorDiv.classList.add('hidden');

    if (!workbookData.students.length) {
        errorDiv.textContent = 'দয়া করে প্রথমে একটি Excel ফাইল আপলোড ও প্রসেস করুন।';
        errorDiv.classList.remove('hidden');
        return;
    }

    if (!t1Name || !t2Name) {
        errorDiv.innerHTML = '⚠️ অনুগ্রহ করে <strong>শ্রেণি শিক্ষক ও ফলাফল প্রস্তুতকারীর নাম</strong> পূরণ করুন।';
        errorDiv.classList.remove('hidden');
        marksheetSec.classList.add('hidden');
        
        if (!t1Name) document.getElementById('teacher1-name-input').focus();
        else if (!t2Name) document.getElementById('teacher2-name-input').focus();
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
    display.innerHTML = generateSingleMarksheetHTML(found, t1Name, t2Name, asstHeadName, examCommitteeName);

    marksheetSec.classList.remove('hidden');
    marksheetSec.scrollIntoView({ behavior: 'smooth' });
}

// View All Students in Section
// Returns true on success (rendered), false if validation failed / nothing to show.
let currentBatchStudents = [];

function viewAllSectionStudents() {
    const sectionSelect = document.getElementById('batch-section-select');
    const selectedSec = sectionSelect ? sectionSelect.value : '';
    const t1Name = document.getElementById('teacher1-name-input').value.trim();
    const t2Name = document.getElementById('teacher2-name-input').value.trim();
    const asstHeadName = document.getElementById('asst-head-name-input').value.trim();
    const examCommitteeName = document.getElementById('exam-committee-name-input').value.trim();
    const errorDiv = document.getElementById('search-error');
    const marksheetSec = document.getElementById('marksheet-section');
    const sectionTitle = document.getElementById('marksheet-section-title');

    errorDiv.classList.add('hidden');

    if (!workbookData.students.length) {
        errorDiv.textContent = 'দয়া করে প্রথমে একটি Excel ফাইল আপলোড ও প্রসেস করুন।';
        errorDiv.classList.remove('hidden');
        return false;
    }

    if (!selectedSec) {
        errorDiv.textContent = 'দয়া করে একটি শাখা নির্বাচন করুন।';
        errorDiv.classList.remove('hidden');
        return false;
    }

    if (!t1Name || !t2Name) {
        errorDiv.innerHTML = '⚠️ অনুগ্রহ করে <strong>শ্রেণি শিক্ষক ও ফলাফল প্রস্তুতকারীর নাম</strong> পূরণ করুন।';
        errorDiv.classList.remove('hidden');
        marksheetSec.classList.add('hidden');

        if (!t1Name) document.getElementById('teacher1-name-input').focus();
        else if (!t2Name) document.getElementById('teacher2-name-input').focus();
        return false;
    }

    const secStudents = workbookData.students.filter(s => s.section === selectedSec);
    if (!secStudents.length) {
        errorDiv.textContent = `শাখা ${selectedSec}-এ কোনো শিক্ষার্থী পাওয়া যায়নি।`;
        errorDiv.classList.remove('hidden');
        return false;
    }

    currentMode = 'batch';
    currentBatchStudents = secStudents;
    sectionTitle.textContent = `শাখা ${selectedSec}-এর সকল শিক্ষার্থীদের মার্কশিট (${secStudents.length} জন)`;

    const display = document.getElementById('marksheet-display');
    display.innerHTML = secStudents.map(st => generateSingleMarksheetHTML(st, t1Name, t2Name, asstHeadName, examCommitteeName)).join('');

    marksheetSec.classList.remove('hidden');
    marksheetSec.scrollIntoView({ behavior: 'smooth' });
    return true;
}

// Generate Marksheet HTML for a student
function generateSingleMarksheetHTML(student, teacher1Name = '', teacher2Name = '', asstHeadName = '', examCommitteeName = '') {
    let totalSubjs = workbookData.subjects.length;
    let rankStr = getBengaliRankStr(student.rankSection);
    let classTitle = workbookData.className || '৬ষ্ঠ';

    let rowsHTML = '';
    workbookData.subjects.forEach((subj, idx) => {
        let stSubj = student.subjects[subj.name] || {
            classTest: { obtained: null, max: subj.classTest.max },
            halfYearly: { obtained: null, max: subj.halfYearly.max, passMark: 33 },
            totalObtained: 0,
            totalMax: subj.halfYearly.max,
            isFailed: false
        };

        let serialBn = toBn(idx + 1);

        // Section-specific highest marks
        const secHighest = (workbookData.sectionHighest[student.section] || {})[subj.name] || { ct: 0, hy: 0, hasCT: false, hasHY: false };
        let ctHighestBn = secHighest.hasCT ? toBn(secHighest.ct) : '—';
        let hyHighestBn = secHighest.hasHY ? toBn(secHighest.hy) : '—';

        let ctVal = stSubj.classTest.obtained;
        let hyVal = stSubj.halfYearly.obtained;

        let ctObtBn = ctVal !== null ? toBn(ctVal) : '—';
        let hyObtBn = hyVal !== null ? (stSubj.isFailed ? `<span class="failed-mark">${toBn(hyVal)} (F)</span>` : toBn(hyVal)) : '—';
        
        // Subject Total is ONLY Half-Yearly mark
        let totalDisplay = stSubj.isFailed 
            ? `<strong class="failed-mark">${toBn(stSubj.totalObtained)} <span class="failed-badge">(F)</span></strong>`
            : `<strong>${toBn(stSubj.totalObtained)}</strong>`;

        let trClass = stSubj.isFailed ? 'class="failed-subject-row"' : '';

        if (idx === 0) {
            rowsHTML += `
                <tr ${trClass}>
                    <td>${serialBn}</td>
                    <td class="subject-name-td">${subj.name}</td>
                    <td>${ctHighestBn}</td>
                    <td>${hyHighestBn}</td>
                    <td>${ctObtBn}</td>
                    <td>${hyObtBn}</td>
                    <td>${totalDisplay}</td>
                    <td rowspan="${totalSubjs}" class="failed-td">${toBn(student.failedCount)}</td>
                    <td rowspan="${totalSubjs}" class="rank-td">${rankStr}</td>
                </tr>
            `;
        } else {
            rowsHTML += `
                <tr ${trClass}>
                    <td>${serialBn}</td>
                    <td class="subject-name-td">${subj.name}</td>
                    <td>${ctHighestBn}</td>
                    <td>${hyHighestBn}</td>
                    <td>${ctObtBn}</td>
                    <td>${hyObtBn}</td>
                    <td>${totalDisplay}</td>
                </tr>
            `;
        }
    });

    let displayRoll = student.roll.includes('-') ? student.roll : `${student.section}-${toBn(student.roll.padStart(2, '0'))}`;
    let t1 = teacher1Name || '—';
    let t2 = teacher2Name || '—';

    return `
        <div class="paper-marksheet-container marksheet-card-item" data-roll="${student.roll}" data-section="${student.section}">
            <!-- Header -->
            <div class="pm-header">
                <div class="pm-header-center">
                    <div class="pm-school-name">${workbookData.schoolName}</div>
                    <div class="pm-school-sub1">স্থাপিত : ১৯৪৭ ইং</div>
                    <div class="pm-school-sub2">ডাকঘর ও উপজেলা : রায়পুরা,  জেলা : নরসিংদী।</div>
                    <div class="pm-school-contact">E-mail : school112766@gmail.com &nbsp;&nbsp;&nbsp; EIIN : 112766 &nbsp;&nbsp;&nbsp; কোড নং : ৩২২৯</div>
                </div>

                <!-- Top Right Grade Table -->
                <div class="pm-grade-table-wrap">
                    <table class="pm-grade-table">
                        <thead>
                            <tr>
                                <th>Marks</th>
                                <th>L.G</th>
                                <th>G.P</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td>80-100</td><td>A+</td><td>5</td></tr>
                            <tr><td>70-79</td><td>A</td><td>4</td></tr>
                            <tr><td>60-69</td><td>A-</td><td>3.5</td></tr>
                            <tr><td>50-59</td><td>B</td><td>3</td></tr>
                            <tr><td>40-49</td><td>C</td><td>2</td></tr>
                            <tr><td>33-39</td><td>D</td><td>1</td></tr>
                            <tr><td>0-32</td><td>F</td><td>0</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!-- Title Box -->
            <div class="pm-title-wrap">
                <div class="pm-title-box">অর্ধবাষিক মূল্যায়ন প্রতিবেদন-২০২৬</div>
            </div>

            <!-- Student Meta (2 Columns) -->
            <div class="pm-student-meta">
                <div class="pm-meta-left">
                    <div class="pm-meta-row">
                        <span class="pm-meta-label">শিক্ষার্থীর নাম :</span>
                        <span class="pm-meta-value">${student.name}</span>
                    </div>
                    <div class="pm-meta-row">
                        <span class="pm-meta-label">শিফট :</span>
                        <span class="pm-meta-value">দিবা</span>
                    </div>
                </div>
                <div class="pm-meta-right">
                    <div class="pm-meta-row">
                        <span class="pm-meta-label">শ্রেণি :</span>
                        <span class="pm-meta-value">${classTitle}</span>
                    </div>
                    <div class="pm-meta-row">
                        <span class="pm-meta-label">রোল :</span>
                        <span class="pm-meta-value">${displayRoll}</span>
                    </div>
                </div>
            </div>

            <!-- Main Marks Table -->
            <table class="pm-table">
                <thead>
                    <tr>
                        <th rowspan="2" style="width: 7%;">ক্রমিক<br>নং</th>
                        <th rowspan="2" style="width: 26%;">বিষয় এর নাম</th>
                        <th colspan="2" style="width: 20%;">শ্রেণিতে বিষয় ভিত্তিক<br>সর্বোচ্চ প্রাপ্ত নম্বর</th>
                        <th colspan="3" style="width: 29%;">প্রাপ্ত নম্বর</th>
                        <th rowspan="2" style="width: 9%;">অকৃতকার্য বিষয়<br>(অর্ধবাষিক)</th>
                        <th rowspan="2" style="width: 9%;">মেধাক্রম</th>
                    </tr>
                    <tr>
                        <th style="width: 10%;">১ম ক্লাস</th>
                        <th style="width: 10%;">অর্ধবার্ষিক</th>
                        <th style="width: 9%;">১ম ক্লাস</th>
                        <th style="width: 10%;">অর্ধবার্ষিক</th>
                        <th style="width: 10%;">মোট<br>নম্বর</th>
                    </tr>
                </thead>
                <tbody>
                    ${rowsHTML}
                    <tr>
                        <td colspan="6" class="total-row-label">সর্বমোট প্রাপ্ত নম্বর</td>
                        <td class="total-row-value">${toBn(student.totalObtained)}</td>
                    </tr>
                </tbody>
            </table>

            <!-- Signatures Section -->
            <div class="pm-signatures">
                <div class="pm-sig-box">
                    <div class="pm-sig-line"></div>
                    <div class="pm-sig-name">${examCommitteeName || '—'}</div>
                    <div class="pm-sig-title">অভ্যন্তরীন পরীক্ষা কমিটি</div>
                </div>
                <div class="pm-sig-box">
                    <div class="pm-sig-line"></div>
                    <div class="pm-sig-name">${t2}</div>
                    <div class="pm-sig-title">ফলাফল প্রস্তুতকারী</div>
                </div>
                <div class="pm-sig-box">
                    <div class="pm-sig-line"></div>
                    <div class="pm-sig-name">${t1}</div>
                    <div class="pm-sig-title">শ্রেণি শিক্ষক</div>
                </div>
                <div class="pm-sig-box">
                    <div class="pm-sig-line"></div>
                    <div class="pm-sig-name">${asstHeadName || '—'}</div>
                    <div class="pm-sig-title">সহকারী প্রধান শিক্ষক</div>
                </div>
                <div class="pm-sig-box">
                    <div class="pm-sig-line"></div>
                    <div class="pm-sig-name">সোহরাব উদ্দিন</div>
                    <div class="pm-sig-title">প্রধান শিক্ষক</div>
                </div>
            </div>
        </div>
    `;
}

// Download Single or Batch PDF
async function downloadPDF() {
    if (currentMode === 'single') {
        await downloadSinglePDF();
    } else {
        await downloadAllSectionPDF();
    }
}

// Download Single Student PDF with filename format: rollnum_class-section.pdf (e.g. 01_6-A.pdf)
async function downloadSinglePDF() {
    const elem = document.querySelector('.marksheet-card-item');
    if (!elem) {
        alert('কোনো মার্কশিট পাওয়া যায়নি!');
        return;
    }

    const downloadBtn = document.getElementById('pdf-download-btn');
    const origText = downloadBtn ? downloadBtn.innerHTML : '';
    if (downloadBtn) {
        downloadBtn.disabled = true;
        downloadBtn.innerHTML = `<span class="spinner"></span> PDF তৈরি হচ্ছে...`;
    }

    try {
        const canvas = await html2canvas(elem, {
            scale: 2,
            useCORS: true,
            allowTaint: true,
            backgroundColor: '#ffffff'
        });

        const imgData = canvas.toDataURL('image/png');
        const jsPDFLib = window.jspdf ? window.jspdf.jsPDF : window.jsPDF;
        const pdf = new jsPDFLib({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });
        
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

        const fileName = getPDFFileName(currentStudent);

        pdf.addImage(imgData, 'PNG', 0, 5, pdfWidth, pdfHeight);
        pdf.save(fileName);
    } catch (err) {
        console.error('PDF Generation error:', err);
        alert('PDF জেনারেট করতে সমস্যা হয়েছে: ' + err.message);
    } finally {
        if (downloadBtn) {
            downloadBtn.disabled = false;
            downloadBtn.innerHTML = origText;
        }
    }
}

// Progress bar helpers for batch PDF download
function showBatchProgress() {
    const container = document.getElementById('batch-progress-container');
    if (container) container.classList.remove('hidden');
}

function hideBatchProgress() {
    const container = document.getElementById('batch-progress-container');
    if (container) container.classList.add('hidden');
}

function updateBatchProgress(done, total, label) {
    const fill = document.getElementById('batch-progress-bar-fill');
    const text = document.getElementById('batch-progress-text');
    const percentEl = document.getElementById('batch-progress-percent');
    const percent = total > 0 ? Math.round((done / total) * 100) : 0;

    if (fill) fill.style.width = percent + '%';
    if (percentEl) percentEl.textContent = toBn(percent) + '%';

    if (text) {
        if (done < total) {
            text.textContent = `ডাউনলোড হচ্ছে: ${label} (${toBn(done + 1)}/${toBn(total)})`;
        } else {
            text.textContent = `✅ সম্পন্ন! মোট ${toBn(total)} টি মার্কশিট ডাউনলোড হয়েছে।`;
        }
    }
}

// Small delay helper so the browser has time to process each download
// before the next one starts (avoids browsers silently blocking rapid downloads)
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Download All Section Students' PDFs one-by-one (separate file per student)
// with a live progress bar showing current progress.
async function downloadAllSectionPDF() {
    // Render (and validate) the section's marksheets first
    const rendered = viewAllSectionStudents();
    if (!rendered) return;

    const selectedSec = document.getElementById('batch-section-select').value;
    const secStudents = currentBatchStudents;

    if (!secStudents || !secStudents.length) {
        alert(`শাখা ${selectedSec}-এ কোনো শিক্ষার্থী পাওয়া যায়নি।`);
        return;
    }

    // Give the DOM a moment to fully paint the newly rendered marksheet cards
    await delay(150);

    const cardElems = document.querySelectorAll('.marksheet-card-item');
    if (!cardElems.length) {
        alert('মার্কশিট রেন্ডার করতে সমস্যা হয়েছে, আবার চেষ্টা করুন।');
        return;
    }

    const batchDownloadBtn = document.getElementById('batch-download-btn');
    const batchViewBtn = document.getElementById('batch-view-btn');
    const pdfDownloadBtn = document.getElementById('pdf-download-btn');
    const origBatchBtnText = batchDownloadBtn ? batchDownloadBtn.innerHTML : '';

    if (batchDownloadBtn) {
        batchDownloadBtn.disabled = true;
        batchDownloadBtn.innerHTML = `<span class="spinner"></span> প্রসেসিং হচ্ছে...`;
    }
    if (batchViewBtn) batchViewBtn.disabled = true;
    if (pdfDownloadBtn) pdfDownloadBtn.disabled = true;

    const total = cardElems.length;
    showBatchProgress();
    updateBatchProgress(0, total, secStudents[0] ? secStudents[0].name : '');

    let successCount = 0;
    let failedStudents = [];

    try {
        const jsPDFLib = window.jspdf ? window.jspdf.jsPDF : window.jsPDF;

        for (let i = 0; i < cardElems.length; i++) {
            const elem = cardElems[i];
            const student = secStudents[i] || null;
            const label = student ? student.name : `#${i + 1}`;

            updateBatchProgress(i, total, label);

            try {
                const canvas = await html2canvas(elem, {
                    scale: 2,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff'
                });

                const imgData = canvas.toDataURL('image/png');
                const pdf = new jsPDFLib({
                    orientation: 'portrait',
                    unit: 'mm',
                    format: 'a4'
                });

                const pdfWidth = pdf.internal.pageSize.getWidth();
                const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

                pdf.addImage(imgData, 'PNG', 0, 5, pdfWidth, pdfHeight);

                const fileName = getPDFFileName(student);
                pdf.save(fileName);
                successCount++;
            } catch (innerErr) {
                console.error(`PDF generation failed for student index ${i}:`, innerErr);
                failedStudents.push(label);
            }

            updateBatchProgress(i + 1, total, label);

            // Small pause between downloads so the browser can process each
            // save before the next one starts (also avoids popup/download blocking)
            if (i < cardElems.length - 1) {
                await delay(600);
            }
        }

        if (failedStudents.length) {
            alert(`${successCount} টি মার্কশিট সফলভাবে ডাউনলোড হয়েছে।\nনিম্নলিখিত ${failedStudents.length} জনের মার্কশিট তৈরি করতে সমস্যা হয়েছে:\n${failedStudents.join(', ')}`);
        }
    } catch (err) {
        console.error('Batch PDF error:', err);
        alert('Batch PDF তৈরি করতে সমস্যা হয়েছে: ' + err.message);
    } finally {
        if (batchDownloadBtn) {
            batchDownloadBtn.disabled = false;
            batchDownloadBtn.innerHTML = origBatchBtnText;
        }
        if (batchViewBtn) batchViewBtn.disabled = false;
        if (pdfDownloadBtn) pdfDownloadBtn.disabled = false;

        // Keep the "done" state visible briefly, then hide the progress bar
        setTimeout(hideBatchProgress, 2500);
    }
}
