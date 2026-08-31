/* ==========================================================================
   TAQA HR PORTAL - APPLICATION LOGIC (PWA / Vanilla JS + Firebase + SheetJS)
   ========================================================================== */

// --- INITIAL SAMPLE EMPLOYEE DATABASE ---
const DEFAULT_EMPLOYEE_DATABASE = [];

// --- APP STATE ---
let employeeDatabase = [];
let submissionsList = [];
let currentEmployee = null;
let dbFirebase = null;

// --- INITIALIZATION ---
document.addEventListener('DOMContentLoaded', () => {
    // Initialize Lucide Icons
    if (window.lucide) {
        lucide.createIcons();
    }

    // Load Database & Submissions from LocalStorage
    loadEmployeeDatabase();
    loadSubmissions();

    // Initialize Firebase if Config exists
    initFirebaseIfAvailable();

    // Update UI Stats
    updateStatsUI();

    // Setup Live Cross-Tab Instant Synchronization (Zero Refresh)
    setupCrossTabSync();
});

function setupCrossTabSync() {
    if (window.BroadcastChannel) {
        try {
            const channel = new BroadcastChannel('taqa_hr_portal_sync');
            channel.onmessage = (event) => {
                if (event.data && event.data.type) {
                    loadSubmissions();
                    loadEmployeeDatabase();
                    renderSubmissionsTable();
                    updateStatsUI();
                }
            };
        } catch (e) {}
    }

    window.addEventListener('storage', (e) => {
        if (e.key === 'taqa_submissions' || e.key === 'taqa_employee_db') {
            loadSubmissions();
            loadEmployeeDatabase();
            renderSubmissionsTable();
            updateStatsUI();
        }
    });
}

function broadcastDataChange() {
    if (window.BroadcastChannel) {
        try {
            const channel = new BroadcastChannel('taqa_hr_portal_sync');
            channel.postMessage({ type: 'DATA_UPDATED', timestamp: Date.now() });
        } catch (e) {}
    }
}

/* ==========================================================================
   1. EMPLOYEE DATABASE MANAGEMENT
   ========================================================================== */

function loadEmployeeDatabase() {
    const storedDb = localStorage.getItem('taqa_employee_db');
    if (storedDb) {
        try {
            const parsed = JSON.parse(storedDb);
            // Purge old sample database if it contained the 6 test employees
            if (Array.isArray(parsed) && parsed.some(e => e.id === "100001" || e.id === "1001" || e.nameAr === "أحمد محمود علي السيد")) {
                localStorage.removeItem('taqa_employee_db');
                employeeDatabase = [];
            } else {
                employeeDatabase = parsed;
            }
        } catch (e) {
            console.error("Error parsing stored DB", e);
            employeeDatabase = [];
        }
    } else {
        employeeDatabase = [];
    }
}

function saveEmployeeDatabase() {
    localStorage.setItem('taqa_employee_db', JSON.stringify(employeeDatabase));
    updateStatsUI();
    broadcastDataChange();
}

function handleExcelDbUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });
            
            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                alert("عذراً! الملف لا يحتوي على أي صفحات (Sheets).");
                return;
            }

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (!jsonSheet || jsonSheet.length === 0) {
                alert("شيت الإكسيل المرفوع فارغ ولا يحتوي على بيانات!");
                return;
            }

            // Get column headers from first row
            const firstRowKeys = Object.keys(jsonSheet[0] || {});

            // Helper for fuzzy key matching
            const getVal = (row, candidates) => {
                for (const key of Object.keys(row)) {
                    const cleanKey = key.trim().toLowerCase();
                    for (const candidate of candidates) {
                        if (cleanKey.includes(candidate.toLowerCase())) {
                            return String(row[key] || "").trim();
                        }
                    }
                }
                return "";
            };

            const parsedEmployees = jsonSheet.map((row, index) => {
                // 1. Employee ID candidates (Matching exact header 'Employee ID' from screenshot)
                let id = getVal(row, ["Employee ID", "EmployeeID", "الرقم الوظيفي", "رقم الموظف", "كود الموظف", "كود", "الرقم", "empid", "emp_id", "emp id", "code", "id"]);
                
                // Fallback to first column if it looks like an ID
                if (!id && firstRowKeys[0] && row[firstRowKeys[0]]) {
                    const candidateVal = String(row[firstRowKeys[0]]).trim();
                    if (/^\d+$/.test(candidateVal)) {
                        id = candidateVal;
                    }
                }

                // 2. Arabic Name candidates (Matching exact header 'الاسم عربي' from screenshot)
                let nameAr = getVal(row, ["الاسم عربي", "اسم عربي", "الاسم بالعربي", "الاسم باللغة العربية", "الاسم الثلاثي", "الاسم الرباعي", "اسم الموظف", "الاسم", "namear", "arabic name", "ar name"]);

                // 3. English Name candidates (Matching exact header 'Employee Name' from screenshot)
                let nameEn = getVal(row, ["Employee Name", "EmployeeName", "الاسم بالإنجليزي", "الاسم باللغة الإنجليزية", "english name", "nameen", "en name", "english"]);

                return { id, nameAr, nameEn };
            }).filter(emp => emp.id && (emp.nameAr || emp.nameEn));

            if (parsedEmployees.length > 0) {
                employeeDatabase = parsedEmployees;
                saveEmployeeDatabase();

                // Sync to Firebase if connected
                if (dbFirebase) {
                    syncEmployeeDbToFirebase(parsedEmployees);
                }

                alert(`تم استيراد ${parsedEmployees.length} موظف بنجاح إلى قاعدة البيانات! ✨`);
            } else {
                const keysList = firstRowKeys.join("، ");
                alert(`تعذر المطابقة التلقائية لأعمدة الإكسيل.\nالأعمدة الموجودة بالشيت هي: (${keysList})\nيرجى التأكد من اسم عمود الرقم الوظيفي والاسم.`);
            }
        } catch (err) {
            console.error("Excel Read Error:", err);
            alert("حدث خطأ أثناء قراءة ملف Excel: " + (err.message || "يرجى التثبت من صيغة الملف .xlsx أو .xls"));
        }
    };

    reader.onerror = function() {
        alert("حدث خطأ أثناء فتح الملف من الجهاز.");
    };

    reader.readAsArrayBuffer(file);
}

/* ==========================================================================
   2. STEP FORM & EMPLOYEE VERIFICATION
   ========================================================================== */

function handleIdVerification(event) {
    if (event) event.preventDefault();
    const idInput = document.getElementById('employeeIdInput').value.trim();
    const errorBox = document.getElementById('step1Error');
    const errorText = document.getElementById('step1ErrorText');

    errorBox.classList.add('hidden');

    if (!idInput) {
        errorText.innerText = "برجاء كتابة الرقم الوظيفي أولاً.";
        errorBox.classList.remove('hidden');
        return;
    }

    // Always fetch latest submissions state from LocalStorage before checking
    loadSubmissions();

    // Search in employeeDatabase with smart zero-padding matching (e.g. 000084 vs 84)
    const cleanInput = idInput.replace(/^0+/, '');
    const emp = employeeDatabase.find(e => {
        const rawId = String(e.id).trim();
        const cleanId = rawId.replace(/^0+/, '');
        return rawId === idInput || (cleanInput !== "" && cleanId === cleanInput);
    });

    if (emp) {
        currentEmployee = emp;
        
        // Check if employee has ALREADY submitted an agreement (Smart zero-padding match: e.g. 000084 vs 84)
        const existingSubmission = submissionsList.find(s => {
            const rawSubId = String(s.empId).trim();
            const cleanSubId = rawSubId.replace(/^0+/, '');
            return rawSubId === idInput || (cleanInput !== "" && cleanSubId === cleanInput);
        });

        if (existingSubmission) {
            currentEmployee = null;
            errorText.innerText = "عذراً! تم تسجيل وإرسال الإقرار مسبقاً لهذا الرقم الوظيفي، ولا يُسمح بإعادة التقديم مرة أخرى.";
            errorBox.classList.remove('hidden');
            return;
        }

        // Populate Step 2 Profile
        document.getElementById('displayEmpId').innerText = emp.id;
        document.getElementById('displayArabicName').innerText = emp.nameAr || "غير مسجل";
        document.getElementById('displayEnglishName').innerText = emp.nameEn || "Not Registered";

        // Reset step 2 inputs
        document.getElementById('nationalIdInput').value = "";
        document.getElementById('nationalIdCounter').innerText = "0 / 14";
        document.getElementById('agree1').checked = false;
        document.getElementById('agree2').checked = false;
        document.getElementById('agree3').checked = false;

        // Transition to Step 2
        goToStep(2);
    } else {
        errorText.innerText = `الرقم الوظيفي (${idInput}) غير مسجل في قاعدة البيانات. برجاء التأكد والتمس المصادقة مع الأدمن.`;
        errorBox.classList.remove('hidden');
    }
}

function clearEmployeeId() {
    document.getElementById('employeeIdInput').value = "";
    document.getElementById('step1Error').classList.add('hidden');
    document.getElementById('employeeIdInput').focus();
}

function validateNationalInput(input) {
    // Keep digits only
    input.value = input.value.replace(/[^0-9]/g, '');
    const currentLength = input.value.length;
    const counter = document.getElementById('nationalIdCounter');
    counter.innerText = `${currentLength} / 14`;

    if (currentLength === 14) {
        counter.style.color = 'var(--accent-blue)';
    } else {
        counter.style.color = 'var(--text-muted)';
    }
}

function handleFinalSubmission(event) {
    event.preventDefault();
    const nationalId = document.getElementById('nationalIdInput').value.trim();
    const errorBox = document.getElementById('step2Error');
    const errorText = document.getElementById('step2ErrorText');
    const agree1 = document.getElementById('agree1').checked;
    const agree2 = document.getElementById('agree2').checked;
    const agree3 = document.getElementById('agree3').checked;
    const submitBtn = document.getElementById('submitBtn');

    errorBox.classList.add('hidden');

    // Check if current employee already submitted
    if (currentEmployee) {
        const currRawId = String(currentEmployee.id).trim();
        const currCleanId = currRawId.replace(/^0+/, '');
        const existing = submissionsList.find(s => {
            const rawSubId = String(s.empId).trim();
            const cleanSubId = rawSubId.replace(/^0+/, '');
            return rawSubId === currRawId || (currCleanId !== "" && cleanSubId === currCleanId);
        });

        if (existing) {
            errorText.innerText = "عذراً! لقد تم إرسال وتأكيد الإقرار لهذا الرقم الوظيفي مسبقاً، ولا يُسمح بالتكرار إطلاقاً.";
            errorBox.classList.remove('hidden');
            return;
        }
    }

    if (nationalId.length !== 14) {
        errorText.innerText = "عذراً! الرقم القومي يجب أن يتكون من 14 رقم بشكل دقيق.";
        errorBox.classList.remove('hidden');
        return;
    }

    if (!agree1 || !agree2 || !agree3) {
        errorText.innerText = "برجاء تحديد جميع الموافقات والتعهدات الرسمية للمتابعة.";
        errorBox.classList.remove('hidden');
        return;
    }

    // Disable button immediately to prevent double submissions
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<span>جارٍ الحفظ والإرسال...</span> <i data-lucide="check"></i>';
    }

    const now = new Date();
    const timestampFormatted = now.toLocaleString('ar-EG', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
    });

    const submissionData = {
        id: Date.now().toString(),
        empId: currentEmployee.id,
        nameAr: currentEmployee.nameAr,
        nameEn: currentEmployee.nameEn,
        nationalId: nationalId,
        submittedAt: timestampFormatted,
        timestampIso: now.toISOString(),
        agreementsConfirmed: true
    };

    // Save Submission safely
    try {
        saveSubmissionRecord(submissionData);
    } catch (e) {
        console.error("Storage save error:", e);
    }

    // Populate Summary Screen (Step 3)
    document.getElementById('summaryEmpId').innerText = submissionData.empId;
    document.getElementById('summaryEmpName').innerText = submissionData.nameAr;
    document.getElementById('summaryNationalId').innerText = submissionData.nationalId;
    document.getElementById('summaryTimestamp').innerText = submissionData.submittedAt;

    const titleElem = document.getElementById('successTitle');
    const descElem = document.getElementById('successDesc');
    if (titleElem) titleElem.innerText = "تم إرسال بياناتك بنجاح! ✓";
    if (descElem) descElem.innerText = "تم تسجيل إقرارك في النظام وتوثيق الوقت والتاريخ بنجاح. شكرًا لتعاونكم.";

    // Restore submit button state for future uses
    if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<span>إرسال البيانات والتأكيد</span> <i data-lucide="send"></i>';
        if (window.lucide) lucide.createIcons();
    }

    // Transition to Step 3 instantly
    goToStep(3);
}

function goToStep(stepNumber) {
    // Hide all steps
    document.getElementById('step1').classList.add('hidden');
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.add('hidden');

    // Remove active/completed classes from indicators
    const step1Ind = document.getElementById('stepIndicator1');
    const step2Ind = document.getElementById('stepIndicator2');
    const step3Ind = document.getElementById('stepIndicator3');
    const line1 = document.getElementById('stepLine1');
    const line2 = document.getElementById('stepLine2');

    step1Ind.className = "step-item";
    step2Ind.className = "step-item";
    step3Ind.className = "step-item";
    line1.className = "step-line";
    line2.className = "step-line";

    if (stepNumber === 1) {
        document.getElementById('step1').classList.remove('hidden');
        step1Ind.classList.add('active');
    } else if (stepNumber === 2) {
        document.getElementById('step2').classList.remove('hidden');
        step1Ind.classList.add('completed');
        line1.classList.add('active');
        step2Ind.classList.add('active');
    } else if (stepNumber === 3) {
        document.getElementById('step3').classList.remove('hidden');
        step1Ind.classList.add('completed');
        line1.classList.add('active');
        step2Ind.classList.add('completed');
        line2.classList.add('active');
        step3Ind.classList.add('active');
    }
}

function resetForm() {
    currentEmployee = null;
    document.getElementById('employeeIdInput').value = "";
    document.getElementById('step1Error').classList.add('hidden');
    document.getElementById('step2Error').classList.add('hidden');
    
    const submitBtn = document.getElementById('submitBtn');
    if (submitBtn) {
        submitBtn.disabled = false;
        submitBtn.innerHTML = '<span>إرسال البيانات والتأكيد</span> <i data-lucide="send"></i>';
        if (window.lucide) lucide.createIcons();
    }

    goToStep(1);
}

/* ==========================================================================
   3. SUBMISSIONS LOCAL & FIREBASE STORAGE
   ========================================================================== */

function loadSubmissions() {
    const stored = localStorage.getItem('taqa_submissions');
    if (stored) {
        try {
            submissionsList = JSON.parse(stored);
        } catch (e) {
            submissionsList = [];
        }
    }
    renderSubmissionsTable();
}

function saveSubmissionRecord(record) {
    // Save locally
    submissionsList.unshift(record);
    localStorage.setItem('taqa_submissions', JSON.stringify(submissionsList));

    // Broadcast instant sync across tabs
    broadcastDataChange();

    // Save to Firebase Firestore if initialized
    if (dbFirebase) {
        dbFirebase.collection('submissions').add(record)
            .then(() => console.log("Successfully synced submission to Firebase"))
            .catch(err => console.error("Firebase sync error:", err));
    }

    renderSubmissionsTable();
    updateStatsUI();
}

function clearAllSubmissionsData() {
    if (confirm("هل أنت تأكد من رغبتك في حذف جميع الإقرارات المسجلة؟")) {
        submissionsList = [];
        localStorage.removeItem('taqa_submissions');
        renderSubmissionsTable();
        updateStatsUI();
    }
}

/* ==========================================================================
   4. ADMIN PANEL & TABLE & EXCEL EXPORT
   ========================================================================== */

function toggleAdminModal() {
    const modal = document.getElementById('adminModal');
    modal.classList.toggle('hidden');
    if (!modal.classList.contains('hidden')) {
        renderSubmissionsTable();
        updateStatsUI();
    }
}

function renderSubmissionsTable(filterText = "") {
    const tbody = document.getElementById('submissionsTableBody');
    if (!tbody) return; // Safeguard when running on index.html page
    tbody.innerHTML = "";

    const filtered = submissionsList.filter(item => {
        if (!filterText) return true;
        const q = filterText.toLowerCase();
        return (
            item.empId.toLowerCase().includes(q) ||
            item.nameAr.toLowerCase().includes(q) ||
            item.nameEn.toLowerCase().includes(q) ||
            item.nationalId.includes(q)
        );
    });

    if (filtered.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" style="text-align: center; color: var(--text-muted); padding: 2rem;">
                    لا توجد إقرارات مسجلة حتى الآن.
                </td>
            </tr>
        `;
        return;
    }

    filtered.forEach((item, idx) => {
        const tr = document.createElement('tr');
        const uniqueId = item.id || item.empId;
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><strong style="color: var(--accent-cyan); font-family: var(--font-en);">${item.empId}</strong></td>
            <td>${item.nameAr || '-'}</td>
            <td dir="ltr" style="text-align: right; font-family: var(--font-en);">${item.nameEn || '-'}</td>
            <td><code style="font-family: var(--font-en); font-size: 0.9rem; color: var(--accent-blue);">${item.nationalId}</code></td>
            <td style="font-size: 0.82rem; color: var(--text-secondary);">${item.submittedAt}</td>
            <td><span class="status-badge">✓ مؤكد وموافق</span></td>
            <td>
                <button 
                    type="button" 
                    onclick="deleteSingleSubmission('${uniqueId}')" 
                    style="background: #fee2e2; color: #dc2626; border: 1px solid #fca5a5; padding: 0.25rem 0.65rem; border-radius: 6px; cursor: pointer; font-size: 0.8rem; font-weight: 600; display: inline-flex; align-items: center; gap: 0.25rem; transition: all 0.2s;"
                    onmouseover="this.style.background='#fca5a5'"
                    onmouseout="this.style.background='#fee2e2'"
                    title="مسح إقرار الموظف"
                >
                    <i data-lucide="trash-2" style="width: 14px; height: 14px;"></i>
                    <span>حذف</span>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    if (window.lucide) lucide.createIcons();
}

function deleteSingleSubmission(submissionId) {
    const item = submissionsList.find(s => (s.id && String(s.id) === String(submissionId)) || String(s.empId) === String(submissionId));
    if (!item) return;

    const empIdToDelete = String(item.empId).trim();
    const cleanIdToDelete = empIdToDelete.replace(/^0+/, '');
    const confirmMsg = `هل أنت تأكد من رغبتك في حذف إقرار الموظف (${item.nameAr || item.empId})؟\n\nتنويه: سيتم مسح الإقرار نهائياً وإتاحة إمكانية التسجيل مرة أخرى للموظف.`;

    if (confirm(confirmMsg)) {
        // 1. Also clear from deleted blacklist if present
        removeDeletedSubmission(empIdToDelete);

        // 2. Remove ALL matching submission records for this employee ID (smart zero-padding match)
        submissionsList = submissionsList.filter(s => {
            const rawSubId = String(s.empId).trim();
            const cleanSubId = rawSubId.replace(/^0+/, '');
            const rawItemId = String(s.id || '').trim();
            
            const isMatch = (rawSubId === empIdToDelete) ||
                            (cleanIdToDelete !== "" && cleanSubId === cleanIdToDelete) ||
                            (rawItemId !== "" && rawItemId === String(submissionId));
            return !isMatch;
        });

        // 3. Save updated list to LocalStorage
        localStorage.setItem('taqa_submissions', JSON.stringify(submissionsList));

        // 4. Delete matching documents from Firebase Firestore
        if (dbFirebase) {
            if (item.id) {
                dbFirebase.collection('submissions').doc(String(item.id)).delete().catch(() => {});
            }

            dbFirebase.collection('submissions').get()
                .then(snapshot => {
                    snapshot.forEach(doc => {
                        const data = doc.data();
                        const docEmpId = String(data.empId || '').trim();
                        const cleanDocEmpId = docEmpId.replace(/^0+/, '');
                        if (docEmpId === empIdToDelete || (cleanIdToDelete !== "" && cleanDocEmpId === cleanIdToDelete)) {
                            doc.ref.delete();
                        }
                    });
                })
                .catch(err => console.error("Firebase doc delete error:", err));
        }

        renderSubmissionsTable();
        updateStatsUI();
        broadcastDataChange();
    }
}

function filterSubmissionsTable() {
    const q = document.getElementById('tableSearchInput').value.trim();
    renderSubmissionsTable(q);
}

function exportSubmissionsToExcel() {
    if (submissionsList.length === 0) {
        alert("لا توجد إقرارات لتصديرها!");
        return;
    }

    // Format data cleanly for Excel columns
    const exportData = submissionsList.map((item, index) => ({
        "م": index + 1,
        "الرقم الوظيفي": item.empId,
        "الاسم بالعربي": item.nameAr,
        "الاسم بالإنجليزي": item.nameEn,
        "الرقم القومي (14 رقم)": item.nationalId,
        "تاريخ ووقت التقديم": item.submittedAt,
        "حالة الموافقة": "تمت الموافقة على جميع التعهدات"
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    
    // Auto-fit column widths
    const colWidths = [
        { wch: 5 },
        { wch: 15 },
        { wch: 30 },
        { wch: 30 },
        { wch: 22 },
        { wch: 25 },
        { wch: 30 }
    ];
    worksheet['!cols'] = colWidths;

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "الإقرارات المستلمة");

    const fileName = `إقرارات_الموظفين_طاقة_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
}

function updateStatsUI() {
    document.getElementById('empDbCount').innerText = employeeDatabase.length;
    document.getElementById('submissionsCount').innerText = submissionsList.length;
}

/* ==========================================================================
   5. FIREBASE DYNAMIC INTEGRATION
   ========================================================================== */

const DEFAULT_FIREBASE_CONFIG = {
  apiKey: "AIzaSyCeeB1vYO0WxhJA3A5G9GwtvLOHGoxMtLQ",
  authDomain: "project-hr-taqa.firebaseapp.com",
  projectId: "project-hr-taqa",
  storageBucket: "project-hr-taqa.firebasestorage.app",
  messagingSenderId: "304602057029",
  appId: "1:304602057029:web:aec731b889faf8ef4fd62c",
  measurementId: "G-FWE7DZ0T8Z"
};

function toggleFirebaseConfigModal() {
    const modal = document.getElementById('firebaseConfigModal');
    modal.classList.toggle('hidden');
}

function saveFirebaseConfig() {
    const config = {
        apiKey: document.getElementById('fbApiKey').value.trim(),
        authDomain: document.getElementById('fbAuthDomain').value.trim(),
        projectId: document.getElementById('fbProjectId').value.trim(),
        storageBucket: document.getElementById('fbStorageBucket').value.trim(),
        messagingSenderId: document.getElementById('fbMessagingSenderId').value.trim(),
        appId: document.getElementById('fbAppId').value.trim()
    };

    if (!config.apiKey || !config.projectId) {
        alert("يرجى ملء المفاتيح الأساسية (apiKey و projectId) على الأقل.");
        return;
    }

    localStorage.setItem('taqa_firebase_config', JSON.stringify(config));
    alert("تم حفظ مفاتيح Firebase بنجاح! جاري الاتصال...");
    toggleFirebaseConfigModal();
    initFirebaseIfAvailable();
}

function initFirebaseIfAvailable() {
    const pill = document.getElementById('firebaseStatusPill');
    const text = document.getElementById('firebaseStatusText');

    let config = DEFAULT_FIREBASE_CONFIG;
    const savedConfig = localStorage.getItem('taqa_firebase_config');
    if (savedConfig) {
        try {
            config = JSON.parse(savedConfig);
        } catch (e) {
            config = DEFAULT_FIREBASE_CONFIG;
        }
    }

    if (!config || !config.apiKey || !config.projectId) {
        if (pill && text) {
            pill.querySelector('.status-dot').className = "status-dot";
            text.innerText = "Firebase غير مهيأ (يرجى إدخال المفاتيح من الإعدادات)";
        }
        return;
    }

    try {
        // Fill modal fields for convenience
        if (document.getElementById('fbApiKey')) {
            document.getElementById('fbApiKey').value = config.apiKey || "";
            document.getElementById('fbAuthDomain').value = config.authDomain || "";
            document.getElementById('fbProjectId').value = config.projectId || "";
            document.getElementById('fbStorageBucket').value = config.storageBucket || "";
            document.getElementById('fbMessagingSenderId').value = config.messagingSenderId || "";
            document.getElementById('fbAppId').value = config.appId || "";
        }

        if (window.firebase && !firebase.apps.length) {
            firebase.initializeApp(config);
        }
        
        if (window.firebase) {
            dbFirebase = firebase.firestore();

            // Update UI pill status
            if (pill && text) {
                pill.querySelector('.status-dot').className = "status-dot success";
                text.innerText = `Firebase متصل بنجاح: (${config.projectId})`;
            }

            // Listen live for cloud submissions if available
            listenToFirebaseSubmissions();
            listenToFirebaseEmployeeRegistry();
        }

    } catch (e) {
        console.error("Firebase Init Error:", e);
        if (pill && text) {
            pill.querySelector('.status-dot').className = "status-dot warning";
            text.innerText = "خطأ في الاتصال بـ Firebase (يرجى مراجعة الإعدادات)";
        }
    }
}

function listenToFirebaseSubmissions() {
    if (!dbFirebase) return;
    
    dbFirebase.collection('submissions').onSnapshot((snapshot) => {
        const remoteSubmissions = [];
        snapshot.forEach(doc => {
            remoteSubmissions.push(doc.data());
        });

        if (remoteSubmissions.length > 0) {
            // Sort by date desc
            remoteSubmissions.sort((a, b) => new Date(b.timestampIso || 0) - new Date(a.timestampIso || 0));

            // Merge with local list avoiding duplicates
            const map = new Map();
            [...remoteSubmissions, ...submissionsList].forEach(item => {
                if (item.id && !map.has(item.id)) {
                    map.set(item.id, item);
                }
            });
            submissionsList = Array.from(map.values());
            localStorage.setItem('taqa_submissions', JSON.stringify(submissionsList));
            renderSubmissionsTable();
            updateStatsUI();
        }
    }, (err) => {
        console.warn("Firestore live snapshot error / fallback to local storage:", err);
    });
}

function listenToFirebaseEmployeeRegistry() {
    if (!dbFirebase) return;

    dbFirebase.collection('employee_registry').onSnapshot((snapshot) => {
        const remoteEmployees = [];
        snapshot.forEach(doc => {
            remoteEmployees.push(doc.data());
        });

        if (remoteEmployees.length > 0) {
            employeeDatabase = remoteEmployees;
            saveEmployeeDatabase();
        }
    }, (err) => {
        console.warn("Firestore employee registry listener fallback:", err);
    });
}

function syncEmployeeDbToFirebase(dbArray) {
    if (!dbFirebase) return;
    const batch = dbFirebase.batch();
    dbArray.forEach(emp => {
        const ref = dbFirebase.collection('employee_registry').doc(emp.id);
        batch.set(ref, emp);
    });
    batch.commit()
        .then(() => console.log("Uploaded employee registry batch to Firestore"))
        .catch(err => console.error("Firestore DB batch error:", err));
}

