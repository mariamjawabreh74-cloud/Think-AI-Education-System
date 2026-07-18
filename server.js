require('dotenv').config();
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const axios = require('axios');

const app = express();
const upload = multer({ dest: 'uploads/' });
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

if (!fs.existsSync(path.join(__dirname, 'uploads'))) {
    fs.mkdirSync(path.join(__dirname, 'uploads'));
}

const QUESTIONS_FILE = path.join(__dirname, 'questions.json');
const ADMINS_FILE = path.join(__dirname, 'admins.json');
const ARCHIVE_META_FILE = path.join(__dirname, 'archive.json');
const ARCHIVE_DIR = path.join(__dirname, 'uploads', 'archive');
if (!fs.existsSync(ARCHIVE_DIR)) {
    fs.mkdirSync(ARCHIVE_DIR, { recursive: true });
}
app.use('/uploads/archive', express.static(ARCHIVE_DIR));

function saveToArchive(schoolId, type, req, rowCount) {
    const archive = readJSON(ARCHIVE_META_FILE);
    const list = Array.isArray(archive.items) ? archive.items : [];
    const id = `arc_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
    const storedName = `${id}${path.extname(req.file.originalname) || '.xlsx'}`;
    fs.copyFileSync(req.file.path, path.join(ARCHIVE_DIR, storedName));
    list.push({
        id,
        schoolId,
        type, // 'students' أو 'teachers'
        originalName: req.file.originalname,
        storedName,
        rowCount,
        uploadedAt: new Date().toISOString()
    });
    writeJSON(ARCHIVE_META_FILE, { items: list });
}

function getQuestions() {
    if (!fs.existsSync(QUESTIONS_FILE)) return [];
    return JSON.parse(fs.readFileSync(QUESTIONS_FILE, 'utf8'));
}
function saveQuestions(questions) {
    fs.writeFileSync(QUESTIONS_FILE, JSON.stringify(questions, null, 4));
}
function readJSON(filePath) {
    if (!fs.existsSync(filePath)) return {};
    try {
        return JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch (error) {
        return {};
    }
}
function writeJSON(filePath, data) {
    fs.writeFileSync(filePath, JSON.stringify(data, null, 4));
}
app.post('/api/superadmin/register-school', (req, res) => {
    const providedSecret = req.body.secret || req.headers['x-super-admin-secret'];
    if (!process.env.SUPER_ADMIN_SECRET) {
        return res.status(500).json({ success: false, message: 'لم يتم ضبط SUPER_ADMIN_SECRET على السيرفر بعد. راجع البرمجي المسؤول.' });
    }
    if (!providedSecret || providedSecret !== process.env.SUPER_ADMIN_SECRET) {
        return res.status(403).json({ success: false, message: 'المفتاح السري غير صحيح.' });
    }

    const { adminId, adminName, schoolName } = req.body;
    if (!adminId || !adminName || !schoolName) {
        return res.status(400).json({ success: false, message: 'البيانات غير مكتملة (adminId, adminName, schoolName).' });
    }

    const admins = readJSON(ADMINS_FILE);
    const idStr = String(adminId).trim();
    if (admins[idStr]) {
        return res.status(400).json({ success: false, message: 'رقم هوية المدير/ة هذا مسجل مسبقاً.' });
    }

    const schoolId = `sch_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
    admins[idStr] = {
        name: String(adminName).trim(),
        schoolName: String(schoolName).trim(),
        schoolId: schoolId,
        role: 'admin'
    };
    writeJSON(ADMINS_FILE, admins);

    res.json({ success: true, message: 'تم إنشاء المدرسة والمدير/ة بنجاح.', schoolId: schoolId });
});
app.post('/api/superadmin/upload-schools',
upload.single('excelFile'),
(req,res)=>{

    const providedSecret=req.headers['x-super-admin-secret'];

    if(providedSecret!==process.env.SUPER_ADMIN_SECRET){
        return res.status(403).json({
            success:false,
            message:'غير مصرح.'
        });
    }

    if(!req.file){
        return res.status(400).json({
            success:false,
            message:'لم يتم اختيار ملف.'
        });
    }

    try{

        const workbook=xlsx.readFile(req.file.path);

        const sheet=workbook.Sheets[workbook.SheetNames[0]];

        const rows=xlsx.utils.sheet_to_json(sheet);

        const admins=readJSON(ADMINS_FILE);

        rows.forEach(row=>{

            const schoolId="sch_"+Date.now()+"_"+Math.random().toString(36).substr(2,5);

            admins[String(row.رقم_الهوية)]={
                name:row.اسم_المديرة,
                schoolName:row.اسم_المدرسة,
                schoolId,
                role:"admin"
            };

        });

        writeJSON(ADMINS_FILE,admins);

        fs.unlinkSync(req.file.path);

        res.json({
            success:true,
            message:"تم رفع الملف بنجاح."
        });

    }catch(err){

        res.status(500).json({
            success:false,
            message:err.message
        });

    }

});

app.get('/api/superadmin/schools', (req, res) => {
    const providedSecret = req.query.secret || req.headers['x-super-admin-secret'];
    if (!process.env.SUPER_ADMIN_SECRET) {
        return res.status(500).json({ success: false, message: 'لم يتم ضبط SUPER_ADMIN_SECRET على السيرفر بعد. راجع البرمجي المسؤول.' });
    }
    if (!providedSecret || providedSecret !== process.env.SUPER_ADMIN_SECRET) {
        return res.status(403).json({ success: false, message: 'المفتاح السري غير صحيح.' });
    }


    const admins = readJSON(ADMINS_FILE);
    const schools = Object.entries(admins).map(([adminId, info]) => ({
        adminId,
        adminName: info.name,
        schoolName: info.schoolName,
        schoolId: info.schoolId
    }));
    res.json({ success: true, schools: schools });
});

app.post('/upload-students', upload.single('excelFile'), (req, res) => {
    if (!req.file) return res.status(400).send('لم يتم إرفاق أي ملف.');
    const schoolId = req.body.schoolId;
    if (!schoolId) { fs.unlinkSync(req.file.path); return res.status(400).send('لم يتم تحديد المدرسة (schoolId)، الرجاء تسجيل الدخول من جديد.'); }
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        const filePath = path.join(__dirname, 'students.json');

        const studentsDatabase = readJSON(filePath);
        jsonData.forEach(student => {
            studentsDatabase[String(student.رقم_الهوية).trim()] = {
                name: String(student.الاسم).trim(),
                grade: String(student.الصف).trim(),
                role: 'student',
                schoolId: schoolId
            };
        });
         
        writeJSON(filePath, studentsDatabase);
        saveToArchive(schoolId, 'students', req, jsonData.length);
        fs.unlinkSync(req.file.path);
        res.send(`تم إضافة/تحديث ${jsonData.length} طالب بنجاح (تمت إضافة الملف للأرشيف)!`);
    } catch (error) {
        res.status(500).send('حدث خطأ أثناء قراءة ملف الطلاب: ' + error.message);
    }
});
  app.delete('/api/superadmin/delete-school', (req, res) => {

    const providedSecret = req.headers['x-super-admin-secret'];

    if (!providedSecret || providedSecret !== process.env.SUPER_ADMIN_SECRET) {
        return res.status(403).json({
            success: false,
            message: 'المفتاح السري غير صحيح.'
        });
    }

    const { adminId } = req.body;

    const admins = readJSON(ADMINS_FILE);

    if (!admins[adminId]) {
        return res.status(404).json({
            success: false,
            message: 'المدرسة غير موجودة.'
        });
    }

    const schoolId = admins[adminId].schoolId;

    delete admins[adminId];
    writeJSON(ADMINS_FILE, admins);

    const studentsFile = path.join(__dirname, 'students.json');
    if (fs.existsSync(studentsFile)) {
        const students = readJSON(studentsFile);

        Object.keys(students).forEach(id => {
            if (students[id].schoolId === schoolId) {
                delete students[id];
            }
        });

        writeJSON(studentsFile, students);
    }

    const teachersFile = path.join(__dirname, 'teachers.json');
    if (fs.existsSync(teachersFile)) {
        const teachers = readJSON(teachersFile);

        Object.keys(teachers).forEach(id => {
            if (teachers[id].schoolId === schoolId) {
                delete teachers[id];
            }
        });

        writeJSON(teachersFile, teachers);
    }

    let questions = getQuestions();
    questions = questions.filter(q => q.schoolId !== schoolId);
    saveQuestions(questions);

    const archive = readJSON(ARCHIVE_META_FILE);

    const remain = [];

    (archive.items || []).forEach(item => {

        if (item.schoolId === schoolId) {

            const file = path.join(ARCHIVE_DIR, item.storedName);

            if (fs.existsSync(file)) {
                fs.unlinkSync(file);
            }

        } else {
            remain.push(item);
        }

    });

    writeJSON(ARCHIVE_META_FILE, {
        items: remain
    });

    res.json({
        success: true,
        message: 'تم حذف المدرسة وجميع بياناتها.'
    });


});
app.post('/upload-teachers', upload.single('excelFile'), (req, res) => {
    if (!req.file) return res.status(400).send('لم يتم إرفاق أي ملف.');
    const schoolId = req.body.schoolId;
    if (!schoolId) { fs.unlinkSync(req.file.path); return res.status(400).send('لم يتم تحديد المدرسة (schoolId)، الرجاء تسجيل الدخول من جديد.'); }
    try {
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        const filePath = path.join(__dirname, 'teachers.json');

        const teachersDatabase = readJSON(filePath);
        jsonData.forEach(teacher => {
            teachersDatabase[String(teacher.رقم_الهوية).trim()] = {
                name: String(teacher.الاسم).trim(),
                classes: String(teacher.المواد_والصفوف).trim(),
                role: 'teacher',
                schoolId: schoolId
            };
        });
        writeJSON(filePath, teachersDatabase);
        saveToArchive(schoolId, 'teachers', req, jsonData.length);
        fs.unlinkSync(req.file.path);
        res.send(`تم إضافة/تحديث ${jsonData.length} معلم/ة بنجاح (تمت إضافة الملف للأرشيف)!`);
    } catch (error) {
        res.status(500).send('حدث خطأ أثناء قراءة ملف المعلمين' + error.message);
    }
});

app.get('/api/admin/archive', (req, res) => {
    const { schoolId, type } = req.query;
    if (!schoolId) return res.status(400).json({ success: false, message: 'schoolId مطلوب.' });
    const archive = readJSON(ARCHIVE_META_FILE);
    let list = Array.isArray(archive.items) ? archive.items : [];
    list = list.filter(item => item.schoolId === schoolId && (!type || item.type === type));
    list.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
    res.json({ success: true, items: list });
});

app.delete('/api/admin/archive/:id', (req, res) => {
    const { id } = req.params;
    const { schoolId } = req.body;
    if (!schoolId) return res.status(400).json({ success: false, message: 'schoolId مطلوب.' });

    const archive = readJSON(ARCHIVE_META_FILE);
    const list = Array.isArray(archive.items) ? archive.items : [];
    const index = list.findIndex(item => item.id === id && item.schoolId === schoolId);
    if (index === -1) return res.status(404).json({ success: false, message: 'الملف غير موجود.' });

    const [removed] = list.splice(index, 1);
    const storedPath = path.join(ARCHIVE_DIR, removed.storedName);
    if (fs.existsSync(storedPath)) fs.unlinkSync(storedPath);
    writeJSON(ARCHIVE_META_FILE, { items: list });

    res.json({ success: true, message: 'تم حذف الملف من الأرشيف.' });
});

app.post('/api/login', (req, res) => {
    const { nationalId, role } = req.body;
    if (!nationalId || !role) return res.status(400).json({ success: false, message: 'الرجاء إدخال البيانات كاملة.' });
    const idStr = String(nationalId).trim();

    if (role === 'admin') {
        const admins = readJSON(ADMINS_FILE);
        if (admins[idStr]) {
            return res.json({ success: true, redirect: 'admin.html', name: admins[idStr].name, schoolId: admins[idStr].schoolId });
        }
        return res.status(400).json({ success: false, message: 'رقم هوية المدير/ة غير صحيح.' });
    }

    if (role === 'student') {
        const filePath = path.join(__dirname, 'students.json');
        if (!fs.existsSync(filePath)) return res.status(400).json({ success: false, message: 'لم يتم رفع بيانات الطلاب بعد من قبل الإدارة.' });
        const studentsData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        if (studentsData[idStr]) {
            return res.json({ success: true, redirect: 'student.html', name: studentsData[idStr].name, grade: studentsData[idStr].grade, schoolId: studentsData[idStr].schoolId });
        }
    }

    if (role === 'teacher') {
        const filePath = path.join(__dirname, 'teachers.json');
        if (!fs.existsSync(filePath)) return res.status(400).json({ success: false, message: 'لم يتم رفع بيانات المعلمين بعد من قبل الإدارة.' });
        const teachersData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        if (teachersData[idStr]) {
            return res.json({ success: true, redirect: 'teacher.html', name: teachersData[idStr].name, classes: teachersData[idStr].classes, schoolId: teachersData[idStr].schoolId });
        }
    }
    res.status(400).json({ success: false, message: 'عذراً، رقم الهوية غير مسجل في القسم المختار.' });
});

app.use((req,res,next)=>{
    console.log(req.method, req.url);
    next();
});

app.post('/api/questions/send', upload.single('questionImage'), async(req, res) => {

    console.log("وصل سؤال من الطالب");
    const { text, subject, grade, schoolId } = req.body;

    // التأكد من البيانات أولاً
    if (!text || !subject || !grade || !schoolId) {
        return res.status(400).json({
            success: false,
            message: 'البيانات غير مكتملة.'
        });
    }


    // إرسال السؤال إلى نموذج الذكاء الاصطناعي
    const aiResponse = await axios.post(
        "http://127.0.0.1:5000/predict",
        {
            text: text
        }
    );

    console.log("AI RESPONSE:", aiResponse.data);


    // إذا رفض الموديل السؤال نوقف الحفظ
    if (aiResponse.data.result === "rejected") {
        return res.json({
            success: false,
            message: "تم رفض السؤال من نظام الذكاء الاصطناعي."
        });
    }


    const questions = getQuestions();

    const newQuestion = {
        id: Date.now().toString(),
        text: text,
        imageUrl: req.file ? `/uploads/${req.file.filename}` : null,
        subject: subject,
        grade: grade,
        schoolId: schoolId,
        status: 'pending',
        answer: null,
        teacherName: null
    };

    questions.push(newQuestion);

    saveQuestions(questions);

    res.json({
        success: true,
        message: 'تم إرسال سؤالك بنجاح وبسرية تامة!'
    });

});
const cleanStr = (t) => t ? String(t).trim().toLowerCase().replace(/^ال/, '') : '';

app.get('/api/subjects-by-grade', (req, res) => {
    const { grade, schoolId } = req.query;
    if (!grade) return res.status(400).json({ success: false, message: 'الصف مطلوب.' });
    if (!schoolId) return res.status(400).json({ success: false, message: 'المدرسة (schoolId) مطلوبة.' });

    const filePath = path.join(__dirname, 'teachers.json');
    if (!fs.existsSync(filePath)) return res.json([]);

    let teachersData;
    try {
        teachersData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch (error) {
        return res.status(500).json({ success: false, message: 'خطأ في قراءة بيانات المعلمين.' });
    }

    const clean = (t) => {
        if (!t) return '';
        let str = String(t).trim().toLowerCase().replace(/[^a-zA-Z0-9آ-ي]/g, '');
        if (str.startsWith('ال')) str = str.substring(2);
        return str;
    };
    const targetGrade = clean(grade);

    const subjectsSet = new Set();

    Object.values(teachersData)
        .filter(teacher => teacher.schoolId === schoolId)
        .forEach(teacher => {
            if (!teacher.classes) return;
            const rooms = String(teacher.classes).split(',');
            rooms.forEach(room => {
                if (!room.includes('-')) return;
                const parts = room.split('-');
                const subject = parts[0].trim();
                const teacherGrade = clean(parts[1]);
                if (!subject || !teacherGrade) return;

                const matchGrade = teacherGrade === targetGrade || teacherGrade.includes(targetGrade) || targetGrade.includes(teacherGrade);
                if (matchGrade) subjectsSet.add(subject);
            });
        });

    res.json(Array.from(subjectsSet));
});

app.get('/api/questions/pending', (req, res) => {
    const { grade, subject, schoolId } = req.query;
    const questions = getQuestions();
    
    const clean = (text) => {
        if (!text) return "";
        let str = String(text).trim().toLowerCase().replace(/[^a-zA-Z0-9آ-ي]/g, '');
        if (str.startsWith("ال")) str = str.substring(2);
        return str;
    };

    const filtered = questions.filter(q => {
        const isPending = q.status === 'pending';
        const matchSchool = q.schoolId === schoolId;
        const matchSubject = clean(q.subject) === clean(subject);
        
        const qGrade = clean(q.grade);
        const tGrade = clean(grade);
        const matchGrade = qGrade === tGrade || qGrade.includes(tGrade) || tGrade.includes(qGrade);
        
        return isPending && matchSchool && matchSubject && matchGrade;
    });

    res.json(filtered);
});


app.post('/api/questions/answer', (req, res) => {
    const { questionId, answer, teacherName, schoolId } = req.body;
    let questions = getQuestions();
    const idx = questions.findIndex(q => q.id === questionId);
    if (idx !== -1) {
        if (schoolId && questions[idx].schoolId !== schoolId) {
            return res.status(403).json({ success: false, message: 'لا يمكنك الإجابة على سؤال من مدرسة أخرى.' });
        }
        questions[idx].status = 'answered';
        questions[idx].answer = answer;
        questions[idx].teacherName = teacherName;
        saveQuestions(questions);
        return res.json({ success: true, message: 'تم نشر الإجابة بنجاح!' });
    }
    res.status(404).json({ success: false, message: 'السؤال غير موجود.' });
});

app.get('/api/questions/approved', (req, res) => {
    const { grade, subject, schoolId } = req.query;
    const questions = getQuestions();
    const filtered = questions.filter(q => q.status === 'answered' && q.schoolId === schoolId && cleanStr(q.grade) === cleanStr(grade) && cleanStr(q.subject) === cleanStr(subject));
    res.json(filtered);
});
app.get('/api/admin/all-questions', (req, res) => {
    try {
        const { schoolId } = req.query;
        const questions = getQuestions();
        const filtered = schoolId ? questions.filter(q => q.schoolId === schoolId) : [];
        res.json(filtered);
    } catch (error) {
        res.status(500).json({ success: false, message: 'خطأ في جلب البيانات' });
    }
});
app.get('/api/admin/filter-options', (req, res) => {

    const { schoolId } = req.query;

    if (!schoolId)
        return res.json({
            grades: [],
            subjects: []
        });

    const teachersFile = path.join(__dirname, 'teachers.json');

    if (!fs.existsSync(teachersFile))
        return res.json({
            grades: [],
            subjects: []
        });

    const teachers = readJSON(teachersFile);

    const grades = new Set();
    const subjects = new Set();

    Object.values(teachers).forEach(t => {

        if (t.schoolId !== schoolId) return;

        if (!t.classes) return;

        String(t.classes).split(',').forEach(room => {

            if (!room.includes('-')) return;

            const parts = room.split('-');

            const subject = parts[0].trim();
            const grade = parts[1].trim();

            if (subject)
                subjects.add(subject);

            if (grade)
                grades.add(grade);

        });

    });

    res.json({
        grades: [...grades],
        subjects: [...subjects]
    });

});
app.listen(3000, () => console.log(' السيرفر يعمل: http://localhost:3000'));