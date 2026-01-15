const xlsx = require('xlsx');

// Normalization Helpers
const normalizeCourseCode = (code) => {
    if (!code) return '';
    // Replace comma with pipe as requested (AI3001,BAI-5A -> AI3001|BAI-5A)
    // Also ensuring it's trimmed
    return code.toString().replace(/,/g, '|').trim();
};

const normalizeStudentId = (id) => {
    if (!id) return '';
    // Remove hyphens (20p-0048 -> 20p0048) and lowercase?
    // User example: 20p-0048 -> 20p0048. It kept the 'p' lowercase.
    return id.toString().replace(/-/g, '').trim();
};

const parseExcel = (buffer) => {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const result = {
        courses: {},
        slots: [],
        enrollments: {},
        counts: {
            courses: 0,
            slots: 0,
            students: 0
        },
        errors: []
    };

    // --- Sheet 1: Courses ---
    const sheet1Name = workbook.SheetNames[0]; // Assuming order or specific names? User said "Sheet-1: Courses"
    // Better to find sheet by name if possible, or assume 1st sheet.
    // User Prompt: "Sheet-1: Courses", "Sheet-2: Slots", "Sheet-3: Enrollments"
    // I will try to match loosely or fallback to index.
    const coursesSheet = workbook.Sheets['Courses'] || workbook.Sheets[sheet1Name];
    const coursesData = xlsx.utils.sheet_to_json(coursesSheet);

    coursesData.forEach((row, index) => {
        // Expected: code, batch, subject, teacher
        const rawCode = row.code;
        if (!rawCode) {
            result.errors.push(`Row ${index + 2} in Courses: Missing code`);
            return;
        }

        const normalizedKey = normalizeCourseCode(rawCode);

        result.courses[normalizedKey] = {
            subject: row.subject,
            teacher: row.teacher
        };
    });

    result.counts.courses = Object.keys(result.courses).length;

    // --- Sheet 2: Slots ---
    const sheet2Name = workbook.SheetNames[1];
    const slotsSheet = workbook.Sheets['Slots'] || workbook.Sheets[sheet2Name];
    const slotsData = xlsx.utils.sheet_to_json(slotsSheet);

    slotsData.forEach((row, index) => {
        // Expected: day, location, slot, code, Slots (time)
        // Output: day, slot, code, room, time
        const rawCode = row.code;
        if (!rawCode) {
            // It's possible to have slots without code? Maybe logic to ignore or error.
            // For now, let's error if code is missing, or skip?
            // Prompt doesn't specify partial slots. Assuming valid.
            result.errors.push(`Row ${index + 2} in Slots: Missing code`);
            return;
        }

        const normalizedCode = normalizeCourseCode(rawCode);

        // Validation: Slot codes exist in Courses
        if (!result.courses[normalizedCode]) {
            result.errors.push(`Row ${index + 2} in Slots: Code '${normalizedCode}' not found in Courses`);
            // We still parse it, or check? "Validate: Slot codes exist in Courses... Store parsed data".
            // Usually if validation fails, we might still store or reject.
            // "Return: ... errors: []".
            // I will add to errors and skip adding to valid slots? Or add anyway?
            // "Validation... Store parsed data". I'll skip invalid ones to keep data integrity clean?
            // Or just log error. Logic: "Slot codes exist in Courses". Implies strict dependency.
            return;
        }

        result.slots.push({
            day: row.day,
            slot: row.slot,
            code: normalizedCode,
            room: row.location, // Mapping location -> room
            time: row.Slots     // Mapping Slots -> time
        });
    });

    result.counts.slots = result.slots.length;

    // --- Sheet 3: Enrollments ---
    const sheet3Name = workbook.SheetNames[2];
    const enrollSheet = workbook.Sheets['Enrollments'] || workbook.Sheets[sheet3Name];
    const enrollData = xlsx.utils.sheet_to_json(enrollSheet);

    enrollData.forEach((row, index) => {
        // Expected: rollnumber, code
        const rawRoll = row.rollnumber;
        const rawCode = row.code;

        if (!rawRoll || !rawCode) return;

        const normalizedRoll = normalizeStudentId(rawRoll);
        const normalizedCode = normalizeCourseCode(rawCode);

        // Validation: Enrollment codes exist in Courses
        if (!result.courses[normalizedCode]) {
            result.errors.push(`Row ${index + 2} in Enrollments: Code '${normalizedCode}' not found in Courses`);
            return;
        }

        if (!result.enrollments[normalizedRoll]) {
            result.enrollments[normalizedRoll] = [];
        }

        // Avoid duplicates if necessary, but array is requested.
        result.enrollments[normalizedRoll].push(normalizedCode);
    });

    result.counts.students = Object.keys(result.enrollments).length;

    return result;
};

module.exports = { parseExcel };
