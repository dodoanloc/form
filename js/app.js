const form = document.getElementById('businessForm');
const toggleOptionalBtn = document.getElementById('toggleOptionalBtn');
const optionalFields = document.getElementById('optionalFields');

function pad2(value) {
    return String(value).padStart(2, '0');
}

function formatDateSlash(value) {
    if (!value) return '';
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(value)) return value;
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return value;
    return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
}

function formatDateFullText(value) {
    const d = value ? new Date(value) : new Date();
    if (Number.isNaN(d.getTime())) return value || '';
    return `ngày ${pad2(d.getDate())} tháng ${pad2(d.getMonth() + 1)} năm ${d.getFullYear()}`;
}

function todayIso() {
    const d = new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function normalizeVietnameseUpper(text) {
    if (!text) return '';
    return text
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/Đ/g, 'D')
        .toUpperCase();
}

function fallback(value, defaultValue = '') {
    return value && String(value).trim() ? String(value).trim() : defaultValue;
}

function xmlEscape(text) {
    return String(text ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function resetForm() {
    form.reset();
    showToast('Đã nhập lại form', 'success');
}

function toggleOptionalFields() {
    optionalFields.classList.toggle('is-open');
    const isOpen = optionalFields.classList.contains('is-open');
    toggleOptionalBtn.textContent = isOpen
        ? 'Ẩn bớt trường không bắt buộc'
        : 'Hiện thêm trường không bắt buộc';
}

function buildMappings(data) {
    const businessName = fallback(data.businessName);
    const ownerIdNumber = fallback(data.ownerIdNumber);
    const ownerPhone = fallback(data.ownerPhone);
    const businessAddress = fallback(data.businessAddress);
    const businessNameEnDefault = normalizeVietnameseUpper(businessName);
    const today = todayIso();

    return {
        'Ngaydaydu': formatDateFullText(today),
        'Ngay//': formatDateSlash(today),
        'MsKH': fallback(data.customerCode, '................'),
        'SoTK': fallback(data.accountNumber, '................'),
        'TenVN': businessName,
        'TenVTVn': businessName,
        'TenTA': businessNameEnDefault,
        'TenVTTA': businessNameEnDefault,
        'Loai Giấy Tờ': 'ĐKKD',
        'SoGT': fallback(data.businessLicense),
        'MST': fallback(data.taxCode, ownerIdNumber),
        'Ngcap': formatDateSlash(data.businessIssueDate),
        'NoiCap': fallback(data.businessIssuePlace, '................'),
        'Ng.HH': '',
        'Lv.KD': fallback(data.businessField, '................'),
        'Cty.M': '',
        'Nc.KD': '',
        'DC': businessAddress,
        'HoTen.GD': fallback(data.ownerName),
        'GT.GD': fallback(data.ownerGender),
        'NgS.GD': formatDateSlash(data.ownerBirthDate),
        'QT.GD': '',
        'SoCC.GD': ownerIdNumber,
        'NgC.GD': formatDateSlash(data.ownerIssueDate),
        'NC.GD': fallback(data.ownerIssuePlace),
        'NgHH.GD': formatDateSlash(data.ownerExpiryDate),
        'DC.GD': fallback(data.ownerAddress, businessAddress),
        'SoDT.GD': ownerPhone,
        'CV.GD': '',
        'NgC.TL': '',
        'QDBN.TL': '......................',
        'DT': fallback(data.contactPhone, ownerPhone),
        'Em.Cty': fallback(data.companyEmail, '................'),
        'TenChiNhanh': fallback(data.branchName, 'Thọ Xuân Thanh Hóa'),
        'Diadanh': fallback(data.locationName, 'Thanh Hóa'),
        'GDV': fallback(data.staffCode, '')
    };
}

function replacePlaceholdersInXml(xml, mappings) {
    let output = xml;
    for (const [key, value] of Object.entries(mappings)) {
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`\\[${escapedKey}\\]`, 'g');
        output = output.replace(regex, xmlEscape(value));
    }
    return output;
}

async function generateDocxFromTemplate(mappings) {
    const response = await fetch('mau_bieu.docx');
    if (!response.ok) throw new Error('Không tải được file mẫu DOCX');

    const arrayBuffer = await response.arrayBuffer();
    const PizZip = window.PizZip;
    if (!PizZip) throw new Error('Thư viện PizZip chưa tải');

    const zip = new PizZip(arrayBuffer);
    const xmlTargets = Object.keys(zip.files).filter(name =>
        name.startsWith('word/') &&
        name.endsWith('.xml') &&
        !name.includes('theme/')
    );

    xmlTargets.forEach((name) => {
        const file = zip.file(name);
        if (!file) return;
        const content = file.asText();
        const replaced = replacePlaceholdersInXml(content, mappings);
        zip.file(name, replaced);
    });

    const blob = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `mau_bieu_ho_kinh_doanh_${Date.now()}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function showToast(message, type = 'success') {
    const existingToast = document.querySelector('.toast');
    if (existingToast) existingToast.remove();

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
        ${type === 'success'
            ? '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>'
            : '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>'}
        <span>${message}</span>
    `;

    document.body.appendChild(toast);
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const submitBtn = form.querySelector('button[type="submit"]');
    submitBtn.classList.add('loading');

    try {
        const formData = new FormData(form);
        const data = Object.fromEntries(formData.entries());
        const mappings = buildMappings(data);
        await generateDocxFromTemplate(mappings);
        showToast('Xuất file thành công!', 'success');
    } catch (error) {
        console.error(error);
        showToast(`Có lỗi xảy ra: ${error.message}`, 'error');
    } finally {
        submitBtn.classList.remove('loading');
    }
});

if (toggleOptionalBtn) {
    toggleOptionalBtn.addEventListener('click', toggleOptionalFields);
}

window.resetForm = resetForm;