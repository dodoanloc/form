// Form App - JavaScript
// =========================================

// DOM Elements
const form = document.getElementById('businessForm');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    console.log('Form App initialized');
});

// Form Submission
form.addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const submitBtn = form.querySelector('button[type="submit"]');
    submitBtn.classList.add('loading');
    
    try {
        // Collect form data
        const formData = new FormData(form);
        const data = Object.fromEntries(formData.entries());
        
        // Map form fields to DOCX placeholders
        const mappings = {
            'TenVN': data.businessName || '',
            'SoGT': data.businessLicense || '',
            'Ngcap': data.businessIssueDate || '',
            'NoiCap': data.businessIssuePlace || '',
            'Ng.HH': data.ownerExpiryDate || '',
            'MST': data.ownerIdNumber || '',
            'TenChiNhanh': '',
            'TenTA': '',
            'TenVTVn': '',
            'TenVTTA': '',
            'MsKH': '',
            'SoTK': ''
        };
        
        console.log('Mapped data:', mappings);
        
        // Generate DOCX from template
        await generateDOCXFromTemplate(mappings);
        
        showToast('Xuất file thành công!', 'success');
    } catch (error) {
        console.error('Error:', error);
        showToast('Có lỗi xảy ra: ' + error.message, 'error');
    } finally {
        submitBtn.classList.remove('loading');
    }
});

// Reset Form
function resetForm() {
    form.reset();
    showToast('Đã nhập lại form', 'success');
}

// Generate DOCX from template
async function generateDOCXFromTemplate(data) {
    try {
        // Fetch the template
        const templateUrl = 'mau_bieu.docx';
        const response = await fetch(templateUrl);
        if (!response.ok) {
            throw new Error('Không tải được template');
        }
        
        const arrayBuffer = await response.arrayBuffer();
        
        // Use PizZip from window
        const PizZip = window.PizZip;
        const Docxtemplater = window.Docxtemplater;
        
        if (!PizZip) {
            throw new Error('Thư viện PizZip chưa tải');
        }
        if (!Docxtemplater) {
            throw new Error('Thư viện Docxtemplater chưa tải');
        }
        
        const zip = new PizZip(arrayBuffer);
        
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
        });
        
        // Render the document
        doc.render(data);
        
        // Get the generated file
        const out = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        
        // Download
        const url = URL.createObjectURL(out);
        const a = document.createElement('a');
        a.href = url;
        a.download = `mau_bieu_ho_kinh_doanh_${Date.now()}.docx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Error generating DOCX:', error);
        throw error;
    }
}

// Show Toast
function showToast(message, type = 'success') {
    const existingToast = document.querySelector('.toast');
    if (existingToast) {
        existingToast.remove();
    }
    
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
        ${type === 'success' 
            ? '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>'
            : '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>'
        }
        <span>${message}</span>
    `;
    
    document.body.appendChild(toast);
    
    setTimeout(() => {
        toast.classList.add('show');
    }, 10);
    
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

window.resetForm = resetForm;