// PPT Automation Frontend JavaScript

let templates = [];
let slides = [];
let slideCount = 3;
let currentOutputId = null;

const API_BASE = '';

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    loadTemplates();
});

// Update Slide Count
function updateSlideCount() {
    slideCount = parseInt(document.getElementById('slideCount').value) || 3;
}

// Proceed to Slides Configuration
function proceedToSlides() {
    if (slideCount < 1) {
        alert('Please select a valid number of slides');
        return;
    }
    initializeSlides();
    showStep('step-slides');
}

// Proceed to Template Step
function proceedToTemplate() {
    // Validate all slides have files
    let allHaveFiles = true;
    for (let i = 0; i < slides.length; i++) {
        if (!slides[i].file_id) {
            alert(`Slide ${i + 1} needs an Excel file uploaded`);
            allHaveFiles = false;
            break;
        }
    }
    
    if (allHaveFiles) {
        showStep('step-template');
    }
}

// Proceed to Generate
function proceedToGenerate() {
    const templateSelect = document.getElementById('templateSelect');
    if (!templateSelect.value) {
        alert('Please select a template');
        return;
    }
    showStep('step-generate');
}

// Initialize Slides Based on Count
function initializeSlides() {
    slides = [];
    for (let i = 0; i < slideCount; i++) {
        slides.push({
            slide_number: i + 1,
            slide_type: 'table',
            title: `Slide ${i + 1}`,
            file_id: null,
            file_name: null,
            file_analysis: null,
            sheet: '',
            columns: [],
            header_row: 0,
            data_preview: null
        });
    }
    renderSlides();
}

// Upload File for Specific Slide
async function uploadFileForSlide(slideIndex) {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx,.xlsb,.xls';
    fileInput.style.display = 'none';
    
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        await handleSlideFileUpload(slideIndex, file);
        document.body.removeChild(fileInput);
    });
    
    document.body.appendChild(fileInput);
    fileInput.click();
}

// Handle File Upload for Slide
async function handleSlideFileUpload(slideIndex, file) {
    const slide = slides[slideIndex];
    const uploadArea = document.getElementById(`uploadArea-${slideIndex}`);
    
    uploadArea.innerHTML = '<p>⏳ Uploading and analyzing...</p>';
    showProgress();
    
    try {
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch(`${API_BASE}/api/analyze-excel`, {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            throw new Error('Failed to analyze file');
        }
        
        const analysis = await response.json();
        
        // Update slide with file info
        slide.file_id = analysis.file_id;
        slide.file_name = analysis.filename;
        slide.file_analysis = analysis;
        
        // Set default sheet
        if (analysis.sheets.length > 0) {
            slide.sheet = analysis.sheets[0].name;
        }
        
        // Load data preview
        await loadDataPreview(slideIndex);
        
        // Re-render slide
        renderSlides();
        checkAllSlidesReady();
        
        hideProgress();
        
    } catch (error) {
        uploadArea.innerHTML = `<p style="color: red;">Error: ${error.message}</p>`;
        hideProgress();
    }
}

// Load Data Preview for Slide
async function loadDataPreview(slideIndex) {
    const slide = slides[slideIndex];
    if (!slide.file_id || !slide.sheet) return;
    
    try {
        const response = await fetch(`${API_BASE}/api/excel-columns?file_id=${slide.file_id}&sheet=${encodeURIComponent(slide.sheet)}`);
        const data = await response.json();
        
        // Get sample data from analysis
        const fileAnalysis = slide.file_analysis;
        const sheetInfo = fileAnalysis.sheets.find(s => s.name === slide.sheet);
        
        if (sheetInfo && sheetInfo.sample_data) {
            slide.data_preview = {
                columns: data.columns,
                sample_rows: sheetInfo.sample_data.slice(0, 5)
            };
        }
        
    } catch (error) {
        console.error('Error loading preview:', error);
    }
}

// Load Templates
async function loadTemplates() {
    try {
        const response = await fetch(`${API_BASE}/api/templates`);
        const data = await response.json();
        templates = data.templates;
        
        const select = document.getElementById('templateSelect');
        select.innerHTML = '<option value="">Select a template...</option>';
        
        if (templates.length === 0) {
            select.innerHTML += '<option value="templates/template.pptx">Default Template</option>';
        } else {
            templates.forEach(template => {
                const option = document.createElement('option');
                option.value = template.path;
                option.textContent = template.name;
                select.appendChild(option);
            });
        }
        
        select.addEventListener('change', function() {
            if (this.value) {
                showStep('step-slides');
                if (slides.length === 0) {
                    addSlide();
                }
            }
        });
        
    } catch (error) {
        console.error('Error loading templates:', error);
    }
}

// No need for addSlide/removeLastSlide - slides are fixed based on count

// Render Slides Configuration
function renderSlides() {
    const container = document.getElementById('slidesContainer');
    container.innerHTML = '';
    
    slides.forEach((slide, index) => {
        const slideDiv = document.createElement('div');
        slideDiv.className = 'slide-config';
        slideDiv.innerHTML = `
            <div class="slide-header">
                <h3>📊 Slide ${slide.slide_number}</h3>
                ${slide.file_id ? '<span class="badge badge-success">✓ File Uploaded</span>' : '<span class="badge badge-warning">⚠ No File</span>'}
            </div>
            
            <div class="form-group">
                <label>Upload Excel File for This Slide</label>
                <div class="upload-area-small" id="uploadArea-${index}" onclick="uploadFileForSlide(${index})">
                    ${slide.file_id ? `
                        <div class="file-uploaded">
                            <p>📁 <strong>${slide.file_name}</strong></p>
                            <p class="file-info-text">${slide.file_analysis ? slide.file_analysis.sheets.length + ' sheet(s)' : ''}</p>
                            <button class="btn btn-small" onclick="event.stopPropagation(); uploadFileForSlide(${index})">Change File</button>
                        </div>
                    ` : `
                        <div class="upload-content-small">
                            <p>📁 Click to upload Excel file</p>
                            <p class="upload-hint">.xlsx, .xlsb, or .xls</p>
                        </div>
                    `}
                </div>
            </div>
            
            ${slide.file_id ? `
                <div class="form-group">
                    <label>Slide Type</label>
                    <select class="form-control" onchange="updateSlide(${index}, 'slide_type', this.value)">
                        <option value="title" ${slide.slide_type === 'title' ? 'selected' : ''}>Title Slide</option>
                        <option value="table" ${slide.slide_type === 'table' ? 'selected' : ''}>Table</option>
                        <option value="content" ${slide.slide_type === 'content' ? 'selected' : ''}>Content</option>
                    </select>
                </div>
                
                <div class="form-group">
                    <label>Title</label>
                    <input type="text" class="form-control" value="${slide.title || ''}" 
                           onchange="updateSlide(${index}, 'title', this.value)">
                </div>
                
                ${slide.slide_type === 'table' ? renderTableConfig(slide, index) : ''}
                
                ${slide.data_preview ? renderDataPreview(slide, index) : ''}
            ` : ''}
        `;
        
        container.appendChild(slideDiv);
    });
}

// Render Data Preview
function renderDataPreview(slide, index) {
    if (!slide.data_preview) return '';
    
    const preview = slide.data_preview;
    let html = `
        <div class="data-preview">
            <h4>📋 Data Preview</h4>
            <div class="preview-table">
                <table>
                    <thead>
                        <tr>
                            ${preview.columns.map(col => `<th>${col.name}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${preview.sample_rows.map(row => `
                            <tr>
                                ${preview.columns.map(col => `<td>${row[col.name] !== undefined ? row[col.name] : ''}</td>`).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
            <p class="preview-note">Showing first 5 rows</p>
        </div>
    `;
    
    return html;
}

// Check if All Slides Ready
function checkAllSlidesReady() {
    const allReady = slides.every(slide => slide.file_id);
    if (allReady) {
        document.getElementById('proceedToTemplateBtn').style.display = 'block';
    } else {
        document.getElementById('proceedToTemplateBtn').style.display = 'none';
    }
}

// Render Table Configuration
function renderTableConfig(slide, index) {
    if (!slide.file_id || !slide.file_analysis) return '';
    
    const fileAnalysis = slide.file_analysis;
    
    let html = `
        <div class="form-group">
            <label>Sheet</label>
            <select class="form-control" onchange="updateSlideSheet(${index}, this.value)">
                ${fileAnalysis.sheets.map(s => 
                    `<option value="${s.name}" ${s.name === slide.sheet ? 'selected' : ''}>${s.name} (${s.row_count} rows)</option>`
                ).join('')}
            </select>
        </div>
        
        <div class="form-group">
            <label>Header Row (0-indexed)</label>
            <input type="number" class="form-control" value="${slide.header_row || 0}" 
                   onchange="updateSlide(${index}, 'header_row', parseInt(this.value))">
        </div>
        
        <div class="form-group">
            <label>Select Columns to Include</label>
            <div id="columns-${index}" class="column-selector">
                <p>Loading columns...</p>
            </div>
        </div>
    `;
    
    // Load columns for current sheet
    if (slide.sheet) {
        setTimeout(() => loadColumns(index, slide.file_id, slide.sheet), 100);
    }
    
    return html;
}

// Update Slide Sheet
async function updateSlideSheet(slideIndex, sheetName) {
    const slide = slides[slideIndex];
    slide.sheet = sheetName;
    slide.columns = [];
    
    await loadDataPreview(slideIndex);
    await loadColumns(slideIndex, slide.file_id, sheetName);
    
    renderSlides();
}

// Load Columns for a Sheet
async function loadColumns(slideIndex, fileId, sheetName) {
    if (!fileId) return;
    
    try {
        const response = await fetch(`${API_BASE}/api/excel-columns?file_id=${fileId}&sheet=${encodeURIComponent(sheetName)}`);
        const data = await response.json();
        
        const container = document.getElementById(`columns-${slideIndex}`);
        if (!container) return;
        
        let html = '<div class="column-checkboxes">';
        data.columns.forEach(col => {
            const checked = slides[slideIndex].columns.includes(col.name) ? 'checked' : '';
            html += `
                <div class="checkbox-item">
                    <input type="checkbox" id="col-${slideIndex}-${col.name}" ${checked}
                           onchange="toggleColumn(${slideIndex}, '${col.name}', this.checked)">
                    <label for="col-${slideIndex}-${col.name}">${col.name}</label>
                </div>
            `;
        });
        html += '</div>';
        
        container.innerHTML = html;
        
    } catch (error) {
        console.error('Error loading columns:', error);
    }
}

// Toggle Column Selection
function toggleColumn(slideIndex, columnName, checked) {
    const slide = slides[slideIndex];
    if (checked) {
        if (!slide.columns.includes(columnName)) {
            slide.columns.push(columnName);
        }
    } else {
        slide.columns = slide.columns.filter(c => c !== columnName);
    }
}

// Update Slide Property
function updateSlide(index, property, value) {
    slides[index][property] = value;
    
    // Re-render if slide type changed
    if (property === 'slide_type') {
        renderSlides();
    }
}

// Generate PPT
async function generatePPT() {
    const templateSelect = document.getElementById('templateSelect');
    const templatePath = templateSelect.value || 'templates/template.pptx';
    
    if (slides.length === 0) {
        showError('Please configure at least one slide');
        return;
    }
    
    // Validate all slides have file selected
    for (let i = 0; i < slides.length; i++) {
        if (!slides[i].file_id) {
            showError(`Slide ${i + 1} needs an Excel file uploaded`);
            return;
        }
        if (slides[i].slide_type === 'table' && slides[i].columns.length === 0) {
            showError(`Slide ${i + 1} needs at least one column selected`);
            return;
        }
    }
    
    const generateBtn = document.getElementById('generateBtn');
    generateBtn.disabled = true;
    generateBtn.textContent = '⏳ Generating...';
    
    showProgress();
    showStatus('info', 'Generating PowerPoint deck... This may take a moment.');
    
    try {
        // Prepare slides config with file mappings
        const slidesConfig = slides.map(slide => {
            const config = {
                slide_number: slide.slide_number,
                slide_type: slide.slide_type,
                title: slide.title,
                file_id: slide.file_id,
                data_source: slide.file_name.replace(/\.(xlsx|xlsb|xls)$/i, ''),
                sheet: slide.sheet,
                columns: slide.columns,
                header_row: slide.header_row
            };
            return config;
        });
        
        // Prepare uploaded files info
        const uploadedFilesInfo = {};
        slides.forEach(slide => {
            if (slide.file_id && !uploadedFilesInfo[slide.file_id]) {
                uploadedFilesInfo[slide.file_id] = {
                    name: slide.file_name,
                    analysis: slide.file_analysis
                };
            }
        });
        
        const response = await fetch(`${API_BASE}/api/generate-ppt`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                uploaded_files: uploadedFilesInfo,
                template_path: templatePath,
                slides_config: slidesConfig
            })
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Failed to generate PPT');
        }
        
        const result = await response.json();
        currentOutputId = result.output_id;
        
        showStatus('success', 'PowerPoint deck generated successfully!');
        showDownload();
        hideProgress();
        
    } catch (error) {
        showError('Error generating PPT: ' + error.message);
        hideProgress();
    } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = '🚀 Generate PowerPoint Deck';
    }
}

// Show Download Section
function showDownload() {
    const downloadSection = document.getElementById('downloadSection');
    const downloadLink = document.getElementById('downloadLink');
    
    downloadLink.href = `${API_BASE}/api/download/${currentOutputId}`;
    downloadSection.style.display = 'block';
}

// Utility Functions
function showStep(stepId) {
    document.getElementById(stepId).style.display = 'block';
}

function showProgress() {
    document.getElementById('progressBar').style.display = 'block';
}

function hideProgress() {
    document.getElementById('progressBar').style.display = 'none';
}

function showStatus(type, message) {
    const statusDiv = document.getElementById('generateStatus');
    statusDiv.className = `status-message ${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
}

function showError(message) {
    showStatus('error', message);
}

