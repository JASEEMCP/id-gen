'use client';

import React, { useState, useRef, useEffect } from 'react';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';
import * as XLSX from 'xlsx';

// --- Types ---
interface DynamicField {
  id: string;
  label: string;
  value: string;
}

interface IdData {
  headerTitle: string;
  headerSubtitle: string;
  headerBanner: string;
  photoUrl: string | null;
  signatureUrl: string | null;
  topSpaceHeight: number;
  cardWidth: number;
  cardHeight: number;
  pdfCardWidthMm: number;
  pdfCardHeightMm: number;
  pdfMarginX: number;
  pdfMarginY: number;
  pdfGridCols: number;
  pdfGridRows: number;
  headerColor: string;
  bodyColor: string;
  headerTextColor: string;
  bannerTextColor: string;
  bodyTextColor: string;
  dynamicFields: DynamicField[];
}

// ... Components ...

const ECILogo = () => (
  <img 
    src="/eci-logo.png" 
    alt="ECI Logo" 
    style={{ width: '140%', height: '140%', objectFit: 'contain'}} 
  />
);
 
export default function Home() {
  const [data, setData] = useState<IdData>({
    headerTitle: 'IDENTITY CARD',
    headerSubtitle: 'GEKLA - 2024 - 013 THALASSERY LAC',
    headerBanner: 'OFFICER ON ELECTION DUTY',
    photoUrl: null,
    signatureUrl: null,
    topSpaceHeight: 80,
    cardWidth: 650,
    cardHeight: 550,
    pdfCardWidthMm: 103,
    pdfCardHeightMm: 98,
    pdfMarginX: 1,
    pdfMarginY: 10,
    pdfGridCols: 2,
    pdfGridRows: 2,
    headerColor: '#1e3a8a',
    bodyColor: '#dbeafe',
    headerTextColor: '#ffffff',
    bannerTextColor: '#1e3a8a',
    bodyTextColor: '#1e3a8a',
    dynamicFields: [
      { id: '1', label: 'Name', value: 'JOHN DOE' },
      { id: '2', label: 'Designation', value: 'ASSISTANT PROFESSOR' },
      { id: '3', label: 'Office', value: 'GOVT. COLLEGE, KANNUR' },
      { id: '4', label: 'PEN', value: '12345678' }
    ],
  });

  const [bulkData, setBulkData] = useState<any[]>([]);
  const [selectedIndices, setSelectedIndices] = useState<number[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const excelInputRef = useRef<HTMLInputElement>(null);

  const photoInputRef = useRef<HTMLInputElement>(null);
  const signatureInputRef = useRef<HTMLInputElement>(null);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value, type } = e.target;
    setData(prev => ({ 
      ...prev, 
      [name]: type === 'number' ? parseInt(value) || 0 : value 
    }));
  };

  const handleFieldChange = (id: string, value: string) => {
    setData(prev => ({
      ...prev,
      dynamicFields: prev.dynamicFields.map(f => f.id === id ? { ...f, value } : f)
    }));
  };

  const handleLabelChange = (id: string, label: string) => {
    setData(prev => ({
      ...prev,
      dynamicFields: prev.dynamicFields.map(f => f.id === id ? { ...f, label } : f)
    }));
  };

  const addField = () => {
    setData(prev => ({
      ...prev,
      dynamicFields: [...prev.dynamicFields, { id: Date.now().toString(), label: '', value: '' }]
    }));
  };

  const removeField = (id: string) => {
    setData(prev => ({
      ...prev,
      dynamicFields: prev.dynamicFields.filter(f => f.id !== id)
    }));
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'photo' | 'signature') => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        setData(prev => ({
          ...prev,
          [type === 'photo' ? 'photoUrl' : 'signatureUrl']: event.target?.result as string
        }));
      };
      reader.readAsDataURL(file);
    }
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const importedData = XLSX.utils.sheet_to_json(ws);
        setBulkData(importedData);
        setSelectedIndices(importedData.map((_, i) => i)); // Select all by default
        alert(`Successfully imported ${importedData.length} records!`);
      };
      reader.readAsBinaryString(file);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const handleDownloadBulkPDF = async () => {
    const cardElement = document.getElementById('printable-card');
    if (!cardElement) return;

    const getGoogleDriveDirectLink = (url: string) => {
      if (!url || typeof url !== 'string') return url;
      if (url.includes('drive.google.com') || url.includes('docs.google.com')) {
        // Look for ID in ?id=... or /d/ID pattern
        // Regex extracts common Drive IDs (25-50 characters)
        const idMatch = url.match(/[-\w]{25,50}/);
        if (idMatch) {
          // This endpoint is the fastest and has the best CORS support for canvas capture
          return `https://lh3.googleusercontent.com/d/${idMatch[0]}`;
        }
      }
      return url;
    };

    // Filter by selected indices if bulkData exists, otherwise use demo
    let dataSource: any[] = [];
    if (bulkData.length > 0) {
      dataSource = selectedIndices.length > 0 
        ? selectedIndices.map(idx => bulkData[idx])
        : bulkData;
    } else {
      dataSource = [
        { 'NAME (in BLOCK LETTERS)': 'JOHN DOE', 'DESIGNATION': 'ASSISTANT PROFESSOR', 'NAME OF OFFICE': 'OFFICE 1', 'PEN': '11111111' },
      { 'NAME (in BLOCK LETTERS)': 'SARAH SMITH', 'DESIGNATION': 'LECTURER', 'NAME OF OFFICE': 'OFFICE 2', 'PEN': '22222222' },
      { 'NAME (in BLOCK LETTERS)': 'MICHAEL BROWN', 'DESIGNATION': 'PRINCIPAL', 'NAME OF OFFICE': 'OFFICE 3', 'PEN': '33333333' },
      { 'NAME (in BLOCK LETTERS)': 'EMILY DAVIS', 'DESIGNATION': 'COORDINATOR', 'NAME OF OFFICE': 'OFFICE 4', 'PEN': '44444444' },
      { 'NAME (in BLOCK LETTERS)': 'DAVID WILSON', 'DESIGNATION': 'MANAGER', 'NAME OF OFFICE': 'OFFICE 5', 'PEN': '55555555' },
      { 'NAME (in BLOCK LETTERS)': 'ANNA WHITE', 'DESIGNATION': 'SECRETARY', 'NAME OF OFFICE': 'OFFICE 6', 'PEN': '66666666' }
    ];
    }

    const pdf = new jsPDF('p', 'mm', 'a4');
    setIsGenerating(true);
    setProgress({ current: 0, total: dataSource.length });
    
    const cardWidthMm = data.pdfCardWidthMm; 
    const cardHeightMm = data.pdfCardHeightMm; 
    const marginX = data.pdfMarginX;
    const marginY = data.pdfMarginY;

    const originalFields = [...data.dynamicFields];
    
    for (let i = 0; i < dataSource.length; i++) {
      const user = dataSource[i];
      
      const findValue = (keywords: string[]) => {
        const key = Object.keys(user).find(k => 
          keywords.some(kw => k.toLowerCase().includes(kw.toLowerCase()))
        );
        return key ? user[key] : null;
      };

      const photoLink = findValue(['Upload Photo', 'UPLOAD PHOTOGRAPH']) || null;
      const directPhotoLink = photoLink ? getGoogleDriveDirectLink(photoLink) : data.photoUrl;

      const newData = {
        ...data,
        photoUrl: directPhotoLink,
        dynamicFields: data.dynamicFields.map(f => {
          const lowerLabel = f.label.toLowerCase();
          if (lowerLabel === 'name') return { ...f, value: findValue(['NAME (in BLOCK LETTERS)', 'name']) || '' };
          if (lowerLabel === 'designation') return { ...f, value: findValue(['DESIGNATION']) || '' };
          if (lowerLabel === 'office') return { ...f, value: findValue(['NAME OF OFFICE', 'office']) || '' };
          if (lowerLabel === 'pen') return { ...f, value: findValue(['PEN']) || '' };
          return f;
        })
      };
      
      // Using a small trick to wait for state/render and image load
      setData(newData);
      // Wait significantly longer for Google Drive external images to resolve and load
      await new Promise(r => setTimeout(r, 1500)); 
      
      const canvas = await html2canvas(cardElement, { 
        scale: 3,
        useCORS: true, 
        allowTaint: true 
      });
      const imgData = canvas.toDataURL('image/png');
      
      // Dynamic grid placement
      const col = i % data.pdfGridCols;
      const row = Math.floor(i / data.pdfGridCols) % data.pdfGridRows;
      const cardsPerPage = data.pdfGridCols * data.pdfGridRows;

      const x = col * (cardWidthMm + marginX) + marginX;
      const y = row * (cardHeightMm + marginY) + marginY;
      
      // Add new page if we exceed cardsPerPage
      if (i > 0 && i % cardsPerPage === 0) {
        pdf.addPage();
      }

      pdf.addImage(imgData, 'PNG', x, y, cardWidthMm, cardHeightMm);
      setProgress(prev => ({ ...prev, current: i + 1 }));
    }
    
    // Restore original data
    setData(prev => ({ ...prev, dynamicFields: originalFields }));
    
    pdf.save('bulk-id-cards.pdf');
    setIsGenerating(false);
  };

  return (
    <main className="main-container">
      <header className="mb-8 text-center">
        <h1 className="text-4xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-blue-900 to-blue-600 leading-tight">
          Premium ID Generator
        </h1>
        <p className="text-slate-500 mt-2">Create highly dynamic identity cards in seconds.</p>
      </header>

      <div className="id-card-section">
        {/* Card Preview */}
        <section className="id-card-preview">
          <div 
            className="id-card" 
            id="printable-card" 
            style={{ 
              width: `${data.cardWidth}px`, 
              height: `${data.cardHeight}px`,
              backgroundColor: data.bodyColor
            }}
          >
            <div style={{ height: `${data.topSpaceHeight}px`, backgroundColor: data.bodyColor }} className="id-card-top-space">
              <div className="id-card-punch-hole"></div>
            </div>

            <div className="id-card-header" style={{ backgroundColor: data.headerColor }}>
              <div className="id-card-logo-box">
                <ECILogo />
              </div>
              <div className="id-card-header-text">
                <div className="id-card-header-title" style={{ color: data.headerTextColor }}>{data.headerTitle}</div>
                <div className="id-card-header-subtitle" style={{ color: data.headerTextColor }}>{data.headerSubtitle}</div>
              </div>
            </div>

            <div className="w-full bg-white flex items-center justify-center py-4">
              <div className="id-card-header-banner" style={{ color: data.bannerTextColor }}>{data.headerBanner}</div>
            </div>

            <div className="id-card-body" style={{ color: data.bodyTextColor }}>
              <div className="id-card-photo-box">
                {data.photoUrl ? (
                  <img 
                    src={data.photoUrl} 
                    alt="ID Photo" 
                    crossOrigin="anonymous"
                    referrerPolicy="no-referrer"
                    key={`photo-${data.photoUrl}`}
                  />
                ) : (
                  <div className="id-card-photo-placeholder">
                    PHOTO<br />PLACEHOLDER
                  </div>
                )}
              </div>

              <div className="id-card-fields">
                {data.dynamicFields.map(field => (
                  <div key={field.id} className="id-card-field">
                    
                    <span className="id-card-field-label">{field.label}</span>
                    <span className="id-card-field-colon">:</span>
                    <span className="id-card-field-value">{field.value}</span>
                  </div>
                ))}
              </div>

              <div className="id-card-footer">
                <span className="id-card-signature-label">(Signature)</span>
                <div className="id-card-footer-middle"></div>
                <span className="id-card-signature-label">(Signature of ARO)</span>
                <div className="id-card-signature-line">
                  {data.signatureUrl && (
                    <img
                      src={data.signatureUrl}
                      alt="Signature"
                      style={{ height: '35px', width: '120px', objectFit: 'contain', position: 'absolute', bottom: '10px', right: '0' }}
                    />
                  )}
                </div>
              </div>
            </div>
          </div>
          <p style={{ textAlign: 'center', marginTop: '1rem', fontSize: '0.875rem', color: 'var(--secondary)' }}>Live Preview (Scale: 1:1)</p>
        </section>

        {/* Editor Pane */}
        <section className="editor-pane">
          <h2 style={{ marginBottom: '1.5rem', fontSize: '1.25rem' }}>Card Customization</h2>

          <div className="form-group">
            <label className="form-label">Layout</label>
            <div className="grid grid-cols-2 gap-4">
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Width (px)</span>
                <input 
                  type="number" 
                  name="cardWidth" 
                  className="form-input" 
                  value={data.cardWidth || 650} 
                  onChange={handleInputChange} 
                  min="400"
                  max="1000"
                />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Height (px)</span>
                <input 
                  type="number" 
                  name="cardHeight" 
                  className="form-input" 
                  value={data.cardHeight || 550} 
                  onChange={handleInputChange} 
                  min="400"
                  max="1000"
                />
              </div>
              <div className="form-group col-span-2" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Top Space (px)</span>
                <input 
                  type="number" 
                  name="topSpaceHeight" 
                  className="form-input" 
                  value={data.topSpaceHeight || 80} 
                  onChange={handleInputChange} 
                  min="0"
                  max="300"
                />
              </div>
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Header Details</label>
            <input
              type="text"
              name="headerTitle"
              className="form-input"
              value={data.headerTitle || ''}
              onChange={handleInputChange}
              placeholder="Card Title"
            />
          </div>

          <div className="form-row">
            <div className="form-group">
              <input
                type="text"
                name="headerSubtitle"
                className="form-input"
                value={data.headerSubtitle || ''}
                onChange={handleInputChange}
                placeholder="Upper Subtitle"
                style={{ fontSize: '0.875rem' }}
              />
            </div>
          </div>

          <div className="form-group">
            <input
              type="text"
              name="headerBanner"
              className="form-input"
              value={data.headerBanner || ''}
              onChange={handleInputChange}
              placeholder="Bottom Banner"
              style={{ fontWeight: 700 }}
            />
          </div>

          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Photo</label>
              <label className="file-input-label">
                {data.photoUrl ? 'Change Photo' : 'Upload Photo'}
                <input
                  type="file"
                  ref={photoInputRef}
                  onChange={(e) => handleFileUpload(e, 'photo')}
                  accept="image/*"
                  hidden
                />
              </label>
            </div>
            <div className="form-group">
              <label className="form-label">Signature</label>
              <label className="file-input-label">
                {data.signatureUrl ? 'Change Sign' : 'Upload Sign'}
                <input
                  type="file"
                  ref={signatureInputRef}
                  onChange={(e) => handleFileUpload(e, 'signature')}
                  accept="image/*"
                  hidden
                />
              </label>
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Colors</label>
            <div className="grid grid-cols-2 gap-4">
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold">Header BG</span>
                <input type="color" name="headerColor" className="w-full h-8 cursor-pointer rounded overflow-hidden" value={data.headerColor} onChange={handleInputChange} />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold">Base BG</span>
                <input type="color" name="bodyColor" className="w-full h-8 cursor-pointer rounded overflow-hidden" value={data.bodyColor} onChange={handleInputChange} />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold">Header Text</span>
                <input type="color" name="headerTextColor" className="w-full h-8 cursor-pointer rounded overflow-hidden" value={data.headerTextColor} onChange={handleInputChange} />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold">Banner Text</span>
                <input type="color" name="bannerTextColor" className="w-full h-8 cursor-pointer rounded overflow-hidden" value={data.bannerTextColor} onChange={handleInputChange} />
              </div>
              {/* <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-[10px] text-slate-400 uppercase tracking-wider font-semibold">Body Text</span>
                <input type="color" name="bodyTextColor" className="w-full h-8 cursor-pointer rounded overflow-hidden" value={data.bodyTextColor} onChange={handleInputChange} />
              </div> */}
            </div>
          </div>

          <div className="form-group">
            <label className="form-label">Export Settings (PDF in mm)</label>
            <div className="grid grid-cols-2 gap-4">
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Card Width (mm)</span>
                <input 
                  type="number" 
                  name="pdfCardWidthMm" 
                  className="form-input" 
                  value={data.pdfCardWidthMm || 105} 
                  onChange={handleInputChange} 
                />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Card Height (mm)</span>
                <input 
                  type="number" 
                  name="pdfCardHeightMm" 
                  className="form-input" 
                  value={data.pdfCardHeightMm || 99} 
                  onChange={handleInputChange} 
                />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Margin X (mm)</span>
                <input 
                  type="number" 
                  name="pdfMarginX" 
                  className="form-input" 
                  value={data.pdfMarginX || 5} 
                  onChange={handleInputChange} 
                />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">Margin Y (mm)</span>
                <input 
                  type="number" 
                  name="pdfMarginY" 
                  className="form-input" 
                  value={data.pdfMarginY || 10} 
                  onChange={handleInputChange} 
                />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">GridCols</span>
                <input 
                  type="number" 
                  name="pdfGridCols" 
                  className="form-input" 
                  value={data.pdfGridCols || 2} 
                  onChange={handleInputChange} 
                  min="1"
                />
              </div>
              <div className="form-group" style={{ marginBottom: 0 }}>
                <span className="text-xs text-slate-400">GridRows</span>
                <input 
                  type="number" 
                  name="pdfGridRows" 
                  className="form-input" 
                  value={data.pdfGridRows || 2} 
                  onChange={handleInputChange} 
                  min="1"
                />
              </div>
            </div>
          </div>
          
          <div className="dynamic-fields-header">
            <h3 style={{ fontSize: '1rem' }}>Data Fields</h3>
            <button className="btn btn-primary" onClick={addField} style={{ padding: '0.5rem 1rem', fontSize: '0.75rem' }}>
              + Add Field
            </button>
          </div>

          {data.dynamicFields.map((field) => (
            <div key={field.id} className="field-editor-row">
              <input
                type="text"
                className="form-input"
                value={field.label || ''}
                onChange={(e) => handleLabelChange(field.id, e.target.value)}
                placeholder="Label"
                style={{ width: '40%', fontWeight: 600 }}
              />
              <input
                type="text"
                className="form-input"
                value={field.value || ''}
                onChange={(e) => handleFieldChange(field.id, e.target.value)}
                placeholder="Value"
              />
              <button className="btn-danger" onClick={() => removeField(field.id)}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M3 6h18m-2 0v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6m3 0V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"></path>
                </svg>
              </button>
            </div>
          ))}

          <div className="action-bar">
            <button className="btn btn-secondary" style={{ flex: 1 }} onClick={() => excelInputRef.current?.click()}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
              Import Excel
              <input type="file" ref={excelInputRef} onChange={handleExcelImport} accept=".xlsx,.xls,.csv" hidden />
            </button>
            <button className="btn btn-secondary" style={{ flex: 1 }} onClick={handleDownloadBulkPDF}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
              </svg>
              Bulk PDF ({selectedIndices.length || bulkData.length || 6})
            </button>
          </div>

          {bulkData.length > 0 && (
            <div style={{ marginTop: '50px', borderTop: '1px solid #e2e8f0', paddingTop: '10px', paddingBottom: '10px' }}>
              <div className="flex justify-between items-center" style={{ marginBottom: '40px' }}>
                <h3 style={{ fontSize: '1.2rem', fontWeight: 600 }}>Imported Batch Data ({bulkData.length})</h3>
                <div className="flex gap-3">
                  <button 
                    style={{ fontSize: '0.7rem', padding: '4px 10px', border: '1px solid #3b82f6', borderRadius: '4px', color: '#3b82f6', background: 'transparent' }}
                    onClick={() => setSelectedIndices(bulkData.map((_, i) => i))}
                  >
                    Select All
                  </button>
                  <button 
                    style={{ fontSize: '0.7rem', padding: '4px 10px', border: '1px solid #64748b', borderRadius: '4px', color: '#64748b', background: 'transparent' }}
                    onClick={() => setSelectedIndices([])}
                  >
                    Clear All
                  </button>
                  <button 
                    style={{ fontSize: '0.7rem', padding: '4px 10px', border: '1px solid #ef4444', borderRadius: '4px', color: '#ef4444', background: 'transparent' }}
                    onClick={() => { setBulkData([]); setSelectedIndices([]); }}
                  >
                    Remove Data
                  </button>
                </div>
              </div>
              
              <div style={{ maxHeight: '300px', overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: '8px' }}>
                <table style={{ width: '100%', fontSize: '0.75rem', textAlign: 'left', borderCollapse: 'collapse' }}>
                  <thead style={{ background: '#f8fafc', position: 'sticky', top: 0 }}>
                    <tr>
                      <th style={{ padding: '10px 8px', borderBottom: '1px solid #e2e8f0', width: '30px' }}></th>
                      <th style={{ padding: '10px 8px', borderBottom: '1px solid #e2e8f0' }}>NAME</th>
                      <th style={{ padding: '10px 8px', borderBottom: '1px solid #e2e8f0' }}>PEN</th>
                      <th style={{ padding: '10px 8px', borderBottom: '1px solid #e2e8f0' }}>DESIGNATION</th>
                      <th style={{ padding: '10px 8px', borderBottom: '1px solid #e2e8f0' }}>OFFICE</th>
                      <th style={{ padding: '10px 8px', borderBottom: '1px solid #e2e8f0' }}>PHOTO LINK</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bulkData.map((row, idx) => {
                      const photoKey = Object.keys(row).find(k => k.toLowerCase().includes('photo') || k.toLowerCase().includes('photograph')) || 'Upload Photo';
                      const nameKey = Object.keys(row).find(k => k.toLowerCase().includes('name')) || 'NAME (in BLOCK LETTERS)';
                      const penKey = Object.keys(row).find(k => k.toLowerCase().includes('pen')) || 'PEN';
                      const officeKey = Object.keys(row).find(k => k.toLowerCase().includes('office')) || 'NAME OF OFFICE';
                      const designationKey = Object.keys(row).find(k => k.toLowerCase().includes('designation')) || 'DESIGNATION';

                      const updateField = (key: string, value: string) => {
                        const newData = [...bulkData];
                        newData[idx] = { ...row, [key]: value };
                        setBulkData(newData);
                      };

                      return (
                      <tr 
                        key={idx} 
                        style={{ borderBottom: '1px solid #f1f5f9', background: selectedIndices.includes(idx) ? '#f0f9ff' : 'transparent' }}
                      >
                        <td style={{ padding: '8px' }}>
                          <input 
                            type="checkbox" 
                            checked={selectedIndices.includes(idx)}
                            onChange={() => {
                              setSelectedIndices(prev => 
                                prev.includes(idx) ? prev.filter(i => i !== idx) : [...prev, idx]
                              );
                            }}
                          />
                        </td>
                        <td style={{ padding: '8px' }}>
                          <input 
                            type="text" 
                            className="text-[0.65rem] w-full p-1 border rounded"
                            value={String(row[nameKey] || '')}
                            onChange={(e) => updateField(nameKey, e.target.value)}
                          />
                        </td>
                        <td style={{ padding: '8px' }}>
                          <input 
                            type="text" 
                            className="text-[0.65rem] w-full p-1 border rounded"
                            value={String(row[penKey] || '')}
                            onChange={(e) => updateField(penKey, e.target.value)}
                          />
                        </td>
                        <td style={{ padding: '8px' }}>
                          <input 
                            type="text" 
                            className="text-[0.65rem] w-full p-1 border rounded"
                            value={String(row[designationKey] || '')}
                            onChange={(e) => updateField(designationKey, e.target.value)}
                          />
                        </td>
                        <td style={{ padding: '8px' }}>
                          <input 
                            type="text" 
                            className="text-[0.65rem] w-full p-1 border rounded"
                            value={String(row[officeKey] || '')}
                            onChange={(e) => updateField(officeKey, e.target.value)}
                          />
                        </td>
                        <td style={{ padding: '8px' }}>
                          <input 
                            type="text" 
                            className="text-[0.65rem] w-full p-1 border rounded"
                            value={String(row[photoKey] || '')}
                            onChange={(e) => updateField(photoKey, e.target.value)}
                          />
                        </td>
                      </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </section>
      </div>

      {isGenerating && (
        <div style={{ 
          position: 'fixed', bottom: '2rem', right: '2rem', background: '#fff', 
          padding: '1.25rem', borderRadius: '12px', boxShadow: '0 10px 25px -5px rgba(0,0,0,0.15)',
          border: '1px solid #e2e8f0', zIndex: 9999, display: 'flex', alignItems: 'center', gap: '1rem',
          minWidth: '240px'
        }}>
          <div style={{ 
            width: '24px', height: '24px', borderRadius: '50%', border: '3px solid #f1f5f9', 
            borderTopColor: '#3b82f6', animation: 'spin 1s linear infinite' 
          }} />
          <div>
            <div style={{ fontWeight: 600, fontSize: '0.875rem' }}>Generating PDF...</div>
            <div style={{ color: '#64748b', fontSize: '0.75rem' }}>Processing {progress.current} of {progress.total}</div>
          </div>
          <style dangerouslySetInnerHTML={{ __html: `
            @keyframes spin { to { transform: rotate(360deg); } }
          `}} />
        </div>
      )}
    </main>
  );
}
