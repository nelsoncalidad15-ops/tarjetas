/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { 
  Plus, 
  Trash2, 
  Copy, 
  Printer, 
  Sparkles, 
  Sliders, 
  Layers, 
  RotateCw, 
  Layout, 
  Type as FontIcon, 
  Info, 
  Maximize2, 
  Grid, 
  ArrowLeftRight, 
  X, 
  Settings, 
  FolderOpen
} from 'lucide-react';
import { Card, PrintConfig, CardStyle } from './types';
import { CardItem } from './components/CardItem';
import { PrintSheet } from './components/PrintSheet';
import * as XLSX from 'xlsx';

// Initial ISO 14001:2015 demo samples to immediately show the power of the tool
const INITIAL_CARDS: Card[] = [
  {
    id: 'sample-1',
    frontText: '¿Qué debe determinar la organización según el apartado 4.1 de la norma ISO 14001:2015?',
    backText: 'Debe determinar las cuestiones externas e internas que son pertinentes para su propósito y que afectan a su capacidad para lograr los resultados previstos de su sistema de gestión ambiental.',
    useImageFront: false,
    useImageBack: false,
    copies: 2,
    category: 'ISO 14001:2015',
    createdAt: Date.now() - 3000
  },
  {
    id: 'sample-2',
    frontText: '¿Cuál es el objeto primario del Sistema de Gestión Ambiental (SGA)?',
    backText: 'Proporcionar soporte a la organización para proteger el medio ambiente, prevenir la contaminación y responder a las condiciones ambientales cambiantes, en equilibrio con las necesidades socioeconómicas.',
    useImageFront: false,
    useImageBack: false,
    copies: 2,
    category: 'ISO 14001:2015',
    createdAt: Date.now() - 2000
  },
  {
    id: 'sample-3',
    frontText: 'Contingencias o desviaciones: ¿Cómo define la norma ISO 14001 el término "Aspecto Ambiental"?',
    backText: 'Elemento de las actividades, productos o servicios de una organización que interactúa o puede interactuar con el medio ambiente de forma directa o indirecta.',
    useImageFront: false,
    useImageBack: false,
    copies: 1,
    category: 'ISO 14001:2015',
    createdAt: Date.now() - 1000
  }
];

const DEFAULT_STYLE: CardStyle = {
  textSize: 'sm',
  fontFamily: 'sans',
  borderRadiusMm: 3.5,
  textColor: '#1e293b',
  backgroundColor: '#ffffff',
  backTextColor: '#0f172a',
  backBackgroundColor: '#f8fafc',
  alignment: 'center',
  borderStyle: 'solid',
  borderColor: '#cbd5e1'
};

const DEFAULT_CONFIG: PrintConfig = {
  cardWidthMm: 63, // Standard poker width (approx. 2.5 inches)
  cardHeightMm: 88, // Standard poker height (approx. 3.5 inches)
  paperSize: 'a4',
  showCutMarks: true,
  showBorder: true,
  showCardNumbers: true,
  style: DEFAULT_STYLE
};

const normalizePrintConfig = (savedConfig: Partial<PrintConfig> | null | undefined): PrintConfig => ({
  ...DEFAULT_CONFIG,
  ...savedConfig,
  style: {
    ...DEFAULT_STYLE,
    ...(savedConfig?.style || {}),
  },
});

export default function App() {
  // Application state
  const [cards, setCards] = useState<Card[]>(() => {
    // Attempt local storage load
    const saved = localStorage.getItem('flashcards_data');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (parsed.length > 0) return parsed;
      } catch (e) {
        console.error('Error parsing local storage cards', e);
      }
    }
    return INITIAL_CARDS;
  });

  const [config, setConfig] = useState<PrintConfig>(() => {
    const saved = localStorage.getItem('print_config');
    if (saved) {
      try {
        return normalizePrintConfig(JSON.parse(saved));
      } catch (e) {}
    }
    return DEFAULT_CONFIG;
  });

  const [activeTab, setActiveTab] = useState<'editor' | 'instructions'>('editor');
  const [previewSide, setPreviewSide] = useState<'front' | 'back'>('front');
  const [currentPageIndex, setCurrentPageIndex] = useState<number>(0);
  const [searchQuery, setSearchQuery] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [copiedNotification, setCopiedNotification] = useState(false);
  const [sessionToast, setSessionToast] = useState<{ type: 'success' | 'info' | 'error'; message: string } | null>(null);
  const [sidebarViewMode, setSidebarViewMode] = useState<'visual' | 'table'>('table');
  const [previewAllPages, setPreviewAllPages] = useState<boolean>(false);

  const showToast = (message: string, type: 'success' | 'info' | 'error' = 'success') => {
    setSessionToast({ type, message });
    setTimeout(() => {
      setSessionToast(null);
    }, 4500);
  };

  // Bulk text importer states
  const [isBulkCreatorOpen, setIsBulkCreatorOpen] = useState(false);
  const [bulkCreatorMode, setBulkCreatorMode] = useState<'columns' | 'separator' | 'csv'>('columns');
  const [bulkFrontsText, setBulkFrontsText] = useState('');
  const [bulkBacksText, setBulkBacksText] = useState('');
  const [bulkSeparatorText, setBulkSeparatorText] = useState('');
  const [bulkSeparatorChar, setBulkSeparatorChar] = useState('//');
  const [bulkCsvText, setBulkCsvText] = useState('');
  const [bulkCategory, setBulkCategory] = useState('Generales');

  // Excel (.xlsx/.xls) tab storage states
  const [excelSheetNames, setExcelSheetNames] = useState<string[]>([]);
  const [excelActiveSheetName, setExcelActiveSheetName] = useState<string>('');
  const [uploadedWorkbook, setUploadedWorkbook] = useState<any | null>(null);

  // Printer guide interactive states
  const [printerType, setPrinterType] = useState<'drawer' | 'top'>('drawer');
  const [flipDirection, setFlipDirection] = useState<'long' | 'short'>('long');
  const [testPrintSuccess, setTestPrintSuccess] = useState<boolean | null>(null);

  const csvFileInputRef = useRef<HTMLInputElement>(null);

  // Helper helper to clean CSV quotes and double-quotes
  const cleanCSVQuotes = (str: string): string => {
    let cleaned = str.trim();
    if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
      cleaned = cleaned.slice(1, -1);
    }
    cleaned = cleaned.replace(/""/g, '"');
    return cleaned;
  };

  // Keep quoted delimiter intact
  const splitCSVLine = (line: string, delimiter: string): string[] => {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === delimiter && !inQuotes) {
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }
    result.push(current);
    return result;
  };

  // Parse CSV or Tab Separated Values (TSV) from Excel/Google Sheets copy paste
  const parseCsvOrTabSeparated = (text: string) => {
    const lines = text.split(/\r?\n/);
    const cardsToCreate: { front: string; back: string }[] = [];
    
    lines.forEach((line, index) => {
      const trimmed = line.trim();
      if (!trimmed) return;
      
      // Auto-detect delimiter: tab (\t), semicolon (;), or comma (,)
      let delimiter = ',';
      if (trimmed.includes('\t')) {
        delimiter = '\t';
      } else if (trimmed.includes(';')) {
        delimiter = ';';
      }
      
      let parts: string[] = [];
      if (delimiter === '\t') {
        parts = trimmed.split('\t');
      } else {
        parts = splitCSVLine(trimmed, delimiter);
      }
      
      // Skip header lines mimicking anverso/reverso
      if (index === 0 && parts.length >= 2) {
        const col1 = parts[0].toLowerCase();
        const col2 = parts[1].toLowerCase();
        if (
          col1.includes('anverso') || col1.includes('adverso') || col1.includes('front') || col1.includes('pregunta') ||
          col2.includes('reverso') || col2.includes('back') || col2.includes('respuesta') || col2.includes('atras')
        ) {
          return; // Skip headers
        }
      }
      
      if (parts.length >= 1) {
        let frontText = cleanCSVQuotes(parts[0] || '');
        let backText = cleanCSVQuotes(parts[1] || '');
        
        if (frontText || backText) {
          cardsToCreate.push({ front: frontText, back: backText });
        }
      }
    });
    
    return cardsToCreate;
  };

  // Parse typed copy-pasted texts for bulk creation
  const getBulkCardsPreview = () => {
    const cardsToCreate: { front: string; back: string }[] = [];
    if (bulkCreatorMode === 'columns') {
      const fronts = bulkFrontsText.split('\n');
      const backs = bulkBacksText.split('\n');
      const maxLines = Math.max(fronts.length, backs.length);
      
      for (let i = 0; i < maxLines; i++) {
        const frontText = (fronts[i] || '').trim();
        const backText = (backs[i] || '').trim();
        if (frontText || backText) {
          cardsToCreate.push({ front: frontText, back: backText });
        }
      }
    } else if (bulkCreatorMode === 'separator') {
      const lines = bulkSeparatorText.split('\n');
      lines.forEach(line => {
        const trimmed = line.trim();
        if (!trimmed) return;
        
        let frontText = '';
        let backText = '';
        
        if (trimmed.includes(bulkSeparatorChar)) {
          const parts = trimmed.split(bulkSeparatorChar);
          frontText = (parts[0] || '').trim();
          backText = (parts.slice(1).join(bulkSeparatorChar) || '').trim();
        } else {
          frontText = trimmed;
        }
        
        if (frontText || backText) {
          cardsToCreate.push({ front: frontText, back: backText });
        }
      });
    } else if (bulkCreatorMode === 'csv') {
      return parseCsvOrTabSeparated(bulkCsvText);
    }
    return cardsToCreate;
  };

  const handleCreateBulkCards = () => {
    const preview = getBulkCardsPreview();
    if (preview.length === 0) {
      alert('Por favor, ingrese o suba un archivo de texto con contenido estructurado.');
      return;
    }
    
    const targetCat = bulkCategory.trim() || 'Generales';
    
    const newCards: Card[] = preview.map((item, idx) => ({
      id: `bulk-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 5)}`,
      frontText: item.front || 'Pregunta...',
      backText: item.back || 'Respuesta...',
      useImageFront: false,
      useImageBack: false,
      copies: 1,
      category: targetCat,
      createdAt: Date.now()
    }));
    
    setCards(prev => [...newCards, ...prev]);
    
    // Automatically reset search & category filters so the new cards are GUARANTEED to be visible immediately!
    setSearchQuery('');
    setCategoryFilter('all');
    
    // Clear inputs and close the bulk accordion
    setBulkFrontsText('');
    setBulkBacksText('');
    setBulkSeparatorText('');
    setBulkCsvText('');
    setIsBulkCreatorOpen(false);
    
    // Dispatch helpful visual toast notification
    showToast(`🎉 ¡Se agregaron ${newCards.length} tarjetas con éxito en la categoría "${targetCat}"! Míralas en las hojas de la derecha.`, 'success');
  };

  // Download Sample CSV template for Excel
  const handleDownloadSampleCsv = () => {
    const csvContent = 
      "Anverso;Reverso\n" +
      "¿Cuáles son los colores primarios?;Rojo, amarillo y azul\n" +
      "La capital de Francia;París\n" +
      "Fórmula química del agua;H2O\n" +
      "¿De qué color es la caja negra de un avión comercial?;Es de naranja brillante para facilitar su búsqueda\n" +
      "Año del descubrimiento de América;1492";

    // BOM character to ensure Excel respects Spanish accents (UTF-8 encoding)
    const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "plantilla_tarjetas.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const importRawRowsToCards = (rawRows: any[], sheetName?: string, replace: boolean = false) => {
    const cardsToCreate: Card[] = [];
    
    rawRows.forEach((row, idx) => {
      if (!row || !Array.isArray(row)) return;
      
      // Header check: skip if row resembles column titles
      if (idx === 0) {
        const col1 = String(row[0] || '').toLowerCase().trim();
        const col2 = String(row[1] || '').toLowerCase().trim();
        if (
          col1.includes('anverso') || col1.includes('adverso') || col1.includes('front') || col1.includes('pregunta') ||
          col2.includes('reverso') || col2.includes('back') || col2.includes('respuesta') || col2.includes('atras')
        ) {
          return; // Skip headers
        }
      }

      const frontVal = String(row[0] === undefined || row[0] === null ? '' : row[0]).trim();
      const backVal = String(row[1] === undefined || row[1] === null ? '' : row[1]).trim();
      
      // Category from column 3 if available, else sheetName or 'Generales'
      const catVal = String(row[2] === undefined || row[2] === null ? '' : row[2]).trim() || sheetName || 'Generales';

      if (frontVal || backVal) {
        cardsToCreate.push({
          id: `excel-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 5)}`,
          frontText: frontVal || 'Pregunta...',
          backText: backVal || 'Respuesta...',
          useImageFront: false,
          useImageBack: false,
          copies: 1,
          category: catVal || 'Generales',
          createdAt: Date.now() + idx
        });
      }
    });

    if (cardsToCreate.length > 0) {
      if (replace) {
        setCards(cardsToCreate);
        showToast(`📊 Se cargaron ${cardsToCreate.length} tarjetas reemplazando las anteriores con éxito.`, 'success');
      } else {
        setCards(prev => [...prev, ...cardsToCreate]);
        showToast(`🎉 ¡Se añadieron ${cardsToCreate.length} tarjetas de tu Excel directamente!`, 'success');
      }
      setSearchQuery('');
      setCategoryFilter('all');
    } else {
      showToast(`⚠️ No se encontraron filas con contenido válido para importar.`, 'error');
    }
  };

  const handleCsvFileLoad = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || e.target.files.length === 0) return;
    const file = e.target.files[0];
    const fileName = file.name.toLowerCase();
    
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          if (event.target && event.target.result) {
            const data = new Uint8Array(event.target.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            
            setUploadedWorkbook(workbook);
            const sheetNames = workbook.SheetNames;
            setExcelSheetNames(sheetNames);
            
            const firstSheetName = sheetNames[0];
            setExcelActiveSheetName(firstSheetName);
            
            const worksheet = workbook.Sheets[firstSheetName];
            const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
            
            // Generate TSV draft just for bulk textarea fallback representation
            const tsvLines = rawRows.map(row => {
              if (!row || !Array.isArray(row)) return '';
              const cell1 = String(row[0] === undefined || row[0] === null ? '' : row[0]).replace(/\t/g, ' ').trim();
              const cell2 = String(row[1] === undefined || row[1] === null ? '' : row[1]).replace(/\t/g, ' ').trim();
              if (!cell1 && !cell2) return '';
              return `${cell1}\t${cell2}`;
            }).filter(line => line).join('\n');
            setBulkCsvText(tsvLines);

            // Directly import cards! Prompt user for append or replace
            const replace = cards.length > 0 && window.confirm(
              `📊 Detectamos ${rawRows.length} filas en la pestaña "${firstSheetName}" de tu archivo Excel. ¿Deseas REEMPLAZAR tus tarjetas actuales o AÑADIR las nuevas al final?\n\n- Click ACEPTAR para Reemplazar las actuales.\n- Click CANCELAR para Añadirlas al final.`
            );
            importRawRowsToCards(rawRows, firstSheetName, replace);
          }
        } catch (err: any) {
          console.error("Error al procesar archivo Excel", err);
          showToast(`❌ Error al procesar archivo Excel: ${err?.message || 'Archivo dañado o inválido'}`, 'error');
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = (event) => {
        if (event.target && event.target.result) {
          const text = event.target.result as string;
          setBulkCsvText(text);
          setExcelSheetNames([]);
          setExcelActiveSheetName('');
          setUploadedWorkbook(null);
          
          const preview = parseCsvOrTabSeparated(text);
          const rawRows = preview.map(p => [p.front, p.back]);
          const replace = cards.length > 0 && window.confirm(
            `📋 Detectamos ${preview.length} tarjetas en tu archivo de texto. ¿Deseas REEMPLAZAR tus tarjetas actuales o AÑADIR las nuevas al final?\n\n- Click ACEPTAR para Reemplazar.\n- Click CANCELAR para Añadir.`
          );
          importRawRowsToCards(rawRows, 'Generales', replace);
        }
      };
      reader.readAsText(file, "UTF-8");
    }
  };

  const handleExcelSheetChange = (sheetName: string) => {
    if (!uploadedWorkbook) return;
    try {
      setExcelActiveSheetName(sheetName);
      const worksheet = uploadedWorkbook.Sheets[sheetName];
      const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
      
      const tsvLines = rawRows.map(row => {
        if (!row || !Array.isArray(row)) return '';
        const cell1 = String(row[0] === undefined || row[0] === null ? '' : row[0]).replace(/\t/g, ' ').trim();
        const cell2 = String(row[1] === undefined || row[1] === null ? '' : row[1]).replace(/\t/g, ' ').trim();
        if (!cell1 && !cell2) return '';
        return `${cell1}\t${cell2}`;
      }).filter(line => line).join('\n');
      setBulkCsvText(tsvLines);

      const replace = cards.length > 0 && window.confirm(
        `📄 Cambiando a la pestaña "${sheetName}". ¿Deseas REEMPLAZAR tus tarjetas de impresión actuales por las de esta hoja, o AÑADIRLAS al final?\n\n- Click ACEPTAR para Reemplazar.\n- Click CANCELAR para Añadir.`
      );
      importRawRowsToCards(rawRows, sheetName, replace);
    } catch (err: any) {
      console.error(err);
      showToast('❌ Error al cambiar de pestaña en el archivo Excel', 'error');
    }
  };

  // Sync state to local storage
  useEffect(() => {
    localStorage.setItem('flashcards_data', JSON.stringify(cards));
  }, [cards]);

  useEffect(() => {
    localStorage.setItem('print_config', JSON.stringify(config));
  }, [config]);

  // Math sizing & capacity variables for printing
  const pageWidth = config.paperSize === 'a4' ? 210 : 215.9;
  const pageHeight = config.paperSize === 'a4' ? 297 : 279.4;
  const marginMm = 8; // standard safe page border margins
  const cols = Math.max(1, Math.floor((pageWidth - marginMm * 2) / config.cardWidthMm));
  const rows = Math.max(1, Math.floor((pageHeight - marginMm * 2) / config.cardHeightMm));
  const cardsPerPage = cols * rows;

  // Flatten normal cards with copy counter multiplier
  const flattenedPrintCards: Card[] = [];
  cards.forEach((card, index) => {
    for (let c = 0; c < card.copies; c++) {
      flattenedPrintCards.push({
        ...card,
        studyNumber: index + 1,
      });
    }
  });

  const totalPagesNeeded = Math.ceil(flattenedPrintCards.length / cardsPerPage);

  // Keep index within boundaries
  useEffect(() => {
    if (currentPageIndex >= totalPagesNeeded && totalPagesNeeded > 0) {
      setCurrentPageIndex(totalPagesNeeded - 1);
    }
  }, [flattenedPrintCards.length, cardsPerPage, searchQuery, categoryFilter]);

  // Handlers
  const handleAddNewCard = () => {
    const newCard: Card = {
      id: `card-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      frontText: 'Escriba la pregunta o término del anverso aquí...',
      backText: 'Escriba la respuesta o explicación del reverso aquí...',
      useImageFront: false,
      useImageBack: false,
      copies: 1,
      category: categoryFilter !== 'all' ? categoryFilter : 'Generales',
      createdAt: Date.now()
    };
    setCards([newCard, ...cards]);
  };

  const handleAddNewPrintSheet = () => {
    const blankCards: Card[] = [];
    const targetCat = categoryFilter !== 'all' ? categoryFilter : 'Generales';
    
    // Create exactly cardsPerPage empty card templates
    for (let i = 0; i < cardsPerPage; i++) {
      blankCards.push({
        id: `card-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
        frontText: `Anverso [Hoja ${totalPagesNeeded + 1} - Tarjeta ${i + 1}]`,
        backText: `Reverso [Hoja ${totalPagesNeeded + 1} - Tarjeta ${i + 1}]`,
        useImageFront: false,
        useImageBack: false,
        copies: 1,
        category: targetCat,
        createdAt: Date.now() + i
      });
    }
    
    setCards([...cards, ...blankCards]);
    showToast(`📄 ¡Se ha añadido una nueva hoja de impresión (con ${cardsPerPage} tarjetas de plantilla listas para editar)! Míralas al final.`, 'success');
    
    // Automatically focus navigation on the final page
    setTimeout(() => {
      setCurrentPageIndex(totalPagesNeeded);
    }, 120);
  };


  const handleUpdateCard = (updatedCard: Card) => {
    setCards(cards.map(c => c.id === updatedCard.id ? updatedCard : c));
  };

  const handleDeleteCard = (id: string) => {
    setCards(cards.filter(c => c.id !== id));
  };

  const handleDuplicateCard = (cardToDuplicate: Card) => {
    const duplicate: Card = {
      ...cardToDuplicate,
      id: `card-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`,
      createdAt: Date.now()
    };
    // Insert right after the duplicated card
    const targetIdx = cards.findIndex(c => c.id === cardToDuplicate.id);
    const updated = [...cards];
    updated.splice(targetIdx + 1, 0, duplicate);
    setCards(updated);
  };

  const handleBulkSetCopies = (count: number) => {
    setCards(cards.map(c => ({ ...c, copies: Math.max(1, count) })));
  };

  const handleBulkSetTextOnly = (frontOrBack: 'both' | 'front' | 'back') => {
    setCards(cards.map(c => {
      const u = { ...c };
      if (frontOrBack === 'both' || frontOrBack === 'front') u.useImageFront = false;
      if (frontOrBack === 'both' || frontOrBack === 'back') u.useImageBack = false;
      return u;
    }));
  };

  const handleResetToDemo = () => {
    if (window.confirm('¿Desea restaurar las tarjetas de ejemplo? Sus tarjetas actuales se reemplazarán.')) {
      setCards(INITIAL_CARDS);
      setConfig(DEFAULT_CONFIG);
    }
  };

  const handleClearAll = () => {
    if (window.confirm('¿Está seguro de de que desea eliminar todas las tarjetas?')) {
      setCards([]);
    }
  };

  // Preset Sizes
  const applyPresetSize = (type: 'poker' | 'bridge' | 'business' | 'small') => {
    if (type === 'poker') {
      setConfig({ ...config, cardWidthMm: 63.5, cardHeightMm: 88.9 });
    } else if (type === 'bridge') {
      setConfig({ ...config, cardWidthMm: 57, cardHeightMm: 88 });
    } else if (type === 'business') {
      setConfig({ ...config, cardWidthMm: 85, cardHeightMm: 55 });
    } else if (type === 'small') {
      setConfig({ ...config, cardWidthMm: 50, cardHeightMm: 75 });
    }
  };

  const applyPresetColor = (preset: 'clean' | 'classic' | 'dark' | 'softBlue' | 'amber') => {
    const currentStyle = { ...config.style };
    if (preset === 'clean') {
      currentStyle.textColor = '#1e293b';
      currentStyle.backgroundColor = '#ffffff';
      currentStyle.backTextColor = '#020617';
      currentStyle.backBackgroundColor = '#fafafa';
      currentStyle.borderColor = '#cbd5e1';
    } else if (preset === 'classic') {
      currentStyle.textColor = '#3b2314';
      currentStyle.backgroundColor = '#fdfbf7';
      currentStyle.backTextColor = '#3b2314';
      currentStyle.backBackgroundColor = '#fbf8f1';
      currentStyle.borderColor = '#e7e0d3';
    } else if (preset === 'dark') {
      currentStyle.textColor = '#f8fafc';
      currentStyle.backgroundColor = '#0f172a';
      currentStyle.backTextColor = '#f8fafc';
      currentStyle.backBackgroundColor = '#1e293b';
      currentStyle.borderColor = '#334155';
    } else if (preset === 'softBlue') {
      currentStyle.textColor = '#1e3a8a';
      currentStyle.backgroundColor = '#eff6ff';
      currentStyle.backTextColor = '#172554';
      currentStyle.backBackgroundColor = '#f0fdf4'; // contrast reverso greenish
      currentStyle.borderColor = '#bfdbfe';
    } else if (preset === 'amber') {
      currentStyle.textColor = '#78350f';
      currentStyle.backgroundColor = '#fffbeb';
      currentStyle.backTextColor = '#78350f';
      currentStyle.backBackgroundColor = '#fef3c7';
      currentStyle.borderColor = '#fde68a';
    }
    setConfig({ ...config, style: currentStyle });
  };


  const handlePrint = () => {
    window.print();
  };

  // Get unique categories list
  const uniqueCategories = ['all', ...Array.from(new Set(cards.map(c => c.category || 'Generales'))).filter(Boolean)];

  // Filter cards to view on left pane
  const filteredCards = cards.filter(card => {
    const query = searchQuery.toLowerCase();
    const textMatches = card.frontText.toLowerCase().includes(query) || 
                        card.backText.toLowerCase().includes(query);
    const categoryMatches = categoryFilter === 'all' || 
                            (card.category || 'Generales') === categoryFilter;
    return textMatches && categoryMatches;
  });

  return (
    <div className="flex h-screen w-full overflow-hidden bg-slate-100 font-sans text-slate-800 antialiased print:block print:bg-white print:h-auto print:overflow-visible">
      
      {/* LEFT SIDEBAR: Config & General Actions */}
      <aside className="no-print w-80 bg-slate-900 text-white flex flex-col border-r border-slate-800 shrink-0">
        
        {/* Sidebar Header in theme style */}
        <div className="p-5 border-b border-slate-800">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-xl font-bold tracking-tight text-blue-400">
                CardPrint<span className="text-white italic">Studio</span>
              </h1>
              <p className="text-[9px] uppercase tracking-widest text-slate-500 mt-0.5">
                Editor e Impresor de Tarjetas
              </p>
            </div>
            <div className="bg-slate-800 py-1 px-2 rounded text-[10px] text-blue-400 font-bold border border-slate-700">
              300 DPI
            </div>
          </div>
        </div>

        {/* Sidebar Body Scrollable Controls */}
        <div className="flex-1 p-5 space-y-6 overflow-y-auto custom-scrollbar">

          {/* Section: Quick Start */}
          <div className="space-y-3 bg-slate-800/40 p-4 rounded-xl border border-slate-800">
            <div className="flex items-center justify-between">
              <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider block">
                Inicio Rapido
              </label>
              <button 
                onClick={handleResetToDemo}
                className="text-[10px] text-blue-400 font-semibold hover:underline"
                title="Restaurar datos de demostracion"
              >
                Cargar Demo
              </button>
            </div>

            <div className="rounded-xl border border-dashed border-slate-700 bg-slate-900/30 p-4 text-xs text-slate-300 space-y-2">
              <p className="font-semibold text-white">La app ahora funciona sin backend ni IA.</p>
              <p>Agrega tarjetas manualmente, pega texto por lotes o importa Excel/CSV para preparar e imprimir tus fichas.</p>
            </div>
          </div>

          {/* Section: Size Calibration */}
          <div className="space-y-4">
            <div className="flex items-center gap-2 pb-1 border-b border-slate-800">
              <Sliders className="w-3.5 h-3.5 text-blue-400" />
              <label className="text-[11px] font-bold text-slate-400 uppercase tracking-widest block">
                Dimensiones de Tarjeta
              </label>
            </div>

            {/* Presets Grid */}
            <div>
              <p className="text-[10px] text-slate-400 mb-2 font-medium">Tamaños estándar comunes:</p>
              <div className="grid grid-cols-2 gap-1.5">
                <button
                  onClick={() => applyPresetSize('poker')}
                  className={`text-[10px] py-1.5 px-2 rounded border transition-colors ${
                    config.cardWidthMm > 62 && config.cardWidthMm < 65
                      ? 'bg-blue-600/20 text-blue-300 border-blue-500/50 font-bold'
                      : 'bg-slate-800/60 text-slate-300 border-slate-800 hover:bg-slate-800'
                  }`}
                >
                  Poker (63 x 88 mm)
                </button>
                <button
                  onClick={() => applyPresetSize('bridge')}
                  className={`text-[10px] py-1.5 px-2 rounded border transition-colors ${
                    config.cardWidthMm === 57
                      ? 'bg-blue-600/20 text-blue-300 border-blue-500/50 font-bold'
                      : 'bg-slate-800/60 text-slate-300 border-slate-800 hover:bg-slate-800'
                  }`}
                >
                  Bridge (57 x 88 mm)
                </button>
                <button
                  onClick={() => applyPresetSize('business')}
                  className={`text-[10px] py-1.5 px-2 rounded border transition-colors ${
                    config.cardWidthMm === 85 && config.cardHeightMm === 55
                      ? 'bg-blue-600/20 text-blue-300 border-blue-500/50 font-bold'
                      : 'bg-slate-800/60 text-slate-300 border-slate-800 hover:bg-slate-800'
                  }`}
                >
                  Ficha (85 x 55 mm)
                </button>
                <button
                  onClick={() => applyPresetSize('small')}
                  className={`text-[10px] py-1.5 px-2 rounded border transition-colors ${
                    config.cardWidthMm === 50
                      ? 'bg-blue-600/20 text-blue-300 border-blue-500/50 font-bold'
                      : 'bg-slate-800/60 text-slate-300 border-slate-800 hover:bg-slate-800'
                  }`}
                >
                  Compacta (50 x 75 mm)
                </button>
              </div>
            </div>

            {/* Custom inputs */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-[10px] text-slate-400 block mb-1">Ancho (mm)</label>
                <input
                  type="number"
                  value={config.cardWidthMm}
                  onChange={(e) => setConfig({ ...config, cardWidthMm: Math.max(10, parseInt(e.target.value) || 10) })}
                  className="w-full bg-slate-800 text-white font-mono text-xs p-2 rounded border border-slate-700 text-center"
                />
              </div>
              <div>
                <label className="text-[10px] text-slate-400 block mb-1">Alto (mm)</label>
                <input
                  type="number"
                  value={config.cardHeightMm}
                  onChange={(e) => setConfig({ ...config, cardHeightMm: Math.max(10, parseInt(e.target.value) || 10) })}
                  className="w-full bg-slate-800 text-white font-mono text-xs p-2 rounded border border-slate-700 text-center"
                />
              </div>
            </div>
          </div>

          {/* Section: Study helpers */}
          <div className="space-y-3 bg-slate-800/40 p-4 rounded-xl border border-slate-800">
            <div className="flex items-center gap-2 pb-1 border-b border-slate-800/70">
              <Settings className="w-3.5 h-3.5 text-blue-400" />
              <label className="text-[11px] font-bold text-slate-400 uppercase tracking-widest block">
                Herramientas de estudio
              </label>
            </div>
            <p className="text-[10px] text-slate-400 leading-normal">
              Estas opciones hacen que las tarjetas sean más útiles para repasar y memorizar.
            </p>
            <label className="flex items-center justify-between gap-3 text-xs text-slate-300 cursor-pointer">
              <span>Numerar tarjetas</span>
              <input
                type="checkbox"
                checked={config.showCardNumbers ?? true}
                onChange={(e) => setConfig({ ...config, showCardNumbers: e.target.checked })}
                className="h-4 w-4 rounded border-slate-600 bg-slate-900 text-blue-500 focus:ring-blue-500"
              />
            </label>
          </div>

          {/* Section: Typography & Styling Options */}
          <div className="space-y-4">
            <div className="flex items-center gap-2 pb-1 border-b border-slate-800">
              <FontIcon className="w-3.5 h-3.5 text-blue-400" />
              <label className="text-[11px] font-bold text-slate-400 uppercase tracking-widest block">
                Diseño Visual Interno
              </label>
            </div>

            {/* Stylings Theme presets */}
            <div>
              <p className="text-[10px] text-slate-400 mb-1.5">Esquemas de Color:</p>
              <div className="flex flex-wrap gap-1.5">
                {[
                  { name: 'Limpio', key: 'clean', bg: 'bg-white text-slate-800 border-slate-300' },
                  { name: 'Sepia', key: 'classic', bg: 'bg-[#fdfbf7] text-[#3b2314] border-[#e7e0d3]' },
                  { name: 'Oscuro', key: 'dark', bg: 'bg-slate-900 text-slate-100 border-slate-800' },
                  { name: 'Celeste', key: 'softBlue', bg: 'bg-blue-100 text-blue-900 border-blue-200' },
                  { name: 'Ámbar', key: 'amber', bg: 'bg-amber-50 text-amber-900 border-amber-200' }
                ].map((item) => (
                  <button
                    key={item.key}
                    onClick={() => applyPresetColor(item.key as any)}
                    className={`text-[9px] py-1 px-2.5 rounded border ${item.bg} hover:brightness-95 transition-all font-medium`}
                  >
                    {item.name}
                  </button>
                ))}
              </div>
            </div>

            {/* Font selector */}
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] text-slate-400 block mb-1">Tipografía</label>
                <select
                  value={config.style.fontFamily}
                  onChange={(e) => setConfig({
                    ...config,
                    style: { ...config.style, fontFamily: e.target.value as any }
                  })}
                  className="w-full bg-slate-800 text-slate-100 text-xs p-2 rounded border border-slate-700"
                >
                  <option value="sans">Moderna (Sans)</option>
                  <option value="mono">Código (Mono)</option>
                  <option value="serif">Elegante (Serif)</option>
                </select>
              </div>

              <div>
                <label className="text-[10px] text-slate-400 block mb-1">Tamaño Texto</label>
                <select
                  value={config.style.textSize}
                  onChange={(e) => setConfig({
                    ...config,
                    style: { ...config.style, textSize: e.target.value as any }
                  })}
                  className="w-full bg-slate-800 text-slate-100 text-xs p-2 rounded border border-slate-700"
                >
                  <option value="xs">Muy Chico (10px)</option>
                  <option value="sm">Chico (12px)</option>
                  <option value="base">Normal (14px)</option>
                  <option value="lg">Mediano (16px)</option>
                  <option value="xl">Grande (18px)</option>
                  <option value="2xl">Muy Grande (20px)</option>
                </select>
              </div>
            </div>

            {/* Border Options */}
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] text-slate-400 block mb-1">Borde Tarjeta</label>
                <select
                  value={config.style.borderStyle}
                  onChange={(e) => setConfig({
                    ...config,
                    style: { ...config.style, borderStyle: e.target.value as any }
                  })}
                  className="w-full bg-slate-800 text-slate-100 text-xs p-2 rounded border border-slate-700"
                >
                  <option value="solid">Línea Sólida</option>
                  <option value="dashed">Segmentado</option>
                  <option value="dotted">Punteado</option>
                  <option value="none">Sin Bordes</option>
                </select>
              </div>

              <div>
                <label className="text-[10px] text-slate-400 block mb-1">Esquinas Rd (mm)</label>
                <input
                  type="number"
                  min="0"
                  max="15"
                  value={config.style.borderRadiusMm}
                  onChange={(e) => setConfig({
                    ...config,
                    style: { ...config.style, borderRadiusMm: Math.max(0, parseFloat(e.target.value) || 0) }
                  })}
                  className="w-full bg-slate-800 text-white font-mono text-xs p-2 rounded border border-slate-700 text-center"
                />
              </div>
            </div>

            {/* Custom color manual pickers */}
            <div className="space-y-2 bg-slate-850 p-2.5 rounded-lg border border-slate-800 text-xs">
              <div className="flex items-center justify-between">
                <span className="text-[10px] text-slate-400 uppercase">Frente Fondo:</span>
                <input
                  type="color"
                  value={config.style.backgroundColor}
                  onChange={(e) => setConfig({
                    ...config,
                    style: { ...config.style, backgroundColor: e.target.value }
                  })}
                  className="w-8 h-4 rounded cursor-pointer"
                />
              </div>
              <div className="flex items-center justify-between">
                <span className="text-[10px] text-slate-400 uppercase">Atrás Fondo:</span>
                <input
                  type="color"
                  value={config.style.backBackgroundColor}
                  onChange={(e) => setConfig({
                    ...config,
                    style: { ...config.style, backBackgroundColor: e.target.value }
                  })}
                  className="w-8 h-4 rounded cursor-pointer"
                />
              </div>
            </div>
          </div>

          {/* Section: Page and layout specs */}
          <div className="space-y-4">
            <div className="flex items-center gap-2 pb-1 border-b border-slate-800">
              <Layout className="w-3.5 h-3.5 text-blue-400" />
              <label className="text-[11px] font-bold text-slate-400 uppercase tracking-widest block">
                Formato del Papel
              </label>
            </div>

            <div className="flex items-center justify-between bg-slate-850 p-2 rounded-lg border border-slate-800">
              <span className="text-xs text-slate-300">Tipo de Papel</span>
              <div className="flex gap-1">
                <button
                  onClick={() => setConfig({ ...config, paperSize: 'a4' })}
                  className={`px-2 py-1 text-[10px] uppercase font-bold rounded transition-colors ${
                    config.paperSize === 'a4'
                      ? 'bg-blue-600 text-white'
                      : 'bg-slate-800 text-slate-400 hover:text-white'
                  }`}
                >
                  A4
                </button>
                <button
                  onClick={() => setConfig({ ...config, paperSize: 'letter' })}
                  className={`px-2 py-1 text-[10px] uppercase font-bold rounded transition-colors ${
                    config.paperSize === 'letter'
                      ? 'bg-blue-600 text-white'
                      : 'bg-slate-800 text-slate-400 hover:text-white'
                  }`}
                >
                  Carta (US)
                </button>
              </div>
            </div>

            {/* Cut marks guides toggling */}
            <div className="flex items-center justify-between text-xs text-slate-300">
              <span>Dibujar marcas de corte</span>
              <input
                type="checkbox"
                checked={config.showCutMarks}
                onChange={(e) => setConfig({ ...config, showCutMarks: e.target.checked })}
                className="w-4 h-4 text-blue-600 border-slate-700 bg-slate-800 rounded focus:ring-0 focus:outline-none"
              />
            </div>

          </div>

          {/* Section: Bulk Action shortcuts */}
          <div className="space-y-3 bg-slate-900/60 p-3.5 border border-slate-800 rounded-xl">
            <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">
              Atajos y Acciones Bulk
            </label>
            <div className="space-y-1.5">
              <div className="grid grid-cols-2 gap-1.5">
                <button 
                  onClick={() => handleBulkSetCopies(1)}
                  className="text-[10px] py-1 bg-slate-800 text-slate-300 hover:text-white rounded border border-slate-700 hover:bg-slate-700"
                >
                  Todo 1 copia
                </button>
                <button 
                  onClick={() => handleBulkSetCopies(2)}
                  className="text-[10px] py-1 bg-slate-800 text-slate-300 hover:text-white rounded border border-slate-700 hover:bg-slate-700"
                >
                  Todo 2 copias
                </button>
              </div>

              <button 
                onClick={() => handleBulkSetTextOnly('both')}
                className="w-full text-left font-mono text-[9px] py-1.5 px-2 bg-slate-800/40 text-slate-400 hover:text-slate-200 border border-slate-800 rounded flex items-center justify-between"
              >
                <span>Usar texto limpio en todas</span>
                <span className="text-blue-400 font-bold">TEXTO CLEAN</span>
              </button>

              <button
                onClick={handleClearAll}
                className="w-full text-left font-mono text-[9px] py-1.5 px-2 bg-red-950/20 text-red-400/80 hover:text-red-300 border border-red-950/20 rounded flex items-center justify-between"
              >
                <span>Eliminar todo el inventario</span>
                <span className="text-red-500 font-bold">BORRAR</span>
              </button>
            </div>
          </div>

          {/* Section: Real-time Layout density optimization metrics */}
          <div className="bg-blue-600/10 border border-blue-500/20 rounded-xl p-4">
            <label className="text-[11px] font-bold text-blue-400 uppercase tracking-widest block mb-1">
              Optimización de Carilla
            </label>
            <p className="text-[10px] text-slate-400 leading-normal mb-3.5">
              Calculamos dinámicamente el layout ideal según las medidas físicas especificadas:
            </p>
            <div className="grid grid-cols-2 gap-2">
              <div className="text-center bg-slate-900 rounded-lg py-2.5 border border-slate-800">
                <div className="text-lg font-bold text-slate-100">{flattenedPrintCards.length}</div>
                <div className="text-[9px] text-slate-500 uppercase tracking-wide">Tarjetas Totales</div>
              </div>
              <div className="text-center bg-slate-900 rounded-lg py-2.5 border border-slate-800">
                <div className="text-lg font-bold text-blue-400">{cardsPerPage}</div>
                <div className="text-[9px] text-slate-500 uppercase tracking-wide">Cupos por Hoja</div>
              </div>
            </div>
            
            <div className="mt-3 py-1.5 border-t border-slate-800 text-[10px] text-slate-400 flex justify-between items-center">
              <span>Hojas de Frente requeridas:</span>
              <span className="font-mono text-slate-200 font-bold">{totalPagesNeeded} {totalPagesNeeded === 1 ? 'hoja' : 'hojas'}</span>
            </div>
            <div className="text-[9px] text-blue-400/80 italic mt-1.5 leading-normal">
              * Distribución exacta en cuadrícula de {cols} columnas x {rows} filas por carilla.
            </div>
            
            {/* Expander button to let the user add empty pages instantly */}
            <button
              onClick={handleAddNewPrintSheet}
              className="mt-3.5 w-full bg-slate-900 border border-blue-500/30 hover:border-blue-500 hover:bg-slate-800 text-white font-bold py-2.5 px-3 rounded-lg text-xs flex items-center justify-center gap-1.5 transition-all cursor-pointer shadow-sm hover:shadow-md"
            >
              <span>📄</span> Añadir Hoja de Impresión (+{cardsPerPage} Tarjetas)
            </button>
          </div>
        </div>

        {/* Action button in matching styling */}
        <div className="p-4 border-t border-slate-800 bg-slate-950/50">
          <button
            onClick={handlePrint}
            className="w-full bg-blue-600 hover:bg-blue-500 text-white font-bold py-3 px-4 rounded-xl flex items-center justify-center gap-2 transition-all cursor-pointer shadow-lg shadow-blue-900/30 text-sm"
          >
            <Printer className="w-4 h-4 shrink-0" />
            Imprimir Frente y Dorso
          </button>
        </div>
      </aside>

      {/* MAIN CONTAINER: Workspace editor and preview panels side-by-side */}
      <main className="flex-1 flex flex-col min-w-0 print:p-0 print:m-0 print:bg-white overflow-hidden">
        
        {/* Workspace Header Navbar: Options toggle */}
        <header className="no-print h-14 bg-white border-b border-gray-200 px-6 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-6">
            <nav className="flex gap-1">
              <button
                onClick={() => setActiveTab('editor')}
                className={`py-4 px-4 text-xs font-bold tracking-wider uppercase border-b-2 transition-all ${
                  activeTab === 'editor'
                    ? 'border-blue-600 text-blue-600 font-bold'
                    : 'border-transparent text-gray-500 hover:text-gray-900'
                }`}
              >
                Diseñador de Tarjetas
              </button>
              <button
                onClick={() => setActiveTab('instructions')}
                className={`py-4 px-4 text-xs font-bold tracking-wider uppercase border-b-2 transition-all ${
                  activeTab === 'instructions'
                    ? 'border-blue-600 text-blue-600 font-bold'
                    : 'border-transparent text-gray-500 hover:text-gray-900'
                }`}
              >
                Guía de Impresión Doble Faz
              </button>
            </nav>
          </div>

          <div className="flex items-center gap-3">
            {/* Quick search input */}
            <div className="relative">
              <input
                type="text"
                placeholder="Buscar en tarjetas..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-48 text-xs bg-gray-50 border border-gray-200 rounded-lg py-1.5 pl-2.5 pr-8 focus:bg-white focus:ring-1 focus:ring-blue-500/20 focus:border-blue-500 outline-none"
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-2 text-gray-400 hover:text-gray-700"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              )}
            </div>

            {/* Category filter pick list */}
            {uniqueCategories.length > 2 && (
              <select
                value={categoryFilter}
                onChange={(e) => setCategoryFilter(e.target.value)}
                className="text-xs bg-gray-50 border border-gray-200 rounded-lg py-1.5 px-2.5 outline-none focus:bg-white"
              >
                <option value="all">Todas las Materias</option>
                {uniqueCategories.filter(cat => cat !== 'all').map((cat) => (
                  <option key={cat} value={cat}>{cat}</option>
                ))}
              </select>
            )}

            <button
              onClick={handleAddNewCard}
              className="bg-slate-900 text-white hover:bg-slate-800 text-xs font-bold py-1.5 px-3 rounded-lg flex items-center gap-1.5 transition-all cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" /> Nueva Tarjeta
            </button>
          </div>
        </header>

        {/* Content body split into list container and sheet preview live pane */}
        {activeTab === 'editor' ? (
          <div className="flex-1 flex overflow-hidden print:overflow-visible">
            
            {/* LEFT CONTAINER: Interactive cards data inventory */}
            <div className="no-print w-1/2 flex flex-col border-r border-gray-200 bg-white overflow-hidden">
              <div className="p-3 bg-slate-50 border-b border-gray-200 flex flex-col gap-2 shrink-0">
                <div className="flex items-center justify-between text-xs text-gray-600">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-800">Listado de Tarjetas</span>
                    <span className="bg-slate-200 text-slate-700 px-1.5 py-0.5 rounded font-mono font-bold">
                      {filteredCards.length}
                    </span>
                  </div>
                  <div className="text-[10px] text-slate-400 font-bold italic">
                    ¡Se actualizan en tiempo real!
                  </div>
                </div>

                {/* View switcher tabs */}
                <div className="flex items-center justify-between gap-1.5 pt-1">
                  <div className="flex bg-slate-250/50 p-0.5 rounded-lg text-[11px] font-bold">
                    <button
                      type="button"
                      onClick={() => setSidebarViewMode('table')}
                      className={`px-3 py-1 rounded-md transition-all cursor-pointer flex items-center gap-1 ${
                        sidebarViewMode === 'table'
                          ? 'bg-white text-blue-700 shadow-xs'
                          : 'text-slate-650 hover:text-slate-900'
                      }`}
                    >
                      <span>📊 Modo Planilla (Excel)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setSidebarViewMode('visual')}
                      className={`px-3 py-1 rounded-md transition-all cursor-pointer flex items-center gap-1 ${
                        sidebarViewMode === 'visual'
                          ? 'bg-white text-blue-700 shadow-xs'
                          : 'text-slate-650 hover:text-slate-900'
                      }`}
                    >
                      <span>🎴 Modo Diseñador 3D</span>
                    </button>
                  </div>

                  <button
                    onClick={handleAddNewCard}
                    className="bg-blue-600 hover:bg-blue-500 text-white font-extrabold py-1 px-2.5 text-[10px] rounded-lg shadow-xs hover:shadow-md transition-all flex items-center gap-1 cursor-pointer"
                  >
                    <span>➕ Fila Nueva</span>
                  </button>
                </div>
              </div>

              {/* Creador Rápido de Tarjetas por copia y pega */}
              <div className="border-b border-gray-100 bg-slate-50/80 p-4 shrink-0 transition-all duration-150">
                <button
                  id="btn-toggle-bulk-creator"
                  onClick={() => setIsBulkCreatorOpen(!isBulkCreatorOpen)}
                  className={`w-full py-2.5 px-4 rounded-xl border text-xs font-bold flex items-center justify-center gap-2 transition-all ${
                    isBulkCreatorOpen 
                      ? 'bg-amber-50 text-amber-950 border-amber-300 hover:bg-amber-100/75' 
                      : 'bg-blue-50/50 hover:bg-blue-50 text-blue-700 border-blue-100 shadow-sm hover:shadow-md'
                  }`}
                >
                  <Sparkles className="w-4 h-4 text-blue-600 animate-pulse" />
                  {isBulkCreatorOpen ? 'Cerrar Importador Rápido por Lotes' : '✨ Pegar Texto: Creador Rápido Frente y Reverso'}
                </button>

                {isBulkCreatorOpen && (
                  <div className="mt-4 bg-white border border-slate-200 rounded-xl p-4 space-y-4 shadow-sm animate-fadeIn">
                    {/* Mode selector */}
                    <div className="flex items-center justify-between border-b border-gray-100 pb-3">
                      <div className="flex gap-1.5 p-0.5 bg-gray-100 rounded-lg">
                        <button
                          type="button"
                          onClick={() => setBulkCreatorMode('columns')}
                          className={`text-xs px-2.5 py-1 rounded-md font-semibold transition-all cursor-pointer ${
                            bulkCreatorMode === 'columns'
                              ? 'bg-white text-slate-800 shadow-xs font-bold'
                              : 'text-slate-500 hover:text-slate-800'
                          }`}
                        >
                          Frente + atrás (2 Col.)
                        </button>
                        <button
                          type="button"
                          onClick={() => setBulkCreatorMode('separator')}
                          className={`text-xs px-2.5 py-1 rounded-md font-semibold transition-all cursor-pointer ${
                            bulkCreatorMode === 'separator'
                              ? 'bg-white text-slate-800 shadow-xs font-bold'
                              : 'text-slate-500 hover:text-slate-800'
                          }`}
                        >
                          Con Separador (//)
                        </button>
                        <button
                          type="button"
                          onClick={() => setBulkCreatorMode('csv')}
                          className={`text-xs px-2.5 py-1 rounded-md font-semibold transition-all cursor-pointer ${
                            bulkCreatorMode === 'csv'
                              ? 'bg-white text-slate-800 shadow-xs font-bold'
                              : 'text-slate-500 hover:text-slate-800'
                          }`}
                        >
                          📁 Excel / CSV
                        </button>
                      </div>
                      
                      {/* Help icon */}
                      <span className="text-[10px] text-gray-500 italic">
                        {bulkCreatorMode === 'columns' 
                          ? 'Une línea a línea las series.' 
                          : bulkCreatorMode === 'separator'
                          ? 'Usa un divisor en la misma línea.'
                          : 'Copia-pega celdas o arrastra un .csv'}
                      </span>
                    </div>

                    {/* Inputs */}
                    {bulkCreatorMode === 'columns' ? (
                      <div className="grid grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block">
                            Pega ANVERSOS (Frente / Preguntas)
                          </label>
                          <textarea
                            value={bulkFrontsText}
                            onChange={(e) => setBulkFrontsText(e.target.value)}
                            placeholder="Ejemplos (una línea por tarjeta):&#10;¿Qué es la fotosíntesis?&#10;¿Quién descubrió América?&#10;¿Qué es un electrón?"
                            rows={6}
                            className="w-full text-xs font-mono p-2.5 bg-slate-50 border border-gray-200 rounded-lg focus:bg-white focus:ring-1 focus:ring-blue-500/20 focus:border-blue-500 outline-none resize-y"
                          />
                          <div className="text-[9px] text-gray-400 font-medium">
                            Total de líneas para Frente: <strong className="text-slate-600">{bulkFrontsText.trim() ? bulkFrontsText.split('\n').filter(l => l.trim()).length : 0}</strong>
                          </div>
                        </div>

                        <div className="space-y-1">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block">
                            Pega REVERSOS (Atrás / Respuestas)
                          </label>
                          <textarea
                            value={bulkBacksText}
                            onChange={(e) => setBulkBacksText(e.target.value)}
                            placeholder="Ejemplos (una línea por tarjeta):&#10;Proceso químico vegetal&#10;Cristóbal Colón&#10;Partícula subatómica negativa"
                            rows={6}
                            className="w-full text-xs font-mono p-2.5 bg-slate-50 border border-gray-200 rounded-lg focus:bg-white focus:ring-1 focus:ring-blue-500/20 focus:border-blue-500 outline-none resize-y"
                          />
                          <div className="text-[9px] text-gray-400 font-medium">
                            Total de líneas para Reverso: <strong className="text-slate-600">{bulkBacksText.trim() ? bulkBacksText.split('\n').filter(l => l.trim()).length : 0}</strong>
                          </div>
                        </div>
                      </div>
                    ) : bulkCreatorMode === 'separator' ? (
                      <div className="space-y-3">
                        <div className="flex items-center justify-between gap-4">
                          <label className="text-[10px] font-bold text-gray-500 uppercase tracking-wider block">
                            Pega texto de preguntas y respuestas combinados
                          </label>
                          <div className="flex items-center gap-1.5 text-xs">
                            <span className="text-gray-400 text-[10px]">Separador:</span>
                            <input
                              type="text"
                              value={bulkSeparatorChar}
                              onChange={(e) => setBulkSeparatorChar(e.target.value || '//')}
                              className="w-12 text-center text-xs font-mono py-0.5 px-1 bg-gray-50 border border-gray-250 rounded outline-none text-slate-800"
                            />
                          </div>
                        </div>
                        <textarea
                          value={bulkSeparatorText}
                          onChange={(e) => setBulkSeparatorText(e.target.value)}
                          placeholder={`Escribe o pega tus tarjetas en una línea cada una, separando frente de atrás con "${bulkSeparatorChar}".&#10;Ejemplo:&#10;¿Qué es un compuesto? ${bulkSeparatorChar} Sustancia química de dos o más elementos&#10;¿Quién pintó la Mona Lisa? ${bulkSeparatorChar} Leonardo da Vinci`}
                          rows={6}
                          className="w-full text-xs font-mono p-2.5 bg-slate-50 border border-gray-200 rounded-lg focus:bg-white focus:ring-1 focus:ring-blue-500/20 focus:border-blue-500 outline-none resize-y"
                        />
                        <div className="text-[9px] text-gray-400 font-medium">
                          Líneas válidas encontradas: <strong className="text-slate-600">{bulkSeparatorText.split('\n').filter(l => l.trim()).length}</strong>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-3 bg-slate-50/50 p-3 rounded-lg border border-slate-200/60">
                        <div className="flex items-center justify-between gap-3 flex-wrap">
                          <div>
                            <span className="text-[10px] font-bold text-slate-700 uppercase tracking-wider block">
                              Importar Excel o archivo .CSV
                            </span>
                            <span className="text-[9px] text-gray-400 block">
                              Detecta automáticamente formato Excel (Tabulado) o formato estándar CSV (como Comas `,` o Puntos y coma `;`)
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={handleDownloadSampleCsv}
                            className="text-[10px] text-blue-600 hover:text-blue-700 font-bold bg-blue-100/60 hover:bg-blue-100 py-1 px-2.5 rounded-lg transition-all cursor-pointer flex items-center gap-1 shrink-0"
                          >
                            <span>📥 Descargar Plantilla .CSV</span>
                          </button>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                          {/* File Uploader area */}
                          <div 
                            onClick={() => csvFileInputRef.current?.click()}
                            className="border border-dashed border-gray-300 hover:border-blue-500 rounded-lg p-3 text-center bg-white hover:bg-slate-50/80 cursor-pointer transition-colors flex flex-col items-center justify-center min-h-[100px]"
                          >
                            <FolderOpen className="w-5 h-5 text-blue-500 mb-1" />
                            <p className="text-[10px] font-bold text-slate-700">Subir Excel (.xlsx) o .CSV</p>
                            <p className="text-[8px] text-gray-400">Archivos Excel, Calc, CSV, TXT</p>
                            <input
                              type="file"
                              ref={csvFileInputRef}
                              accept=".xlsx,.xls,.csv,.tsv,.txt"
                              onChange={handleCsvFileLoad}
                              className="hidden"
                            />
                          </div>

                          {/* Textarea for direct Excel paste */}
                          <div className="md:col-span-2 space-y-1">
                            <span className="text-[10px] font-bold text-gray-500 block uppercase">
                              O copia celdas de Excel y pégalas aquí directamente:
                            </span>
                            <textarea
                              value={bulkCsvText}
                              onChange={(e) => setBulkCsvText(e.target.value)}
                              placeholder="Frente / Pregunta &#9; Reverso / Respuesta &#10;Pregunta de ejemplo 1&#9;Respuesta de ejemplo 1&#10;Pregunta de ejemplo 2&#9;Respuesta de ejemplo 2"
                              rows={4}
                              className="w-full text-[11px] font-mono p-2 bg-white border border-gray-250 rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 focus:bg-white outline-none resize-none"
                            />
                          </div>
                        </div>

                        {/* Excel sheet selection if multiple sheets exist */}
                        {excelSheetNames.length > 1 && (
                          <div className="bg-blue-50/50 border border-blue-200/50 rounded-lg p-2.5 flex items-center justify-between gap-3 text-left mt-1.5 animate-fadeIn">
                            <div>
                              <p className="text-[10px] font-bold text-blue-800 uppercase tracking-wide flex items-center gap-1">
                                <span>📄</span> Multi-pestanas detectadas en el Excel
                              </p>
                              <p className="text-[9px] text-blue-600 block leading-normal">
                                Tu archivo posee varias hojas. Elige cuál quieres importar para tus tarjetas:
                              </p>
                            </div>
                            <select
                              value={excelActiveSheetName}
                              onChange={(e) => handleExcelSheetChange(e.target.value)}
                              className="text-[11px] font-bold text-slate-800 bg-white border border-gray-200 rounded-lg py-1 px-2.5 outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 cursor-pointer"
                            >
                              {excelSheetNames.map(name => (
                                <option key={name} value={name}>{name}</option>
                              ))}
                            </select>
                          </div>
                        )}

                        {bulkCsvText && (
                          <div className="flex items-center justify-between text-[11px] bg-green-50 text-green-800 p-2 border border-green-200/50 rounded-lg mt-1">
                            <span className="font-medium">
                              📝 ¡Texto cargado correctamente! Detectamos <strong>{parseCsvOrTabSeparated(bulkCsvText).length}</strong> registros listos para organizar.
                            </span>
                            <button
                              type="button"
                              onClick={() => setBulkCsvText('')}
                              className="text-red-500 hover:underline font-bold text-[10px]"
                            >
                              Limpiar
                            </button>
                          </div>
                        )}
                      </div>
                    )}

                    {/* Category Selector and Generation Preview */}
                    <div className="grid grid-cols-2 gap-3 items-center pt-2 border-t border-gray-100">
                      <div>
                        <label className="text-[10px] text-gray-500 uppercase tracking-wider font-bold block mb-1">
                          Categoría / Materia de destino:
                        </label>
                        <input
                          type="text"
                          value={bulkCategory}
                          onChange={(e) => setBulkCategory(e.target.value)}
                          placeholder="ej: Geografía, Examen, Farmacia..."
                          className="w-full text-xs py-1.5 px-2.5 bg-slate-50 border border-gray-200 rounded-lg outline-none focus:bg-white focus:border-blue-500 text-slate-800 font-medium"
                        />
                      </div>

                      {/* Align Check and Total Cards Preview indicator */}
                      <div className="text-right">
                        {bulkCreatorMode === 'columns' && (
                          <div className="inline-block text-[11px] mb-1 text-left">
                            {(() => {
                              const fCount = bulkFrontsText.trim() ? bulkFrontsText.split('\n').filter(l => l.trim()).length : 0;
                              const bCount = bulkBacksText.trim() ? bulkBacksText.split('\n').filter(l => l.trim()).length : 0;
                              if (fCount === 0 && bCount === 0) {
                                return <span className="text-gray-400">Sin datos de entrada</span>;
                              } else if (fCount === bCount) {
                                return <span className="text-green-600 font-semibold flex items-center gap-1">✓ ¡Columnas Alineadas! ({fCount} pares)</span>;
                              } else {
                                return (
                                  <span className="text-amber-600 font-medium leading-tight block">
                                    ⚠️ No coinciden ({fCount} ftes. / {bCount} dros.). Se completarán vacíos.
                                  </span>
                                );
                              }
                            })()}
                          </div>
                        )}
                        <p className="text-[10px] text-gray-400">
                          Se crearán un total de <strong className="text-slate-800 text-xs">{getBulkCardsPreview().length}</strong> tarjetas.
                        </p>
                      </div>
                    </div>

                    {/* Table-like Live Parser Preview to ensure user is confident they paste it correctly! */}
                    {getBulkCardsPreview().length > 0 && (
                      <>
                        <div className="bg-slate-50 border border-slate-100 rounded-lg p-2.5 max-h-32 overflow-y-auto custom-scrollbar">
                          <p className="text-[9px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
                            Vista previa en tiempo real de lo que vas a importar:
                          </p>
                          <table className="w-full text-[10px] text-left border-collapse">
                            <thead>
                              <tr className="border-b border-gray-200">
                                <th className="py-1 text-slate-400 font-normal w-6">#</th>
                                <th className="py-1 text-slate-400 font-normal">Anverso (Frente / Pregunta)</th>
                                <th className="py-1 text-slate-400 font-normal">Reverso (Atrás / Respuesta)</th>
                              </tr>
                            </thead>
                            <tbody>
                              {getBulkCardsPreview().map((item, idx) => (
                                <tr key={idx} className="border-b border-gray-100 hover:bg-slate-50 last:border-0">
                                  <td className="py-1 text-slate-400 font-mono">{idx + 1}</td>
                                  <td className="py-1 text-slate-800 truncate max-w-[120px] pr-2" title={item.front}>
                                    {item.front || <span className="text-red-400 bg-red-50 px-1 py-0.2 rounded font-semibold italic">Vacío</span>}
                                  </td>
                                  <td className="py-1 text-slate-600 truncate max-w-[120px]" title={item.back}>
                                    {item.back || <span className="text-red-400 bg-red-50 px-1 py-0.2 rounded font-semibold italic">Vacío</span>}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* Informative Step-Guide Banner */}
                        <div className="bg-blue-50 border border-blue-200 text-blue-900 rounded-xl p-3 text-left text-xs space-y-1 shadow-xs border-l-4 border-l-blue-600 animate-[pulse_2.5s_infinite]">
                          <p className="font-extrabold flex items-center gap-1.5 text-blue-950">
                            <span>📣</span> ¡Borrador Cargado Correctamente!
                          </p>
                          <p className="text-[11px] text-blue-800 leading-normal">
                            Para transferir estas tarjetas al mazo principal y visualizarlas en las hojas de impresión de la derecha, presiona el botón azul de abajo: <strong className="text-blue-950">"Crear {getBulkCardsPreview().length} Tarjetas en Lote"</strong>.
                          </p>
                        </div>
                      </>
                    )}

                    {/* Submit Action Block */}
                    <div className="flex gap-2 justify-end pt-1">
                      <button
                        type="button"
                        onClick={() => {
                          setBulkFrontsText('');
                          setBulkBacksText('');
                          setBulkSeparatorText('');
                        }}
                        className="py-1.5 px-3 bg-gray-100 hover:bg-gray-200 text-slate-700 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
                      >
                        Limpiar Textos
                      </button>
                      <button
                        type="button"
                        onClick={handleCreateBulkCards}
                        className={`py-2 px-4 text-xs font-bold rounded-lg transition-all shadow-md flex items-center gap-1.5 cursor-pointer ${
                          getBulkCardsPreview().length > 0 
                            ? 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white hover:shadow-lg hover:scale-[1.02] ring-2 ring-blue-500/20'
                            : 'bg-gray-200 text-gray-400 cursor-not-allowed'
                        }`}
                        disabled={getBulkCardsPreview().length === 0}
                      >
                        <Plus className="w-4 h-4" />
                        Crear {getBulkCardsPreview().length} Tarjetas en Lote
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Infinite list of cards */}
              <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar">
                {filteredCards.length === 0 ? (
                  <div className="text-center py-12 px-6 border-2 border-dashed border-gray-200 rounded-xl max-w-md mx-auto my-6 space-y-3">
                    <p className="text-xs text-gray-400 font-medium">No se encontraron tarjetas que coincidan con los filtros.</p>
                    <button
                      onClick={handleAddNewCard}
                      className="bg-blue-600 text-white text-xs font-bold py-2 px-4 rounded-lg inline-flex items-center gap-1.5 shadow-sm cursor-pointer"
                    >
                      <Plus className="w-3.5 h-3.5" /> Agregar Nueva Tarjeta
                    </button>
                  </div>
                ) : sidebarViewMode === 'table' ? (
                  <div className="space-y-3">
                    <div className="bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200/80 rounded-xl p-3 text-xs shadow-xs shrink-0 select-none">
                      <div className="font-extrabold text-blue-900">Edicion rapida de tarjetas</div>
                      <div className="text-[10px] text-blue-700/80 leading-normal mt-1">
                        Esta vista es totalmente local: edita frente, reverso y categoria directamente en la tabla sin depender de servicios externos.
                      </div>
                    </div>

                    <div className="bg-slate-50 border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                      <div className="overflow-x-auto">
                        <table className="w-full text-left text-xs border-collapse">
                          <thead className="bg-[#f1f5f9] border-b border-slate-200 text-slate-700 font-extrabold text-[10px] uppercase tracking-wider select-none">
                            <tr>
                              <th className="py-2.5 px-2.5 w-10 text-center font-bold">N°</th>
                              <th className="py-2.5 px-3">Anverso (Frente / Pregunta)</th>
                              <th className="py-2.5 px-3">Reverso (Atrás / Respuesta)</th>
                              <th className="py-2.5 px-2 w-24">Materia</th>
                              <th className="py-2.5 px-2 w-16 text-center">Acción</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-gray-200/85 bg-white">
                            {filteredCards.map((card, idx) => (
                              <tr key={card.id} className="hover:bg-blue-50/20 transition-colors">
                                {/* 1. Badge / Number index */}
                                <td className="py-2 px-1 text-center font-mono font-bold text-slate-500">
                                  <span className="inline-block bg-slate-100 text-slate-700 rounded-md px-1.5 py-0.5 text-[10px]">
                                    #{idx + 1}
                                  </span>
                                </td>
                                
                                {/* 2. Front content input */}
                                <td className="py-1.5 px-2">
                                  <textarea
                                    rows={1}
                                    value={card.frontText}
                                    onChange={(e) => {
                                      handleUpdateCard({ ...card, frontText: e.target.value });
                                    }}
                                    placeholder="Escribe la pregunta..."
                                    className="w-full text-xs font-mono p-1 bg-transparent hover:bg-slate-50/50 border border-transparent hover:border-slate-200 focus:bg-white focus:border-blue-500 focus:ring-1 focus:ring-blue-500/10 rounded-md outline-none transition-all resize-none min-h-[30px]"
                                    style={{ height: 'auto' }}
                                  />
                                </td>
                                
                                {/* 3. Back content input */}
                                <td className="py-1.5 px-2 relative">
                                  <textarea
                                    rows={1}
                                    value={card.backText}
                                    onChange={(e) => {
                                      handleUpdateCard({ ...card, backText: e.target.value });
                                    }}
                                    placeholder="Escribe la respuesta..."
                                    className="w-full text-xs font-mono p-1 bg-transparent hover:bg-slate-50/50 border border-transparent hover:border-slate-200 focus:bg-white focus:border-blue-500 focus:ring-1 focus:ring-blue-500/10 rounded-md outline-none transition-all resize-none min-h-[30px]"
                                    style={{ height: 'auto' }}
                                  />
                                </td>
                                
                                {/* 4. Category category */}
                                <td className="py-1.5 px-2">
                                  <input
                                    type="text"
                                    value={card.category || 'Generales'}
                                    onChange={(e) => {
                                      handleUpdateCard({ ...card, category: e.target.value });
                                    }}
                                    placeholder="Categoría"
                                    className="w-full text-[10px] font-bold text-slate-700 p-1 bg-transparent hover:bg-slate-50/50 border border-transparent hover:border-slate-200 focus:bg-white focus:border-blue-500 rounded-md outline-none"
                                  />
                                </td>
                                
                                {/* 5. Row action triggers */}
                                <td className="py-1.5 px-2">
                                  <div className="flex items-center justify-center gap-1">
                                    {/* Duplicate */}
                                    <button
                                      onClick={() => handleDuplicateCard(card)}
                                      className="p-1 hover:bg-slate-100 rounded text-slate-400 hover:text-slate-650 transition-colors cursor-pointer"
                                      title="Duplicar Fila"
                                    >
                                      <Copy className="w-3 h-3" />
                                    </button>
                                    {/* Delete row */}
                                    <button
                                      onClick={() => handleDeleteCard(card.id)}
                                      className="p-1 hover:bg-red-50 rounded text-red-400 hover:text-red-500 transition-colors cursor-pointer"
                                      title="Eliminar Fila"
                                    >
                                      <Trash2 className="w-3 h-3" />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {/* Footguide */}
                      <div className="bg-amber-50/60 p-2.5 border-t border-slate-200 text-amber-950 text-[10px] font-medium leading-relaxed flex items-center justify-between">
                        <div className="flex items-center gap-1.5">
                          <span>💡 <strong>¡Escribe aquí en dos columnas directamente!</strong> Los cambios se guardan y recalculan al instante.</span>
                        </div>
                        <div className="text-slate-400 italic">
                          ¡Al presionar Tab puedes moverte rápido por fila!
                        </div>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-4">
                    {filteredCards.map((card, idx) => (
                      <CardItem
                        key={card.id}
                        card={card}
                        index={idx}
                        onUpdate={handleUpdateCard}
                        onDelete={handleDeleteCard}
                        onDuplicate={handleDuplicateCard}
                      />
                    ))}
                  </div>
                )}
              </div>
            </div>

            {/* RIGHT CONTAINER: Real scale Sheet visualizer */}
            <div className="flex-1 bg-slate-200 p-6 flex flex-col items-center justify-start overflow-y-auto custom-scrollbar relative print:p-0 print:bg-white print:overflow-visible">
              
              {/* Dynamic instruction box floating */}
              <div className="no-print w-[210mm] max-w-full mb-4 bg-amber-50 border border-amber-200/60 rounded-xl p-4 text-xs text-amber-900 leading-relaxed shadow-xs flex gap-3 items-start">
                <Info className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <p className="font-bold">💡 Impresión Doble Faz Manual de Máxima Alineación:</p>
                  <p>
                    Para recortar sin descuadres, hemos implementado el <b>espejado automático</b> para imprimir reverso/dorso.
                    En la barra del previsualizador, puedes alternar entre <b>Anverso (Frente)</b> y <b>Reverso (Atrás)</b> para examinar las carillas antes de imprimir. El diseño se centra milimétricamente.
                  </p>
                </div>
              </div>

              {/* Interactive Toolbar for Preview Sheet */}
              <div className="no-print w-[210mm] max-w-full bg-white border border-gray-200 p-3.5 rounded-xl mb-4 shadow-xs flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-slate-500 uppercase">Previsualizar Carilla:</span>
                  <div className="bg-gray-100 p-0.5 rounded-lg flex">
                    <button
                      onClick={() => setPreviewSide('front')}
                      className={`text-xs py-1 px-3 rounded-md font-semibold transition-all flex items-center gap-1 cursor-pointer ${
                        previewSide === 'front'
                          ? 'bg-blue-600 text-white shadow-xs'
                          : 'text-gray-500 hover:text-slate-800'
                      }`}
                    >
                      <Layers className="w-3.5 h-3.5" /> Frente (Anverso)
                    </button>
                    <button
                      onClick={() => setPreviewSide('back')}
                      className={`text-xs py-1 px-3 rounded-md font-semibold transition-all flex items-center gap-1 cursor-pointer ${
                        previewSide === 'back'
                          ? 'bg-blue-600 text-white shadow-xs'
                          : 'text-gray-500 hover:text-slate-800'
                      }`}
                    >
                      <RotateCw className="w-3.5 h-3.5" /> Reverso (Atrás Espejado)
                    </button>
                  </div>
                </div>

                <div className="flex items-center gap-3">
                  {/* View all hojas stacked toggle */}
                  <div className="flex bg-gray-100 p-0.5 rounded-lg border border-gray-200">
                    <button
                      type="button"
                      onClick={() => setPreviewAllPages(false)}
                      className={`text-[10px] uppercase tracking-wider font-extrabold px-3 py-1 rounded-md transition-all cursor-pointer ${
                        !previewAllPages
                          ? 'bg-white text-slate-800 shadow-xs font-bold'
                          : 'text-slate-550 hover:text-slate-800'
                      }`}
                      title="Ver una hoja a la vez con controles de navegación"
                    >
                      Una Hoja
                    </button>
                    <button
                      type="button"
                      onClick={() => setPreviewAllPages(true)}
                      className={`text-[10px] uppercase tracking-wider font-extrabold px-3 py-1 rounded-md transition-all cursor-pointer ${
                        previewAllPages
                          ? 'bg-white text-slate-800 shadow-xs font-bold'
                          : 'text-slate-550 hover:text-slate-800'
                      }`}
                      title="Muestra todas las páginas generadas en una lista continua para que puedas scrollearlas juntas"
                    >
                      Ver Todas ({totalPagesNeeded})
                    </button>
                  </div>

                  {/* Persistent Add sheet button right on the main preview board! */}
                  <button
                    onClick={handleAddNewPrintSheet}
                    className="text-[11px] font-bold text-blue-700 hover:text-white bg-blue-50 hover:bg-blue-600 border border-blue-200 hover:border-blue-600 py-1.5 px-3 rounded-xl transition-all flex items-center gap-1 cursor-pointer hover:shadow-sm"
                    title={`Añade un juego completo de ${cardsPerPage} tarjetas vacías para rellenar en una nueva hoja de impresión`}
                  >
                    <span>📄+</span> Añadir Hoja
                  </button>

                  {!previewAllPages && totalPagesNeeded > 1 && (
                    <div className="flex items-center gap-2 border-l border-gray-200 pl-3">
                      <span className="text-[11px] font-bold text-slate-650">
                        Hoja: {currentPageIndex + 1} de {totalPagesNeeded}
                      </span>
                      <div className="flex border rounded-lg overflow-hidden border-gray-200 text-xs font-bold bg-gray-50">
                        <button
                          onClick={() => setCurrentPageIndex(prev => Math.max(0, prev - 1))}
                          disabled={currentPageIndex === 0}
                          className="px-2 py-1 hover:bg-gray-100 disabled:opacity-40 cursor-pointer disabled:cursor-not-allowed"
                        >
                          ←
                        </button>
                        <button
                          onClick={() => setCurrentPageIndex(prev => Math.min(totalPagesNeeded - 1, prev + 1))}
                          disabled={currentPageIndex === totalPagesNeeded - 1}
                          className="px-2 py-1 hover:bg-gray-100 disabled:opacity-40 cursor-pointer disabled:cursor-not-allowed"
                        >
                          →
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* Actually render the A4 sheets list for Print, or the current selected active interactive page for OnScreen Visualization */}
              <div className="print-layout-container flex flex-col gap-8 w-full items-center print:gap-0">
                {/* At print time, render ALL sheets sequentially so everything prints correctly with breaks! */}
                {/* On-screen preview, we render the current active page for speed/user friendliness */}
                <div className="hidden print:block w-full">
                  {Array.from({ length: totalPagesNeeded }).map((_, pageIdx) => (
                    <div key={`print-page-wrapper-front-${pageIdx}`} className="print-page-break-container">
                      <PrintSheet
                        cards={flattenedPrintCards}
                        config={config}
                        side="front"
                        pageIndex={pageIdx}
                        totalPages={totalPagesNeeded}
                      />
                      <PrintSheet
                        cards={flattenedPrintCards}
                        config={config}
                        side="back"
                        pageIndex={pageIdx}
                        totalPages={totalPagesNeeded}
                      />
                    </div>
                  ))}
                </div>

                {/* On interactive screen, render just the selected filter page */}
                <div className="print:hidden">
                  {totalPagesNeeded > 0 ? (
                    previewAllPages ? (
                      <div className="flex flex-col gap-6 w-full items-center">
                        {Array.from({ length: totalPagesNeeded }).map((_, pageIdx) => (
                          <div key={`on-screen-page-stacked-${pageIdx}`} className="relative bg-white rounded-2xl shadow-md border border-gray-200 overflow-hidden flex flex-col items-center p-2.5 pt-8 hover:border-blue-400 hover:shadow-lg transition-all group">
                            {/* Page header banner indicator */}
                            <div className="absolute top-2 left-4 text-[10px] font-extrabold text-blue-800 bg-blue-50 px-2.5 py-0.5 rounded-md select-none">
                              HOJA DE IMPRESIÓN {pageIdx + 1} de {totalPagesNeeded}
                            </div>
                            <div className="absolute top-2 right-4 text-[10px] font-bold text-gray-400 group-hover:text-blue-600 transition-colors">
                              Frente y Reverso Alineados
                            </div>
                            <PrintSheet
                              cards={flattenedPrintCards}
                              config={config}
                              side={previewSide}
                              pageIndex={pageIdx}
                              totalPages={totalPagesNeeded}
                            />
                          </div>
                        ))}
                      </div>
                    ) : (
                      <PrintSheet
                        cards={flattenedPrintCards}
                        config={config}
                        side={previewSide}
                        pageIndex={currentPageIndex}
                        totalPages={totalPagesNeeded}
                      />
                    )
                  ) : (
                    <div className="w-[210mm] max-w-full aspect-[1/1.414] bg-white rounded-xl shadow-md border border-gray-200 flex flex-col items-center justify-center p-12 text-center text-gray-400 space-y-2">
                      <Grid className="w-12 h-12 text-gray-300" />
                      <p className="text-sm font-semibold">No se han ingresado tarjetas para renderizar las hojas.</p>
                      <p className="text-xs">Usa el botón "Agregar Nueva Tarjeta" en el panel de arriba o importa tus datos desde Excel/CSV.</p>
                    </div>
                  )}
                </div>
              </div>

              {/* Status Footer at end */}
              <div className="no-print w-[210mm] max-w-full bg-slate-900 border border-slate-800 text-slate-400 text-[10px] uppercase tracking-wider py-2 px-4 rounded-xl mt-4 flex items-center justify-between">
                <div className="flex gap-4">
                  <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-green-500 animate-pulse"></span> MODO LOCAL</span>
                  <span>PRESTACIONES: 300 DPI ALTA DENSIDAD</span>
                </div>
                <div>
                  Fórmula de cuadrícula: <b className="text-slate-200">{cols} columnas x {rows} filas</b>
                </div>
              </div>
            </div>
          </div>
        ) : (
          /* INSTRUCTIONS MANUAL TAB CONTENT with interactive guided duplex wizard */
          <div className="no-print flex-1 overflow-y-auto p-8 max-w-5xl mx-auto space-y-8">
            {/* Elegant Header Hero */}
            <div className="bg-slate-900 text-white rounded-2xl p-6 shadow-xl border border-slate-800 relative overflow-hidden flex flex-col md:flex-row gap-6 justify-between items-center">
              <div className="absolute right-0 top-0 translate-x-12 -translate-y-12 w-64 h-64 rounded-full bg-blue-500 opacity-10 blur-3xl pointer-events-none" />
              <div className="max-w-2xl text-left">
                <span className="bg-blue-500/20 text-blue-300 font-mono font-bold text-[10px] py-1 px-2.5 rounded-full uppercase tracking-wider">
                  Guía de Impresión Profesional
                </span>
                <h2 className="text-2xl font-bold mt-3 text-slate-100 flex items-center gap-2">
                  <span>🖨️</span> Manual de Doble Cara y Alimentación de Papel
                </h2>
                <p className="text-slate-400 text-sm mt-2 leading-relaxed">
                  Para que las preguntas y respuestas queden perfectamente alineadas al dar vuelta la hoja, el motor reordena e invierte de izquierda a derecha la cuadrícula trasera. Sigue este manual dinámico para evitar desperdicio de papel y lograr un acabado milimétrico.
                </p>
              </div>
              <button
                onClick={() => {
                  setActiveTab('editor');
                  setPreviewSide('front');
                }}
                className="bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs py-2 px-4 rounded-xl shadow-md cursor-pointer transition-all shrink-0 hover:shadow-lg flex items-center gap-1.5"
              >
                <span>Ir a Previsualizar</span> →
              </button>
            </div>

            {/* Micro Simulator Configurator */}
            <div className="bg-white border text-left border-gray-200 rounded-xl p-5 shadow-xs grid grid-cols-1 md:grid-cols-3 gap-6 items-start">
              <div className="space-y-4 md:col-span-1">
                <h3 className="font-bold text-slate-800 text-sm flex items-center gap-1.5 pb-2 border-b border-gray-100">
                  <span className="bg-blue-100 text-blue-700 w-5 h-5 rounded-full flex items-center justify-center text-[10px]">A</span>
                  1. Configura tu Impresora:
                </h3>
                
                {/* Selector 1 */}
                <div className="space-y-1.5">
                  <label className="text-[10px] uppercase tracking-wider text-slate-400 font-extrabold block">
                    ¿Cómo alimenta papel tu impresora?
                  </label>
                  <div className="grid grid-cols-1 gap-1.5">
                    <button
                      onClick={() => setPrinterType('drawer')}
                      className={`text-slate-700 font-medium text-xs p-2.5 rounded-lg border text-left flex items-start gap-2.5 transition-all cursor-pointer ${
                        printerType === 'drawer'
                          ? 'border-blue-500 bg-blue-50/50 ring-1 ring-blue-500/20'
                          : 'border-gray-200 hover:bg-slate-50'
                      }`}
                    >
                      <span className="text-base mt-0.5">📥</span>
                      <div>
                        <p className="font-bold text-slate-800">Cajón Inferior / Bandeja Cerrada</p>
                        <p className="text-[10px] text-gray-500 leading-tight block mt-0.5">
                          Típica láser. El papel está adentro, se aspira hacia el fondo simulando una "U", imprimiendo en la cara que mira haca abajo.
                        </p>
                      </div>
                    </button>

                    <button
                      onClick={() => setPrinterType('top')}
                      className={`text-slate-700 font-medium text-xs p-2.5 rounded-lg border text-left flex items-start gap-2.5 transition-all cursor-pointer ${
                        printerType === 'top'
                          ? 'border-blue-500 bg-blue-50/50 ring-1 ring-blue-500/20'
                          : 'border-gray-200 hover:bg-slate-50'
                      }`}
                    >
                      <span className="text-base mt-0.5">📤</span>
                      <div>
                        <p className="font-bold text-slate-800">Alimentador Superior / Trasero</p>
                        <p className="text-[10px] text-gray-500 leading-tight block mt-0.5">
                          Típica chorro de tinta (Inkjet). Las hojas entran paradas desde arriba y bajan de manera plana para salir impresas por abajo.
                        </p>
                      </div>
                    </button>
                  </div>
                </div>
              </div>

              <div className="space-y-4 md:col-span-1">
                <h3 className="font-bold text-slate-800 text-sm flex items-center gap-1.5 pb-2 border-b border-gray-100">
                  <span className="bg-blue-100 text-blue-700 w-5 h-5 rounded-full flex items-center justify-center text-[10px]">B</span>
                  2. Método de Volteo Elegido:
                </h3>

                {/* Flip Direction Selector */}
                <div className="space-y-1.5">
                  <label className="text-[10px] uppercase tracking-wider text-slate-400 font-extrabold block">
                    Dirección de Giro al Devolver el Papel:
                  </label>
                  <div className="grid grid-cols-1 gap-1.5">
                    <button
                      onClick={() => setFlipDirection('long')}
                      className={`font-medium text-xs p-2.5 rounded-lg border text-left flex items-start gap-2.5 transition-all cursor-pointer ${
                        flipDirection === 'long'
                          ? 'border-blue-500 bg-blue-50/50 ring-1 ring-blue-500/20'
                          : 'border-gray-200 hover:bg-slate-50'
                      }`}
                    >
                      <span className="text-base mt-0.5">📖</span>
                      <div>
                        <p className="font-bold text-slate-800">Giro en Borde Largo (Estilo Libro)</p>
                        <p className="text-[10px] text-gray-500 leading-tight block mt-0.5">
                          Giras la hoja horizontalmente hacia el lado izquierdo, tal como pasas la página de un libro comercial standard.
                        </p>
                      </div>
                    </button>

                    <button
                      onClick={() => setFlipDirection('short')}
                      className={`font-medium text-xs p-2.5 rounded-lg border text-left flex items-start gap-2.5 transition-all cursor-pointer ${
                        flipDirection === 'short'
                          ? 'border-blue-500 bg-blue-50/50 ring-1 ring-blue-500/20'
                          : 'border-gray-200 hover:bg-slate-50'
                      }`}
                    >
                      <span className="text-base mt-0.5">🪙</span>
                      <div>
                        <p className="font-bold text-slate-800">Giro en Borde Corto (Estilo Moneda)</p>
                        <p className="text-[10px] text-gray-500 leading-tight block mt-0.5">
                          Volteas la hoja sobre su eje horizontal (de arriba hacia abajo), igual que al girar una moneda de arriba hacia abajo.
                        </p>
                      </div>
                    </button>
                  </div>
                </div>
              </div>

              {/* 3D Paper Flip Visualizer Widget */}
              <div className="md:col-span-1 bg-slate-50 border border-slate-200 rounded-xl p-4 flex flex-col justify-between h-full min-h-[220px]">
                <div>
                  <span className="text-[9px] font-bold text-blue-600 bg-blue-50 border border-blue-200/50 px-2 py-0.5 rounded-full block w-fit mb-1.5">
                    Orientador 3D Activo
                  </span>
                  <h4 className="font-bold text-slate-800 text-[11px] uppercase tracking-wider">
                    ¿Cómo rotar la hoja antes de la pasada 2?
                  </h4>
                  <p className="text-[10px] text-slate-500 leading-tight block mt-1">
                    Mira la animación interactiva de abajo que representa tu hoja A4 real.
                  </p>
                </div>

                {/* Simulated Sheet of paper flipping with CSS */}
                <div className="flex items-center justify-center p-3">
                  <div className="perspective-1000 w-[70px] h-[100px]">
                    <div 
                      className={`relative w-full h-full preserve-3d h-full rounded border border-gray-300 shadow-sm transition-transform duration-1000 flex items-center justify-center text-center font-bold text-[10px] ${
                        flipDirection === 'long' 
                          ? 'animate-[spin_4s_infinite_linear]' 
                          : 'animate-[spinShort_4s_infinite_linear]'
                      }`}
                    >
                      {/* CSS custom keyframe utility styled here internally with style tag for portability */}
                      <style>{`
                        @keyframes spin {
                          0% { transform: rotateY(0deg); background-color: #ffffff; }
                          45% { transform: rotateY(180deg); background-color: #eff6ff; }
                          55% { transform: rotateY(180deg); background-color: #eff6ff; }
                          100% { transform: rotateY(360deg); background-color: #ffffff; }
                        }
                        @keyframes spinShort {
                          0% { transform: rotateX(0deg); background-color: #ffffff; }
                          45% { transform: rotateX(180deg); background-color: #eff6ff; }
                          50% { transform: rotateX(180deg); background-color: #eff6ff; }
                          100% { transform: rotateX(360deg); background-color: #ffffff; }
                        }
                      `}</style>
                      <div className="absolute inset-0 flex flex-col items-center justify-center bg-white border rounded">
                        <span className="text-[9px] text-slate-400">PASADA 1</span>
                        <span className="text-xs font-bold text-slate-800">Frente</span>
                        <span className="text-[8px] text-blue-600 font-mono">↑ ARRIBA ↑</span>
                      </div>
                    </div>
                  </div>
                </div>

                <p className="text-[9px] text-center text-slate-400 italic">
                  Eje: {flipDirection === 'long' ? 'Giro Horizontal (Y)' : 'Giro de Cabeza (X)'}
                </p>
              </div>
            </div>

            {/* Foolproof Marker Test Segment */}
            <div className="bg-amber-50/55 border border-amber-200 rounded-xl p-4 flex gap-4 text-left items-start">
              <span className="text-xl shrink-0">💡</span>
              <div className="space-y-1">
                <h4 className="font-bold text-amber-900 text-xs uppercase tracking-wider">
                  🧪 El truco infalible de la "Flecha de Lápiz" (Recomendado para primer uso)
                </h4>
                <p className="text-amber-800 text-[11px] leading-relaxed">
                  Para no dudar de cómo traga el papel tu impresora: toma una hoja blanca y, antes de meterla, dibuja una <strong>flecha pequeña de lápiz apuntando hacia arriba (↑) en la esquina superior del lado que queda a la vista</strong>. Haz una prueba imprimiendo solo una tarjeta. Cuando salga la hoja, sabrás exactamente de qué lado imprime y con qué orientación física debe reingresar para el reverso.
                </p>
              </div>
            </div>

            {/* Dynamic Step-by-Step Print Process Instructions */}
            <div className="space-y-4 text-left">
              <h3 className="font-extrabold text-slate-800 text-sm uppercase tracking-wider">
                Procedimiento Paso a Paso de Impresión en 2 Pasadas
              </h3>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                
                {/* Step 1 */}
                <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-xs flex flex-col justify-between space-y-4">
                  <div className="space-y-2.5">
                    <span className="bg-blue-100 text-blue-800 text-[11px] font-extrabold font-mono py-0.5 px-2 rounded-full block w-fit">
                      PASO 1
                    </span>
                    <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider flex items-center gap-1.5">
                      <span>1️⃣</span> Impresión de los Frentes
                    </h4>
                    <p className="text-slate-500 text-xs leading-relaxed">
                      Haz clic en <strong>"Imprimir Tarjetas"</strong> o presiona <kbd className="bg-gray-100 border px-1 rounded text-[10px] font-mono">Ctrl+P</kbd>. 
                    </p>
                    <div className="bg-slate-50 border border-slate-100 rounded-lg p-2.5 space-y-1 text-[11px]">
                      <span className="font-bold text-slate-800 block">En la ventana del Navegador:</span>
                      <ul className="list-disc pl-4 space-y-1 text-slate-600 font-medium">
                        <li>Selecciona <strong>Páginas Impares</strong> en el rango de páginas (o escribe manualmente la secuencia <code className="bg-white px-1 border rounded text-[10px] font-bold">1, 3, 5, 7...</code>).</li>
                        <li>Configura escala al <strong>100% (Real)</strong>.</li>
                        <li>Establece márgenes en <strong>"Ninguno"</strong> o "Mínimos".</li>
                      </ul>
                    </div>
                  </div>
                  <div className="text-[10px] text-green-700 bg-green-50 border border-green-100 rounded-lg p-2">
                    ✓ Esto imprimirá todas las carillas con preguntas en papel rígido.
                  </div>
                </div>

                {/* Step 2 */}
                <div className="bg-white border border-blue-200/80 bg-gradient-to-b from-blue-50/5 to-blue-50/20 rounded-xl p-5 shadow-xs flex flex-col justify-between space-y-4">
                  <div className="space-y-2.5">
                    <span className="bg-amber-100 text-amber-800 text-[11px] font-extrabold font-mono py-0.5 px-2 rounded-full block w-fit">
                      PASO 2
                    </span>
                    <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider flex items-center gap-1.5">
                      <span>2️⃣</span> La Técnica de Giro Exacto
                    </h4>
                    <p className="text-slate-500 text-xs leading-relaxed">
                      Retira las hojas del final del frente impreso y devuélvelas a la bandeja de entrada siguiendo la orientación que determinaron tus elecciones físicas:
                    </p>

                    {/* Highly Targeted Instruction based on interactive selection above */}
                    <div className="bg-white border border-blue-200 rounded-lg p-3 space-y-2 text-[11px] shadow-xs">
                      <p className="font-bold text-blue-900 border-b border-blue-100 pb-1 flex items-center gap-1">
                        👉 Instrucción para tu impresora:
                      </p>
                      
                      {printerType === 'drawer' ? (
                        flipDirection === 'long' ? (
                          <div className="space-y-1.5 text-slate-700">
                            <p className="font-medium">Para <strong>Bandejas de Cajón Inferior</strong> con <strong>Giro Libro</strong>:</p>
                            <p className="italic bg-slate-50 rounded p-1.5 border border-slate-100 text-[10px]">
                              "Toma las hojas. El frente impreso, que antes miraba hacia arriba, ahora debe ponerse <strong>boca abajo (mirando al piso del cajón)</strong>. El borde superior de la hoja (la cabecera) debe seguir apuntando <strong>hacia el fondo</strong> del cajón."
                            </p>
                          </div>
                        ) : (
                          <div className="space-y-1.5 text-slate-700">
                            <p className="font-medium">Para <strong>Bandejas de Cajón Inferior</strong> con <strong>Giro Moneda</strong>:</p>
                            <p className="italic bg-slate-50 rounded p-1.5 border border-slate-100 text-[10px]">
                              "Toma las hojas. El frente impreso debe ponerse <strong>boca abajo (mirando al piso del cajón)</strong>, pero debes <strong>girar la hoja de arriba a abajo</strong> de modo que el borde superior (la cabecera) quede ahora apuntando <strong>hacia afuera (hacia ti)</strong>."
                            </p>
                          </div>
                        )
                      ) : (
                        flipDirection === 'long' ? (
                          <div className="space-y-1.5 text-slate-700">
                            <p className="font-medium">Para <strong>Alimentadores Superiores Traseros</strong> con <strong>Giro Libro</strong>:</p>
                            <p className="italic bg-slate-50 rounded p-1.5 border border-slate-100 text-[10px]">
                              "Inserta las hojas verticalmente en el alimentador de modo que el frente que ya imprimiste quede <strong>boca abajo (mirando hacia atrás, a la pared de soporte)</strong>, manteniendo la cabeza de la hoja apuntando <strong>hacia abajo</strong>."
                            </p>
                          </div>
                        ) : (
                          <div className="space-y-1.5 text-slate-700">
                            <p className="font-medium">Para <strong>Alimentadores Superiores Traseros</strong> con <strong>Giro Moneda</strong>:</p>
                            <p className="italic bg-slate-50 rounded p-1.5 border border-slate-100 text-[10px]">
                              "Inserta las hojas en el alimentador de modo que la parte ya impresa mire <strong>hacia adelante (mirándote a ti)</strong>, pero habiéndolas girado de arriba a abajo (el borde superior queda ahora apuntando <strong>hacia abajo</strong>)."
                            </p>
                          </div>
                        )
                      )}
                    </div>
                  </div>
                  <div className="text-[10px] text-blue-700 font-semibold bg-blue-100/40 p-1.5 text-center rounded border border-blue-100/50">
                    💡 ¡Truco extra! Reinserta siempre en el mismo orden (pág 1 queda arriba).
                  </div>
                </div>

                {/* Step 3 */}
                <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-xs flex flex-col justify-between space-y-4">
                  <div className="space-y-2.5">
                    <span className="bg-green-100 text-green-800 text-[11px] font-extrabold font-mono py-0.5 px-2 rounded-full block w-fit">
                      PASO 3
                    </span>
                    <h4 className="font-bold text-slate-800 text-xs uppercase tracking-wider flex items-center gap-1.5">
                      <span>3️⃣</span> Impresión de los Reversos
                    </h4>
                    <p className="text-slate-500 text-xs leading-relaxed">
                      Con las hojas ya devueltas con la rotación correcta a la impresora:
                    </p>
                    <div className="bg-slate-50 border border-slate-100 rounded-lg p-2.5 space-y-1 text-[11px]">
                      <span className="font-bold text-slate-800 block">En el diálogo de Impresión:</span>
                      <ul className="list-disc pl-4 space-y-1 text-slate-600 font-medium">
                        <li>Selecciona únicamente <strong>Páginas Pares</strong> (imprime la secuencia <code className="bg-white px-1 border rounded text-[10px] font-bold">2, 4, 6, 8...</code>).</li>
                        <li>Configura la misma escala <strong>100% (Real)</strong> anterior.</li>
                        <li>Haz clic en <strong>Imprimir</strong> de nuevo.</li>
                      </ul>
                    </div>
                  </div>
                  <div className="text-[10px] text-slate-500 italic">
                    ⭐ El motor se encargará de reordenar tus tarjetas espejándolas para que la pregunta n°1 coincida exactamente en la parte de atrás de la respuesta n°1.
                  </div>
                </div>

              </div>
            </div>

            {/* General hacks list */}
            <div className="bg-white border border-gray-200 rounded-xl p-5 text-left text-xs space-y-4">
              <h3 className="font-extrabold text-slate-800 text-xs uppercase tracking-wider pb-2 border-b border-gray-100 flex items-center gap-1.5">
                <span>🛠️</span> Hacks Prácticos para Terminar tus Fichas como Compradas:
              </h3>
              
              <ul className="grid grid-cols-1 md:grid-cols-2 gap-4 text-slate-600">
                <li className="flex items-start gap-2.5">
                  <span className="text-green-500 font-bold shrink-0">✓</span>
                  <div>
                    <strong className="text-slate-800 block mb-0.5">Gramaje del Papel recomendando:</strong>
                    Para que no traspase la luz al leer las preguntas, utiliza papel opalina blanca o ilustración mate de <b>240 gramos o más</b>. Puedes usar hojas delgadas estándar y pegarlas entre sí con plasticola en barra si no deseas comprar papel grueso.
                  </div>
                </li>
                <li className="flex items-start gap-2.5">
                  <span className="text-green-500 font-bold shrink-0">✓</span>
                  <div>
                    <strong className="text-slate-800 block mb-0.5">Corte seguro y recto:</strong>
                    Usa regla metálica pesada y un cutter con filo nuevo. Corta deslizando suavemente de 2 a 3 veces sobre las guías grises en vez de forzar un solo corte profundo. Previene desgarros.
                  </div>
                </li>
                <li className="flex items-start gap-2.5">
                  <span className="text-green-500 font-bold shrink-0">✓</span>
                  <div>
                    <strong className="text-slate-800 block mb-0.5">El toque PRO definitivo:</strong>
                    Activa la opción de esquinas redondeadas en el editor. Cuando recortes todas tus tarjetas cuadradas, utiliza una redondeadora / troqueladora manual de esquinas (se consiguen online por precio mínimo). Les da un tacto ultrasuave idéntico a una baraja de naipes comercial.
                  </div>
                </li>
                <li className="flex items-start gap-2.5">
                  <span className="text-green-500 font-bold shrink-0">✓</span>
                  <div>
                    <strong className="text-slate-800 block mb-0.5">Clasificación Automática:</strong>
                    La franja de colores superiores agrupa las fichas por categoría. No necesitas rotularlas a mano al dividirlas por materias en tu organizador físico.
                  </div>
                </li>
              </ul>
            </div>
          </div>
        )}
      </main>

      {/* Floating high-fidelity success/info notification toast */}
      {sessionToast && (
        <div className="no-print fixed bottom-6 right-6 z-[9999] flex items-center gap-3 bg-slate-900/95 backdrop-blur-md border border-slate-700/60 text-white py-3 px-4.5 rounded-xl shadow-2xl animate-fadeIn transition-all duration-300">
          <span className="text-base select-none">{sessionToast.type === 'success' ? '🚀' : 'ℹ️'}</span>
          <div className="text-left max-w-sm">
            <p className="text-[10px] font-extrabold uppercase tracking-wider text-blue-400">Notificación del Sistema</p>
            <p className="text-xs font-medium text-slate-100">{sessionToast.message}</p>
          </div>
          <button 
            onClick={() => setSessionToast(null)} 
            className="text-slate-400 hover:text-slate-200 ml-2 text-xs font-bold cursor-pointer flex items-center justify-center p-1 hover:bg-slate-800 rounded-lg"
          >
            ✕
          </button>
        </div>
      )}
    </div>
  );
}
