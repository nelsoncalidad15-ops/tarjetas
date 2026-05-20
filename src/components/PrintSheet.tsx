import React from 'react';
import { Card, PrintConfig } from '../types';

interface PrintSheetProps {
  cards: Card[]; // Flattened lists of cards based on copy counters
  config: PrintConfig;
  side: 'front' | 'back';
  pageIndex: number;
  totalPages: number;
}

export const PrintSheet: React.FC<PrintSheetProps> = ({
  cards,
  config,
  side,
  pageIndex,
  totalPages,
}) => {
  const { cardWidthMm, cardHeightMm, paperSize, showCutMarks, style } = config;

  // Paper Dimensions in mm
  const pageWidth = paperSize === 'a4' ? 210 : 215.9;
  const pageHeight = paperSize === 'a4' ? 297 : 279.4;

  // Calculate grid layout ensuring safe margins
  const marginMm = 8; // standard safe printer margin
  const cols = Math.max(1, Math.floor((pageWidth - marginMm * 2) / cardWidthMm));
  const rows = Math.max(1, Math.floor((pageHeight - marginMm * 2) / cardHeightMm));
  const cardsPerPage = cols * rows;

  // Compute exact padding to center the grid perfectly
  const gridWidth = cols * cardWidthMm;
  const gridHeight = rows * cardHeightMm;
  const paddingLeftRight = (pageWidth - gridWidth) / 2;
  const paddingTopBottom = (pageHeight - gridHeight) / 2;

  // Get current page cards (fill empty spaces with nulls to complete grid structure)
  const startIndex = pageIndex * cardsPerPage;
  const pageCards: (Card | null)[] = [];
  for (let i = 0; i < cardsPerPage; i++) {
    const cardIndex = startIndex + i;
    if (cardIndex < cards.length) {
      pageCards.push(cards[cardIndex]);
    } else {
      pageCards.push(null);
    }
  }

  // Choose appropriate styles based on config
  const textClass = 
    style.textSize === 'xs' ? 'text-[10px]' :
    style.textSize === 'sm' ? 'text-xs' :
    style.textSize === 'base' ? 'text-sm' :
    style.textSize === 'lg' ? 'text-base' :
    style.textSize === 'xl' ? 'text-lg' : 'text-xl';

  const fontClass = 
    style.fontFamily === 'mono' ? 'font-mono' :
    style.fontFamily === 'serif' ? 'font-serif' : 'font-sans';

  const alignClass = 
    style.alignment === 'left' ? 'text-left' :
    style.alignment === 'right' ? 'text-right' : 'text-center';

  // Helper to render individual card cell
  const renderCell = (card: Card | null, colIdx: number, rowIdx: number) => {
    if (!card) {
      // Empty placeholder cell for padding.
      // We still render a blank cell with cut lines if requested so the page remains a consistent grid.
      return (
        <div
          key={`empty-${rowIdx}-${colIdx}`}
          className="relative box-border"
          style={{
            width: `${cardWidthMm}mm`,
            height: `${cardHeightMm}mm`,
          }}
        >
          {showCutMarks && (
            <div className="absolute inset-0 border border-dashed border-gray-200 pointer-events-none" />
          )}
        </div>
      );
    }

    const isFront = side === 'front';
    const useImage = isFront ? card.useImageFront : card.useImageBack;
    const imageUrl = isFront ? card.frontImage : card.backImage;
    const cardText = isFront ? card.frontText : card.backText;

    // Define card design classes
    const cardBgColor = isFront ? style.backgroundColor : style.backBackgroundColor;
    const cardTextColor = isFront ? style.textColor : style.backTextColor;

    // Borders
    const borderStyleValue = style.borderStyle;
    const borderColorValue = style.borderColor;

    return (
      <div
        key={`card-${card.id}-${rowIdx}-${colIdx}`}
        id={`card-print-${card.id}-${side}`}
        className="relative overflow-hidden box-border flex flex-col justify-between"
        style={{
          width: `${cardWidthMm}mm`,
          height: `${cardHeightMm}mm`,
          backgroundColor: cardBgColor,
          color: cardTextColor,
          borderRadius: `${style.borderRadiusMm}mm`,
          borderWidth: borderStyleValue !== 'none' ? '1.5px' : '0px',
          borderStyle: borderStyleValue !== 'none' ? borderStyleValue : 'solid',
          borderColor: borderColorValue,
          padding: '4.5mm',
        }}
      >
        {/* Background / Content rendered depending on option */}
        {useImage && imageUrl ? (
          <div className="absolute inset-0 w-full h-full flex items-center justify-center p-0">
            <img
              src={imageUrl}
              alt="Card Snapshot"
              className="w-full h-full object-cover"
              referrerPolicy="no-referrer"
            />
          </div>
        ) : (
          <div className="w-full h-full flex flex-col justify-center overflow-hidden">
            {/* Tag/Category in small text */}
            {card.category && (
              <div 
                className="text-[8px] uppercase tracking-wider font-semibold opacity-60 mb-1.5"
                style={{ color: cardTextColor }}
              >
                {card.category}
              </div>
            )}
            {/* Main Text Markdown or plain text */}
            <div className={`${textClass} ${fontClass} ${alignClass} font-medium leading-relaxed overflow-hidden break-words whitespace-pre-line`}>
              {cardText}
            </div>
          </div>
        )}

        {/* Optional corner cut guides drawing */}
        {showCutMarks && (
          <div className="absolute inset-0 border border-solid border-slate-300 opacity-20 pointer-events-none" />
        )}
      </div>
    );
  };

  // Convert flat index array of 1 page cells into a grid representation
  const gridCells = [];
  for (let r = 0; r < rows; r++) {
    const rowCells = [];
    for (let c = 0; c < cols; c++) {
      // Calculate index for the cell based on layout side
      // For frontal pages, normal left-to-right grid order index: r * cols + c
      // For backend pages (mirroring!), right-to-left order on each row: r * cols + (cols - 1 - c)
      const targetCol = side === 'front' ? c : (cols - 1 - c);
      const cardIndex = r * cols + targetCol;
      const card = (startIndex + cardIndex) < cards.length ? cards[startIndex + cardIndex] : null;
      rowCells.push({ card, colIdx: c, rowIdx: r });
    }
    gridCells.push(rowCells);
  }

  // Calculate printable page header notes
  const pageLabel = side === 'front' 
    ? `Página ${pageIndex + 1} de ${totalPages} (Hojas de Frente - Anverso)`
    : `Página ${pageIndex + 1} de ${totalPages} (Hojas de Reverso - Dorso Espejado)`;

  return (
    <div className="flex flex-col items-center">
      {/* Screen-only Page Header, hidden in standard print output */}
      <div className="no-print mt-6 mb-2 text-xs text-slate-500 flex items-center gap-2 justify-between w-[210mm] max-w-full bg-slate-100 py-1.5 px-3 rounded-lg border border-slate-200">
        <span className="font-semibold text-slate-700 uppercase tracking-wide">
          🖨️ {pageLabel}
        </span>
        <span className="bg-slate-300 font-mono font-bold text-slate-800 px-2 py-0.5 rounded text-[10px]">
          {cols * rows} tarjetas ({cols}x{rows})
        </span>
      </div>

      {/* Main Printed Page Area */}
      <div
        className="print-page bg-white shadow-xl border border-gray-200 relative overflow-hidden flex flex-col select-none"
        style={{
          width: `${pageWidth}mm`,
          height: `${pageHeight}mm`,
          paddingLeft: `${paddingLeftRight}mm`,
          paddingRight: `${paddingLeftRight}mm`,
          paddingTop: `${paddingTopBottom}mm`,
          paddingBottom: `${paddingTopBottom}mm`,
          boxSizing: 'border-box',
        }}
      >
        {/* Dynamic CSS Grid based on exact counts */}
        <div 
          className="grid gap-0"
          style={{
            gridTemplateColumns: `repeat(${cols}, ${cardWidthMm}mm)`,
            gridTemplateRows: `repeat(${rows}, ${cardHeightMm}mm)`,
            width: `${gridWidth}mm`,
            height: `${gridHeight}mm`,
          }}
        >
          {gridCells.map((rowCells, rIdx) => 
            rowCells.map(({ card, colIdx, rowIdx }) => renderCell(card, colIdx, rowIdx))
          )}
        </div>
      </div>
    </div>
  );
};
