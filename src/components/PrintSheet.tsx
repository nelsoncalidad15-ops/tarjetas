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
  const { cardWidthMm, cardHeightMm, paperSize, showCutMarks, showCardNumbers, style } = config;

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

  const startIndex = pageIndex * cardsPerPage;

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

  const renderCell = (card: Card | null, colIdx: number, rowIdx: number) => {
    if (!card) {
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
    const cardBgColor = isFront ? style.backgroundColor : style.backBackgroundColor;
    const cardTextColor = isFront ? style.textColor : style.backTextColor;
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
        {showCardNumbers && typeof card.studyNumber === 'number' && (
          <div className="absolute top-[2.2mm] right-[2.2mm] rounded-full border border-slate-300/70 bg-white/90 px-[2.2mm] py-[0.7mm] text-[8px] font-bold leading-none shadow-sm">
            #{card.studyNumber}
          </div>
        )}

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
            {card.category && (
              <div
                className="text-[8px] uppercase tracking-wider font-semibold opacity-60 mb-1.5"
                style={{ color: cardTextColor }}
              >
                {card.category}
              </div>
            )}
            <div className={`${textClass} ${fontClass} ${alignClass} font-medium leading-relaxed overflow-hidden break-words whitespace-pre-line`}>
              {cardText}
            </div>
          </div>
        )}

        {showCutMarks && (
          <div className="absolute inset-0 border border-solid border-slate-300 opacity-20 pointer-events-none" />
        )}
      </div>
    );
  };

  const gridCells = [];
  for (let r = 0; r < rows; r++) {
    const rowCells = [];
    for (let c = 0; c < cols; c++) {
      const targetCol = side === 'front' ? c : (cols - 1 - c);
      const cardIndex = r * cols + targetCol;
      const card = (startIndex + cardIndex) < cards.length ? cards[startIndex + cardIndex] : null;
      rowCells.push({ card, colIdx: c, rowIdx: r });
    }
    gridCells.push(rowCells);
  }

  const pageLabel = side === 'front'
        ? `Pagina ${pageIndex + 1} de ${totalPages} (Hojas de Frente - Anverso)`
    : `Pagina ${pageIndex + 1} de ${totalPages} (Hojas de Reverso - Dorso Espejado)`;

  return (
    <div className="flex flex-col items-center">
      <div className="no-print mt-6 mb-2 text-xs text-slate-500 flex items-center gap-2 justify-between w-[210mm] max-w-full bg-slate-100 py-1.5 px-3 rounded-lg border border-slate-200">
        <span className="font-semibold text-slate-700 uppercase tracking-wide"> [Print] {pageLabel}
        </span>
        <span className="bg-slate-300 font-mono font-bold text-slate-800 px-2 py-0.5 rounded text-[10px]">
          {cols * rows} tarjetas ({cols}x{rows})
        </span>
      </div>

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
        <div
          className="grid gap-0"
          style={{
            gridTemplateColumns: `repeat(${cols}, ${cardWidthMm}mm)`,
            gridTemplateRows: `repeat(${rows}, ${cardHeightMm}mm)`,
            width: `${gridWidth}mm`,
            height: `${gridHeight}mm`,
          }}
        >
          {gridCells.map((rowCells) =>
            rowCells.map(({ card, colIdx, rowIdx }) => renderCell(card, colIdx, rowIdx))
          )}
        </div>
      </div>
    </div>
  );
};
