import React, { useState } from 'react';
import { Card } from '../types';
import { Trash2, Copy, Image as ImageIcon, FileText, ChevronRight, Hash, Plus, Minus, RotateCw, Eye } from 'lucide-react';

interface CardItemProps {
  card: Card;
  index: number;
  onUpdate: (updatedCard: Card) => void;
  onDelete: (id: string) => void;
  onDuplicate: (card: Card) => void;
}

export const CardItem: React.FC<CardItemProps> = ({
  card,
  index,
  onUpdate,
  onDelete,
  onDuplicate,
}) => {
  const [isFlipped, setIsFlipped] = useState(false);

  const updateField = (field: keyof Card, value: any) => {
    onUpdate({
      ...card,
      [field]: value,
    });
  };

  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow-xs overflow-hidden flex flex-col hover:border-gray-300 transition-all duration-200">
      {/* Card Header with index, copies, and delete/duplicate controls */}
      <div className="px-4 py-2.5 bg-gray-50 border-b border-gray-100 flex items-center justify-between text-xs text-gray-500 font-medium">
        <div className="flex items-center gap-2">
          <span className="bg-gray-200 text-gray-700 px-2 py-0.5 rounded-full font-mono font-bold">
            #{index + 1}
          </span>
          {card.category && (
            <span className="bg-blue-50 text-blue-700 px-2 py-0.5 rounded-md border border-blue-100 max-w-[120px] truncate">
              {card.category}
            </span>
          )}
        </div>
        
        <div className="flex items-center gap-1.5">
          {/* Flip indicator pill */}
          <span className={`text-[10px] px-1.5 py-0.5 rounded-md font-semibold transition-all ${
            isFlipped 
              ? 'bg-purple-50 text-purple-700 border border-purple-100' 
              : 'bg-amber-50 text-amber-700 border border-amber-100'
          }`}>
            {isFlipped ? 'Modificando Reverso ↩' : 'Modificando Anverso ↪'}
          </span>

          {/* Duplicate Button */}
          <button
            onClick={() => onDuplicate(card)}
            className="p-1 text-gray-400 hover:text-gray-600 hover:bg-gray-100 rounded transition-colors cursor-pointer"
            title="Duplicar Tarjeta"
          >
            <Copy className="w-3.5 h-3.5" />
          </button>
          {/* Delete Button */}
          <button
            onClick={() => onDelete(card.id)}
            className="p-1 text-red-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors cursor-pointer"
            title="Eliminar Tarjeta"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Main content - Side-by-side tab-controllers to rotate with style */}
      <div className="p-4 flex-1 flex flex-col gap-4">
        
        {/* Visual Rotary Tab Selector (Synchronizes with Flip State!) */}
        <div className="flex bg-slate-100 p-0.5 rounded-lg text-xs font-semibold relative z-10">
          <button
            type="button"
            className={`flex-1 py-1 px-2 rounded-md transition-all flex items-center justify-center gap-1 cursor-pointer ${
              !isFlipped
                ? 'bg-white text-gray-900 shadow-xs font-bold'
                : 'text-gray-500 hover:text-gray-900'
            }`}
            onClick={() => setIsFlipped(false)}
          >
            <span>Anverso (Frente)</span>
          </button>
          <button
            type="button"
            className={`flex-1 py-1 px-2 rounded-md transition-all flex items-center justify-center gap-1 cursor-pointer ${
              isFlipped
                ? 'bg-white text-gray-900 shadow-xs font-bold'
                : 'text-gray-500 hover:text-gray-900'
            }`}
            onClick={() => setIsFlipped(true)}
          >
            <span>Reverso (Atrás)</span>
          </button>
        </div>

        {/* 3D ROTATION CANVAS AREA */}
        <div className="relative w-full h-[180px] perspective-1000">
          <div 
            className={`relative w-full h-full preserve-3d transition-transform duration-500 ease-out ${
              isFlipped ? 'rotate-y-180' : ''
            }`}
          >
            
            {/* FRONT SIDE (ANVERSO) */}
            <div className="absolute inset-0 w-full h-full backface-hidden bg-gradient-to-br from-slate-50 to-slate-100/50 border border-slate-200 rounded-xl p-3 flex flex-col justify-between shadow-xs">
              <div className="space-y-1.5 flex-1 flex flex-col overflow-hidden">
                {/* Image panel if available */}
                {card.frontImage && (
                  <div className="flex items-center justify-between gap-1.5 p-1.5 bg-white/80 rounded-lg border border-slate-200/60 text-[10px]">
                    <div className="flex items-center gap-1.5 overflow-hidden">
                      <img
                        src={card.frontImage}
                        alt="Front Screen"
                        className="w-7 h-7 object-cover rounded border border-gray-200 flex-shrink-0"
                      />
                      <div className="text-left truncate">
                        <p className="font-semibold text-gray-700 truncate">Captura frontal asociada</p>
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        updateField('useImageFront', !card.useImageFront);
                      }}
                      className={`px-1.5 py-0.5 rounded text-[9px] font-bold uppercase transition-colors cursor-pointer ${
                        card.useImageFront
                          ? 'bg-green-100 text-green-700 border border-green-200/50'
                          : 'bg-gray-100 text-gray-600 border border-gray-200/50'
                      }`}
                    >
                      {card.useImageFront ? 'Imagen' : 'Texto Clean'}
                    </button>
                  </div>
                )}

                <div className="flex-1 flex flex-col min-h-0">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-[9px] font-extrabold text-slate-400 uppercase tracking-wider">
                      PREGUNTA (ANVERSO)
                    </span>
                    <button
                      type="button"
                      onClick={() => setIsFlipped(true)}
                      className="text-[9px] text-blue-600 hover:text-blue-700 font-bold flex items-center gap-0.5 cursor-pointer hover:underline"
                      title="Girar para ver respuesta"
                    >
                      <RotateCw className="w-2.5 h-2.5" /> Voltear
                    </button>
                  </div>
                  <textarea
                    value={card.frontText}
                    onChange={(e) => updateField('frontText', e.target.value)}
                    placeholder="Escribe la pregunta o concepto principal aquí..."
                    className="flex-1 w-full text-xs p-2 bg-white border border-slate-200 rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 outline-none resize-none overflow-y-auto"
                    onClick={(e) => e.stopPropagation()} // Prevent flip when typing
                  />
                </div>
              </div>

              {/* Bottom hint of the card itself */}
              <div 
                onClick={() => setIsFlipped(true)}
                className="pt-1.5 border-t border-slate-100 flex items-center justify-between text-[9px] text-slate-400 cursor-pointer hover:text-blue-500 transition-colors"
              >
                <span>Materia: {card.category || 'Generales'}</span>
                <span className="flex items-center gap-0.5 font-bold text-blue-500">
                  Ver Reverso <ChevronRight className="w-3 h-3" />
                </span>
              </div>
            </div>

            {/* BACK SIDE (REVERSO) */}
            <div className="absolute inset-0 w-full h-full backface-hidden rotate-y-180 bg-gradient-to-br from-blue-50/20 to-blue-50/50 border border-blue-200/50 rounded-xl p-3 flex flex-col justify-between shadow-xs">
              <div className="space-y-1.5 flex-1 flex flex-col overflow-hidden">
                {/* Image panel if available */}
                {card.backImage && (
                  <div className="flex items-center justify-between gap-1.5 p-1.5 bg-white/90 rounded-lg border border-blue-100/80 text-[10px]">
                    <div className="flex items-center gap-1.5 overflow-hidden">
                      <img
                        src={card.backImage}
                        alt="Back Screen"
                        className="w-7 h-7 object-cover rounded border border-blue-100 flex-shrink-0"
                      />
                      <div className="text-left truncate">
                        <p className="font-semibold text-slate-700 truncate">Captura reverso asociada</p>
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        updateField('useImageBack', !card.useImageBack);
                      }}
                      className={`px-1.5 py-0.5 rounded text-[9px] font-bold uppercase transition-colors cursor-pointer ${
                        card.useImageBack
                          ? 'bg-blue-600 text-white'
                          : 'bg-gray-100 text-gray-600 border border-gray-200'
                      }`}
                    >
                      {card.useImageBack ? 'Imagen' : 'Texto Clean'}
                    </button>
                  </div>
                )}

                <div className="flex-1 flex flex-col min-h-0">
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-[9px] font-extrabold text-blue-500 uppercase tracking-wider">
                      RESPUESTA / EXPLICACIÓN (REVERSO)
                    </span>
                    <button
                      type="button"
                      onClick={() => setIsFlipped(false)}
                      className="text-[9px] text-slate-500 hover:text-slate-700 font-bold flex items-center gap-0.5 cursor-pointer hover:underline"
                      title="Girar para ver pregunta"
                    >
                      <RotateCw className="w-2.5 h-2.5" /> Voltear
                    </button>
                  </div>
                  <textarea
                    value={card.backText}
                    onChange={(e) => updateField('backText', e.target.value)}
                    placeholder="Escribe la respuesta detallada, datos de memoria o explicación aquí..."
                    className="flex-1 w-full text-xs p-2 bg-white border border-blue-100 rounded-lg focus:border-blue-500 focus:ring-1 focus:ring-blue-500/20 outline-none resize-none overflow-y-auto font-medium text-slate-700"
                    onClick={(e) => e.stopPropagation()} // Prevent flip when typing
                  />
                </div>
              </div>

              {/* Bottom hint of the card itself */}
              <div 
                onClick={() => setIsFlipped(false)}
                className="pt-1.5 border-t border-blue-100/50 flex items-center justify-between text-[9px] text-blue-400 cursor-pointer hover:text-slate-600 transition-colors"
              >
                <span>Examen listo para impresión</span>
                <span className="flex items-center gap-0.5 font-bold text-slate-600">
                  Ver Anverso <ChevronRight className="w-3 h-3" />
                </span>
              </div>
            </div>

          </div>
        </div>

        {/* Inputs for Category and Print Copies */}
        <div className="pt-1.5 border-t border-gray-150 grid grid-cols-2 gap-2.5 mt-auto">
          <div>
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider block mb-1">
              CATEGORÍA / MATERIA
            </label>
            <input
              type="text"
              value={card.category || ''}
              onChange={(e) => updateField('category', e.target.value)}
              placeholder="ISO 14001, Historia..."
              className="w-full text-xs py-1.5 px-2 bg-slate-50 hover:bg-slate-100/50 focus:bg-white border border-gray-250 rounded-lg focus:border-blue-500 outline-none transition-all text-slate-700 font-medium"
            />
          </div>
          <div>
            <label className="text-[10px] font-bold text-gray-400 uppercase tracking-wider block mb-1">
              COPIAS A IMPRIMIR
            </label>
            <div className="flex items-center border border-gray-250 rounded-lg overflow-hidden h-[30px] bg-slate-50">
              <button
                type="button"
                onClick={() => updateField('copies', Math.max(1, card.copies - 1))}
                className="w-8 h-full flex items-center justify-center bg-transparent text-gray-550 hover:bg-gray-150 border-r border-gray-200 transition-colors cursor-pointer"
              >
                <Minus className="w-3" />
              </button>
              <input
                type="number"
                min="1"
                value={card.copies}
                onChange={(e) => updateField('copies', Math.max(1, parseInt(e.target.value) || 1))}
                className="flex-1 h-full text-center text-xs font-bold text-gray-800 bg-white focus:outline-none [appearance:textfield] [&::-webkit-outer-spin-button]:appearance-none [&::-webkit-inner-spin-button]:appearance-none"
              />
              <button
                type="button"
                onClick={() => updateField('copies', card.copies + 1)}
                className="w-8 h-full flex items-center justify-center bg-transparent text-gray-550 hover:bg-gray-150 border-l border-gray-200 transition-colors cursor-pointer"
                title="Aumentar copias"
              >
                <Plus className="w-3" />
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

