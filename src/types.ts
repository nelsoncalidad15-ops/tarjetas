export interface Card {
  id: string;
  frontText: string;
  backText: string;
  frontImage?: string | null; // base64 data url or object URL
  backImage?: string | null;  // base64 data url or object URL
  useImageFront: boolean;
  useImageBack: boolean;
  copies: number; // number of copies to print
  category?: string;
  createdAt: number;
}

export interface CardStyle {
  textSize: 'xs' | 'sm' | 'base' | 'lg' | 'xl' | '2xl';
  fontFamily: 'sans' | 'serif' | 'mono';
  borderRadiusMm: number;
  textColor: string;
  backgroundColor: string;
  backTextColor: string;
  backBackgroundColor: string;
  alignment: 'left' | 'center' | 'right';
  borderStyle: 'none' | 'solid' | 'dashed' | 'dotted';
  borderColor: string;
}

export interface PrintConfig {
  cardWidthMm: number;
  cardHeightMm: number;
  paperSize: 'a4' | 'letter';
  showCutMarks: boolean;
  showBorder: boolean;
  style: CardStyle;
}
