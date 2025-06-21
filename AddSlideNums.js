// メニューに追加する
function onOpen() {
  SlidesApp.getUi()
    .createMenu('ページ番号ツール')
    .addItem('ページ番号を振る', 'addPageNumbersWithCustomColor')
    .addItem('ページ番号を削除', 'removePageNumbers')
    .addToUi();
}

// 共通の色範囲定義
const colorRanges = [
  { start: 1, end: 30, color: '#FFFFFF' },
  { start: 31, end: 43, color: '#000000' },
  { start: 44, end: 74, color: '#FFFFFF' }
];

// ページ番号を振る（デフォルト）
function addPageNumbersWithCustomColor() {
  const presentation = SlidesApp.getActivePresentation();
  const pageCount = presentation.getSlides().length;
  const maxDigits = String(pageCount).length;
  const boxWidth = Math.max(20, maxDigits * 9 + 8);

  addPageNumbers({
    boxWidth: boxWidth,
    precise: false
  });
}

// 精密配置オプション（レイアウトが崩れる場合用）
function addPageNumbersWithPreciseAlignment() {
  const presentation = SlidesApp.getActivePresentation();
  const pageCount = presentation.getSlides().length;
  const maxDigits = String(pageCount).length;

  let boxWidth;
  if (maxDigits <= 1) boxWidth = 20;
  else if (maxDigits === 2) boxWidth = 28;
  else if (maxDigits === 3) boxWidth = 36;
  else boxWidth = maxDigits * 9 + 8;

  addPageNumbers({
    boxWidth: boxWidth,
    precise: true
  });
}

// 共通のページ番号挿入処理
function addPageNumbers(config) {
  const { boxWidth, precise } = config;

  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageWidth = presentation.getPageWidth();
  const pageHeight = presentation.getPageHeight();

  const fontSize = 12;
  const boxHeight = 18;
  const marginX = 6;
  const marginY = 12;

  const baseLeft = pageWidth - boxWidth - marginX;
  const baseTop = pageHeight - boxHeight - marginY;

  slides.forEach((slide, index) => {
    const pageNum = index + 1;
    const color = getColorForPage(pageNum);
    const textbox = slide.insertTextBox(String(pageNum));
    const textRange = textbox.getText();

    applyTextStyle(textRange, fontSize, color, precise);
    positionTextBox(textbox, boxWidth, boxHeight, baseLeft, baseTop);

    textbox.setTitle('PageNumberTag');
  });

  showAlert('ページ番号の挿入が完了しました！');
}

// テキストスタイル設定（共通）
function applyTextStyle(textRange, fontSize, color, precise) {
  const style = textRange.getTextStyle();
  style.setFontSize(fontSize).setBold(true).setForegroundColor(color);

  const paraStyle = textRange.getParagraphStyle();
  paraStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  if (precise) paraStyle.setLineSpacing(100);
}

// テキストボックスのサイズと位置
function positionTextBox(textbox, width, height, left, top) {
  textbox.setWidth(width);
  textbox.setHeight(height);
  textbox.setLeft(left);
  textbox.setTop(top);
}

// ページ番号を削除する
function removePageNumbers() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  slides.forEach(slide => {
    slide.getShapes().forEach(shape => {
      if (shape.getTitle() === 'PageNumberTag') {
        shape.remove();
      }
    });
  });
  showAlert('ページ番号をすべて削除しました。');
}

// 色取得ロジック
function getColorForPage(pageNum) {
  for (const range of colorRanges) {
    if (pageNum >= range.start && pageNum <= range.end) {
      return range.color;
    }
  }
  return '#FFFFFF';
}

// アラート表示（共通）
function showAlert(message) {
  SlidesApp.getUi().alert(message);
}
