function onInstall() {
  onOpen();
}

function onOpen() {
  DocumentApp.getUi()
      .createMenu('CSH (beta)')
      .addItem('Convert Structured HTML', 'convertAndShowHtml')
      .addToUi();
}

/**
 * ドキュメントのコンテンツをHTMLに変換し、ダイアログに表示します。
 */
function convertAndShowHtml() {
  const body = DocumentApp.getActiveDocument().getBody();
  let htmlOutput = [];
  
  // リスト（<ul> or <ol>）の状態を管理する変数
  let inList = false;
  let currentListTag = null; // 'ul' または 'ol'

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const elementType = child.getType();

    if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = child.asListItem();
      const requiredListTag = isOrderedList(listItem.getGlyphType()) ? 'ol' : 'ul';

      // 1. リストがまだ始まっていない場合
      if (!inList) {
        htmlOutput.push(`<${requiredListTag}>`);
        inList = true;
        currentListTag = requiredListTag;
      }
      // 2. リストの種類が切り替わった場合 (例: ul -> ol)
      else if (currentListTag !== requiredListTag) {
        htmlOutput.push(`</${currentListTag}>`); // 前のリストを閉じる
        htmlOutput.push(`<${requiredListTag}>`); // 新しいリストを開始
        currentListTag = requiredListTag;
      }
      
      // リスト項目を処理
      htmlOutput.push(processElement(child));

    } else {
      // リスト項目以外の要素が来た場合、リストが開いていれば閉じる
      if (inList) {
        htmlOutput.push(`</${currentListTag}>`);
        inList = false;
        currentListTag = null;
      }
      // 通常の要素を処理
      const processedHtml = processElement(child);
      if (processedHtml) { // 空文字でない場合のみ配列に追加
        htmlOutput.push(processedHtml);
      }
    }
  }
  
  // ドキュメントの末尾がリストだった場合に備えてリストを閉じる
  if (inList) {
    htmlOutput.push(`</${currentListTag}>`);
  }

  // 結果をダイアログで表示
  const htmlResult = htmlOutput.join('\n');
  const htmlDialogContent = HtmlService.createHtmlOutput(`<pre>${escapeHtml(htmlResult)}</pre>`)
      .setWidth(700)
      .setHeight(500);
  DocumentApp.getUi().showModalDialog(htmlDialogContent, '生成されたHTML');
}

/**
 * 個々のドキュメント要素を対応するHTML文字列に変換します。
 * ★リスト項目の処理がシンプル化されています。
 * @param {GoogleAppsScript.Document.Element} element 処理対象の要素
 * @return {string} 変換後のHTML文字列
 */
function processElement(element) {
  const elementType = element.getType();

  switch (elementType) {
    case DocumentApp.ElementType.PARAGRAPH:
      const paragraph = element.asParagraph();
      if (paragraph.getNumChildren() === 0) return '';
      
      const heading = paragraph.getHeading();
      const textContent = processText(paragraph.asText());
      
      if(textContent.length == 0) {
        break
      }

      switch (heading) {
        case DocumentApp.ParagraphHeading.HEADING1: return `<h1>${textContent}</h1>`;
        case DocumentApp.ParagraphHeading.HEADING2: return `<h2>${textContent}</h2>`;
        case DocumentApp.ParagraphHeading.HEADING3: return `<h3>${textContent}</h3>`;
        case DocumentApp.ParagraphHeading.HEADING4: return `<h4>${textContent}</h4>`;
        case DocumentApp.ParagraphHeading.HEADING5: return `<h5>${textContent}</h5>`;
        case DocumentApp.ParagraphHeading.HEADING6: return `<h6>${textContent}</h6>`;
        default: return `<p>${textContent}</p>`;
      }

    case DocumentApp.ElementType.LIST_ITEM:
      // liタグの中身を生成することに専念
      return `  <li>${processText(element.asListItem().asText())}</li>`;
      
    case DocumentApp.ElementType.TABLE:
      const table = element.asTable();
      let tableHtml = ['<table>'];
      for (let i = 0; i < table.getNumRows(); i++) {
        const row = table.getRow(i);
        tableHtml.push('  <tr>');
        for (let j = 0; j < row.getNumCells(); j++) {
          const cell = row.getCell(j);
          const cellContent = [];
          for (let k = 0; k < cell.getNumChildren(); k++) {
              // 注意: セル内の要素の処理はメインループのロジックを再利用せず、
              // 単純なprocessElement呼び出しにしています。
              // セル内にリストがある場合の挙動は限定的です。
              cellContent.push(processElement(cell.getChild(k)));
          }
          tableHtml.push(`    <td>${cellContent.join('')}</td>`);
        }
        tableHtml.push('  </tr>');
      }
      tableHtml.push('</table>');
      return tableHtml.join('\n');

    case DocumentApp.ElementType.HORIZONTAL_RULE:
      return '<hr>';
      
    default:
      return ``;
  }
}

/**
 * Text要素内の書式（太字, 斜体, 下線）を<b>, <i>, <u>タグに変換します。
 * @param {GoogleAppsScript.Document.Text} textElement 処理対象のText要素
 * @return {string} 変換後のHTML文字列
 */
function processText(textElement) {
  if (!textElement || textElement.getText().trim() === '') return '&nbsp;'; // 空のpタグなどのために空白を返す

  let html = '';
  const text = textElement.getText();
  const indices = textElement.getTextAttributeIndices();

  for (let i = 0; i < indices.length; i++) {
    const startIndex = indices[i];
    const endIndex = (i + 1 < indices.length) ? indices[i + 1] : text.length;
    let part = text.substring(startIndex, endIndex);
    
    // リンク以外はエスケープを先に行う
    let escapedPart = escapeHtml(part);
    
    const attributes = textElement.getAttributes(startIndex);

    if (attributes[DocumentApp.Attribute.UNDERLINE]) {
      escapedPart = `<u>${escapedPart}</u>`;
    }
    if (attributes[DocumentApp.Attribute.ITALIC]) {
      escapedPart = `<i>${escapedPart}</i>`;
    }
    if (attributes[DocumentApp.Attribute.BOLD]) {
      escapedPart = `<b>${escapedPart}</b>`;
    }
    
    const linkUrl = textElement.getLinkUrl(startIndex);
    if (linkUrl) {
        escapedPart = `<a href="${linkUrl}">${escapedPart}</a>`;
    }
    
    html += escapedPart;
  }
  return html;
}

/**
 * リストが順序付きリスト(ol)かどうかを判定します。
 * @param {GoogleAppsScript.Document.GlyphType} glyphType 記号の種類
 * @return {boolean} 順序付きリストならtrue
 */
function isOrderedList(glyphType) {
    const orderedTypes = [
        DocumentApp.GlyphType.NUMBER,
        DocumentApp.GlyphType.LATIN_LOWER,
        DocumentApp.GlyphType.LATIN_UPPER,
        DocumentApp.GlyphType.ROMAN_LOWER,
        DocumentApp.GlyphType.ROMAN_UPPER
    ];
    return orderedTypes.indexOf(glyphType) !== -1;
}

/**
 * HTML特殊文字をエスケープします。
 * @param {string} text エスケープする文字列
 * @return {string} エスケープ後の文字列
 */
function escapeHtml(text) {
    return text
         .replace(/&/g, "&amp;")
         .replace(/</g, "&lt;")
         .replace(/>/g, "&gt;")
         .replace(/"/g, "&quot;")
         .replace(/'/g, "&#039;");
}
