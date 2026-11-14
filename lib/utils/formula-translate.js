const colCache = require('./col-cache');

// 结构变更（插行/删行、插列/删列）后平移公式中的 A1 引用。
//
// 设计取舍（交付说明中亦有记载）：
// 1. 结构性插删对绝对引用（$）与相对引用一视同仁——$ 锁定的是复制填充时的
//    平移行为，而工作表几何形状改变时，Excel 自身同样会改写 $A$1 这类引用。
// 2. 落入被删区间的单点引用按 Excel 语义改写为 #REF!；区间端点落入删除区时
//    夹取到删除点外侧（区间收缩），仅当整个区间都被删除才整体变为 #REF!。
// 3. 跨表引用只在其归属表与发生结构变更的表同名时平移，其余原样保留。
// 4. 字符串字面量、数值、错误字面量、函数名与结构化引用中的伪 A1 文本一律
//    不参与匹配（见 translateFormula 的分段扫描与边界守卫）。

// 可选的表前缀（支持 'Sheet 1'! 与 Sheet1! 两种形态），后跟单元格或 A1:B2
// 区间，$ 符号各自独立捕获。表前缀必须以 ! 收尾，函数名因不含行号、且后面
// 不是 ! 而不会命中。
const refTokenRx =
  /(?:(?:'((?:[^']|'')*)'|([A-Za-z0-9_.\- ]+))!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?/g;

// 字符串字面量（双引号闭合，"" 为转义）以及错误字面量（#REF!、#N/A 等）。
// 注意 '#' 必须在非捕获组内部，否则 N/A、REF! 等会脱离 '#' 被单独匹配。
const skippableRx =
  /"(?:[^"]|"")*"|#(?:REF!|N\/A|NAME\?|VALUE!|DIV\/0!|NULL!|NUM!|SPILL!|CALC!|GETTING_DATA)/g;

function isWordChar(ch) {
  return /[A-Za-z0-9_]/.test(ch);
}

// 收集全局正则的全部匹配（避免在 while 条件内赋值）
function collectMatches(rx, text) {
  const matches = [];
  rx.lastIndex = 0;
  let match = rx.exec(text);
  while (match !== null) {
    matches.push(match);
    match = rx.exec(text);
  }
  return matches;
}

// 按 splice 语义把一个 1 维坐标平移：
//   start        插入/删除位置（1 基）
//   numDelete    删除的行数/列数
//   numInsert    插入的行数/列数
// 返回 null 表示该坐标落在被删区间内（对应 Excel 的 #REF!）
function translateIndex(index, start, numDelete, numInsert) {
  if (index < start) {
    return index;
  }
  if (index < start + numDelete) {
    return null;
  }
  return index + numInsert - numDelete;
}

function sameSheetName(a, b) {
  return a !== undefined && b !== undefined && a.toLowerCase() === b.toLowerCase();
}

function buildSheetPrefix(refSheet, quotedSheet, bareSheet) {
  if (refSheet === undefined) {
    return '';
  }
  if (quotedSheet !== undefined) {
    return `'${quotedSheet}'!`;
  }
  return `${bareSheet}!`;
}

// 重写一个引用 token。orientation: 'row' | 'col'
// crossSheetOnly 为 true 时只平移显式带 sheetName 前缀的跨表引用
function translateToken(
  token,
  orientation,
  sheetName,
  start,
  numDelete,
  numInsert,
  crossSheetOnly
) {
  const {
    quotedSheet,
    bareSheet,
    col1Dollar,
    col1,
    row1Dollar,
    row1,
    col2Dollar,
    col2,
    row2Dollar,
    row2,
    hasRange,
    index,
  } = token;

  const refSheet = quotedSheet !== undefined ? quotedSheet : bareSheet;

  // 跨表模式：只平移显式带目标表前缀的引用，裸 A1 引用属于公式所在表自身，
  // 与发生结构变更的目标表无关，原样保留
  if (crossSheetOnly && refSheet === undefined) {
    return null;
  }

  // 带表前缀且指向别的表：结构变更与该引用无关，原样保留
  if (refSheet !== undefined && !sameSheetName(refSheet, sheetName)) {
    return null;
  }

  const sheetPrefix = buildSheetPrefix(refSheet, quotedSheet, bareSheet);

  const buildAddress = (colDollar, col, rowDollar, row) =>
    `${colDollar || ''}${colCache.n2l(col)}${rowDollar || ''}${row}`;

  const firstCol = colCache.l2n(col1.toUpperCase());
  const firstRow = parseInt(row1, 10);

  if (!hasRange) {
    // 单点引用：落入删除区 -> #REF!
    const movedCol =
      orientation === 'col' ? translateIndex(firstCol, start, numDelete, numInsert) : firstCol;
    const movedRow =
      orientation === 'row' ? translateIndex(firstRow, start, numDelete, numInsert) : firstRow;
    if (movedRow === null || movedCol === null) {
      return {text: `${sheetPrefix}#REF!`, index};
    }
    return {
      text: `${sheetPrefix}${buildAddress(col1Dollar, movedCol, row1Dollar, movedRow)}`,
      index,
    };
  }

  // 区间引用
  const lastCol = colCache.l2n(col2.toUpperCase());
  const lastRow = parseInt(row2, 10);

  if (orientation === 'row') {
    // 整个区间（行向上）都被删除 -> #REF!
    if (firstRow >= start && lastRow < start + numDelete) {
      return {text: `${sheetPrefix}#REF!`, index};
    }
    const topIndex = translateIndex(firstRow, start, numDelete, numInsert);
    const bottomIndex = translateIndex(lastRow, start, numDelete, numInsert);
    const top = topIndex === null ? start - 1 : topIndex;
    const bottom = bottomIndex === null ? start - 1 : bottomIndex;
    return {
      text: `${sheetPrefix}${buildAddress(col1Dollar, firstCol, row1Dollar, top)}:${buildAddress(
        col2Dollar,
        lastCol,
        row2Dollar,
        bottom
      )}`,
      index,
    };
  }

  if (firstCol >= start && lastCol < start + numDelete) {
    return {text: `${sheetPrefix}#REF!`, index};
  }
  const leftIndex = translateIndex(firstCol, start, numDelete, numInsert);
  const rightIndex = translateIndex(lastCol, start, numDelete, numInsert);
  const left = leftIndex === null ? start - 1 : leftIndex;
  const right = rightIndex === null ? start - 1 : rightIndex;
  return {
    text: `${sheetPrefix}${buildAddress(col1Dollar, left, row1Dollar, firstRow)}:${buildAddress(
      col2Dollar,
      right,
      row2Dollar,
      lastRow
    )}`,
    index,
  };
}

// 平移公式字符串里的全部 A1 引用。
// orientation: 'row'（spliceRows）或 'col'（spliceColumns）
// crossSheetOnly: 仅平移显式指向 sheetName 的跨表引用，裸引用保持不变
function translateFormula(
  formula,
  {orientation, sheetName, start, numDelete, numInsert, crossSheetOnly = false}
) {
  if (typeof formula !== 'string' || !formula) {
    return formula;
  }

  // 先收集需要跳过的区间（字符串字面量、错误字面量），避免误伤其中的伪引用
  const skips = collectMatches(skippableRx, formula).map(skipMatch => ({
    start: skipMatch.index,
    end: skipMatch.index + skipMatch[0].length,
  }));
  const isSkipped = idx => skips.some(skip => idx >= skip.start && idx < skip.end);

  const replacements = [];
  collectMatches(refTokenRx, formula).forEach(match => {
    const idx = match.index;
    if (isSkipped(idx)) {
      return;
    }
    // 列字母在原文中的位置：匹配串内，表前缀之后、起始 $ 之后即列字母
    const prefixLength = match[0].indexOf(match[4]);
    const colLetterIdx = idx + prefixLength;
    // 左边界守卫：向前跳过空白，若紧邻标识符组成部分，则当前 token 是更长
    // 名字的尾部（如函数名 LOG10 尾部的 G10、定义名称 X1、结构化引用
    // Table1 等），不是独立引用，跳过。空白只在运算符（&、+ 等）之后出现，
    // 而标识符内部不允许空白，因此跳过空白后检查即可同时覆盖两种写法。
    let probe = colLetterIdx - 1;
    while (probe >= 0 && /\s/.test(formula[probe])) {
      probe--;
    }
    if (probe >= 0 && isWordChar(formula[probe])) {
      return;
    }
    // 右边界守卫：
    //  - 引用后紧跟字母/数字/下划线时，属于更长标识符，跳过；
    //  - 跳过空白后紧跟 '(' 时，该 token 实为函数名（如 LOG10、MYFUNC2 的
    //    尾部 UNC2），真实单元格引用不可能作为被调用函数。
    // 区间续接 ':'、运算符、逗号、括号闭合等均为合法边界。
    const endIdx = idx + match[0].length;
    if (endIdx < formula.length && isWordChar(formula[endIdx])) {
      return;
    }
    let lookAhead = endIdx;
    while (lookAhead < formula.length && /\s/.test(formula[lookAhead])) {
      lookAhead++;
    }
    if (formula[lookAhead] === '(') {
      return;
    }

    const token = {
      quotedSheet: match[1],
      bareSheet: match[2],
      col1Dollar: match[3] || '',
      col1: match[4],
      row1Dollar: match[5] || '',
      row1: match[6],
      hasRange: match[7] !== undefined,
      col2Dollar: match[7] || '',
      col2: match[8],
      row2Dollar: match[9] || '',
      row2: match[10],
      index: idx,
    };

    const replacement = translateToken(
      token,
      orientation,
      sheetName,
      start,
      numDelete,
      numInsert,
      crossSheetOnly
    );
    if (replacement) {
      replacements.push({start: idx, end: endIdx, text: replacement.text});
    }
  });

  if (!replacements.length) {
    return formula;
  }

  let result = '';
  let cursor = 0;
  replacements.forEach(rep => {
    result += formula.slice(cursor, rep.start) + rep.text;
    cursor = rep.end;
  });
  result += formula.slice(cursor);
  return result;
}

module.exports = {
  translateFormula,
  translateIndex,
};
