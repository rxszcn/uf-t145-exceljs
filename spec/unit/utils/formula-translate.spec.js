const {translateFormula, translateIndex} = verquire('utils/formula-translate');

describe('formula-translate', () => {
  describe('translateIndex', () => {
    // splice 语义：start 之前不动，删除区内为 null，其后按 插入-删除 净位移
    it('keeps indices before the splice point', () => {
      expect(translateIndex(2, 3, 2, 1)).to.equal(2);
      expect(translateIndex(1, 3, 0, 2)).to.equal(1);
    });
    it('returns null inside the deleted window', () => {
      expect(translateIndex(3, 3, 2, 0)).to.equal(null);
      expect(translateIndex(4, 3, 2, 1)).to.equal(null);
    });
    it('shifts indices after the splice point', () => {
      expect(translateIndex(5, 3, 0, 2)).to.equal(7); // pure insert
      expect(translateIndex(5, 3, 2, 0)).to.equal(3); // pure delete
      expect(translateIndex(5, 3, 2, 1)).to.equal(4); // replace (net -1)
    });
  });

  describe('translateFormula row orientation', () => {
    const insert = {orientation: 'row', sheetName: 'S1', start: 3, numDelete: 0, numInsert: 2};
    const remove = {orientation: 'row', sheetName: 'S1', start: 3, numDelete: 2, numInsert: 0};

    it('shifts relative references on row insert', () => {
      expect(translateFormula('A1+A3+A10', insert)).to.equal('A1+A5+A12');
    });
    it('expands ranges straddling the row insert point', () => {
      expect(translateFormula('SUM(A1:A10)', insert)).to.equal('SUM(A1:A12)');
    });
    it('keeps absolute column/row dollars but still moves the locked coordinate on structural change', () => {
      // $A 锁列：行插入不影响列；行号 1 在插入点之前，整体不变
      expect(translateFormula('$A$1+A1', insert)).to.equal('$A$1+A1');
      // 相对行号 3 落在插入点 -> +2；$1 锁定但坐标在插入点之前
      expect(translateFormula('A3+A$1', insert)).to.equal('A5+A$1');
    });
    it('turns a deleted single reference into #REF!', () => {
      expect(translateFormula('A3', remove)).to.equal('#REF!');
      expect(translateFormula('A1+A4', remove)).to.equal('A1+#REF!');
    });
    it('shrinks a range straddling the deleted rows', () => {
      expect(translateFormula('A1:A5', remove)).to.equal('A1:A3');
      expect(translateFormula('A3:A10', remove)).to.equal('A2:A8');
    });
    it('returns #REF! when the whole row range is deleted', () => {
      expect(translateFormula('SUM(A3:A4)', remove)).to.equal('SUM(#REF!)');
    });
    it('leaves references before the delete point untouched', () => {
      expect(translateFormula('A1:A2', remove)).to.equal('A1:A2');
    });
  });

  describe('translateFormula column orientation', () => {
    const insert = {orientation: 'col', sheetName: 'S1', start: 3, numDelete: 0, numInsert: 2};
    const remove = {orientation: 'col', sheetName: 'S1', start: 3, numDelete: 2, numInsert: 0};

    it('keeps columns before the insert point and shifts later ones', () => {
      expect(translateFormula('B10', insert)).to.equal('B10');
      expect(translateFormula('C10', insert)).to.equal('E10');
      expect(translateFormula('$C10', insert)).to.equal('$E10');
    });
    it('expands a column range', () => {
      expect(translateFormula('SUM(C1:C10)', insert)).to.equal('SUM(E1:E10)');
    });
    it('turns a deleted column reference into #REF!', () => {
      expect(translateFormula('C2', remove)).to.equal('#REF!');
      expect(translateFormula('$C$2', remove)).to.equal('#REF!');
    });
    it('shrinks a column range straddling the deleted columns', () => {
      expect(translateFormula('B2:E2', remove)).to.equal('B2:C2');
    });
    it('returns #REF! when the whole column range is deleted', () => {
      expect(translateFormula('C1:D1', remove)).to.equal('#REF!');
    });
  });

  describe('cross-sheet references', () => {
    const rowInsert = {orientation: 'row', sheetName: 'S1', start: 3, numDelete: 0, numInsert: 2};

    it('translates same-sheet prefixed references (bare and quoted)', () => {
      expect(translateFormula('S1!A10', rowInsert)).to.equal('S1!A12');
      expect(translateFormula('\'S1\'!A10', rowInsert)).to.equal('\'S1\'!A12');
    });
    it('translates quoted sheet names containing spaces', () => {
      expect(
        translateFormula('\'S 1\'!A10+A10', {...rowInsert, sheetName: 'S 1'})
      ).to.equal('\'S 1\'!A12+A12');
    });
    it('leaves references to other sheets untouched', () => {
      // 两个引用行号都在插入点（第3行）之前，且 Sheet2!A1 指向别的表
      expect(translateFormula('Sheet2!A1+A2', rowInsert)).to.equal('Sheet2!A1+A2');
      expect(translateFormula('Other!A10', rowInsert)).to.equal('Other!A10');
      // 跨表前缀保留，同表裸引用仍按规则平移
      expect(translateFormula('Other!A1+A10', rowInsert)).to.equal('Other!A1+A12');
    });
  });

  describe('literals, function names and pseudo references', () => {
    const rowInsert = {orientation: 'row', sheetName: 'S1', start: 3, numDelete: 0, numInsert: 2};

    it('does not touch addresses inside string literals', () => {
      expect(translateFormula('CONCATENATE("A10", A10)', rowInsert)).to.equal(
        'CONCATENATE("A10", A12)'
      );
    });
    it('does not mistake function names for references', () => {
      expect(translateFormula('LOG10(5)+A10', rowInsert)).to.equal('LOG10(5)+A12');
    });
    it('preserves error literals', () => {
      const remove = {orientation: 'row', sheetName: 'S1', start: 3, numDelete: 2, numInsert: 0};
      expect(translateFormula('#REF!+A1', remove)).to.equal('#REF!+A1');
      expect(translateFormula('IF(A5=0,"#N/A",A5)', remove)).to.equal(
        'IF(A3=0,"#N/A",A3)'
      );
    });
    it('does not touch names without an A1 shape', () => {
      expect(translateFormula('MYNAME+A1', rowInsert)).to.equal('MYNAME+A1');
    });
  });
});
