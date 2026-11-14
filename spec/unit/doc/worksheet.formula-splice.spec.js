const Excel = verquire('exceljs');

// 公式引用平移的插入/删除形态矩阵：行与列两个方向、插入与删除四种形态，
// 外加跨表引用、绝对引用、字符串字面量、共享公式与既有值/样式/名称不回退。
describe('Worksheet', () => {
  describe('Splice formula reference translation', () => {
    it('row insert shifts formula references and expands ranges', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = 1;
      ws.getCell('A5').value = {formula: 'SUM(A1:A3)+A10', result: 0};

      // 在第 3 行前插入 1 行：>=3 的行坐标 +1，跨插入点区间扩张
      ws.spliceRows(3, 0, ['inserted']);

      expect(ws.getCell('A6').formula).to.equal('SUM(A1:A4)+A11');
    });

    it('row delete shifts, shrinks and invalidates references', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A10').value = {formula: 'A1+A3+A5+SUM(A3:A4)', result: 0};

      // 删除第 3、4 行
      ws.spliceRows(3, 2);

      // A3 落删除区 -> #REF!；A5 左移到 A3；整段 A3:A4 被删 -> #REF!
      expect(ws.getCell('A8').formula).to.equal('A1+#REF!+A3+SUM(#REF!)');
    });

    it('column insert shifts formula references and expands ranges', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = {formula: 'B2+SUM(C1:E1)', result: 0};

      // 在 C 列前插入 1 列：>=C 的列坐标 +1，B 列保持
      ws.spliceColumns(3, 0, ['inserted']);

      expect(ws.getCell('A1').formula).to.equal('B2+SUM(D1:F1)');
    });

    it('column delete shifts, shrinks and invalidates references', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = {formula: 'B2+C2+E2+SUM(C1:D1)', result: 0};

      // 删除 C、D 两列
      ws.spliceColumns(3, 2);

      // C2 落删除区 -> #REF!；E2 左移到 C2；整段 C1:D1 被删 -> #REF!
      expect(ws.getCell('A1').formula).to.equal('B2+#REF!+C2+SUM(#REF!)');
    });

    it('moves absolute references the same as relative ones', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = {formula: '$C$5+C$5+$C5', result: 0};

      ws.spliceColumns(3, 0, ['x']);

      expect(ws.getCell('A1').formula).to.equal('$D$5+D$5+$D5');
    });

    it('does not touch pseudo references in strings or function names', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = {
        formula: 'CONCATENATE("A10", LOG10(5), A10)',
        result: 0,
      };

      ws.spliceRows(3, 0, ['x']);

      expect(ws.getCell('A1').formula).to.equal('CONCATENATE("A10", LOG10(5), A11)');
    });

    it('adjusts cross-sheet references on other worksheets', () => {
      const wb = new Excel.Workbook();
      const ws1 = wb.addWorksheet('S1');
      const ws2 = wb.addWorksheet('S2');
      ws2.getCell('A1').value = {formula: 'S1!A10+\'S1\'!A2+A2', result: 0};

      // S1 插行：指向 S1 的跨表引用平移，S2 本地的裸 A2 不动（在插入点之前）
      ws1.spliceRows(3, 0, ['x']);

      expect(ws2.getCell('A1').formula).to.equal('S1!A11+\'S1\'!A2+A2');
    });

    it('keeps cross-sheet references pointing at untouched sheets stable', () => {
      const wb = new Excel.Workbook();
      const ws1 = wb.addWorksheet('S1');
      const ws3 = wb.addWorksheet('S3');
      ws3.getCell('A1').value = {formula: 'S3!A10+A10', result: 0};

      ws1.spliceRows(3, 0, ['x']);

      // S3 未发生结构变更，其表上所有引用保持原样
      expect(ws3.getCell('A1').formula).to.equal('S3!A10+A10');
    });

    it('translates shared formula master pointers, refs and slave formulas', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.fillFormula(
        'A1:A4',
        'B1',
        [1, 2, 3, 4]
      );

      // 第 3 行前插入 1 行：共享区间 A1:A4 扩张为 A1:A5，slave 指针随之下移
      ws.spliceRows(3, 0, ['x']);

      expect(ws.getCell('A1').model.ref).to.equal('A1:A5');
      expect(ws.getCell('A4').value.sharedFormula).to.equal('A1');
      expect(ws.getCell('A5').value.sharedFormula).to.equal('A1');
      // 既有共享公式复制语义不回退：slave 公式由 master 相对平移得到
      expect(ws.getCell('A4').formula).to.equal('B4');
      expect(ws.getCell('A5').formula).to.equal('B5');
    });

    it('does not regress existing value and style movement on row splice', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = 1;
      ws.getCell('A2').value = {formula: 'A1', result: 1};
      ws.getCell('A2').numFmt = '0.00';

      ws.spliceRows(2, 0, ['inserted']);

      // 值与样式随单元格平移到第 3 行，公式引用同步平移
      expect(ws.getCell('A2').value).to.equal('inserted');
      expect(ws.getCell('A3').formula).to.equal('A1');
      expect(ws.getCell('A3').numFmt).to.equal('0.00');
    });

    it('does not regress existing defined-name movement', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('S1');
      ws.getCell('A1').value = 1;
      ws.getCell('A4').value = 2;
      ws.getCell('A4').name = 'myName';
      ws.getCell('A1').value = {formula: 'A4', result: 2};

      ws.spliceRows(3, 0, ['x']);

      // 名称随单元格平移到 A5，公式中的 A4 同步平移到 A5
      expect(ws.getCell('A5').name).to.equal('myName');
      expect(ws.getCell('A1').formula).to.equal('A5');
    });
  });
});
