const Inspection = require('../models/Inspection');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// ─── GET All Inspections (list) ───────────────────────────────────────────────
exports.getAllInspections = async (req, res) => {
  try {
    const inspections = await Inspection.find({}, 'buyer poStyle factory inspDate result createdAt')
      .sort({ createdAt: -1 });
    res.json({ success: true, data: inspections });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
};

// ─── GET Single Inspection ────────────────────────────────────────────────────
exports.getInspection = async (req, res) => {
  try {
    const inspection = await Inspection.findById(req.params.id);
    if (!inspection) return res.status(404).json({ success: false, message: 'Not found' });
    res.json({ success: true, data: inspection });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
};

// ─── CREATE Inspection ────────────────────────────────────────────────────────
exports.createInspection = async (req, res) => {
  try {
    const body = parseBody(req.body);
    const inspection = new Inspection(body);
    await inspection.save();
    res.status(201).json({ success: true, data: inspection });
  } catch (err) {
    res.status(400).json({ success: false, message: err.message });
  }
};

// ─── UPDATE Inspection ────────────────────────────────────────────────────────
exports.updateInspection = async (req, res) => {
  try {
    const body = parseBody(req.body);
    body.updatedAt = new Date();
    const inspection = await Inspection.findByIdAndUpdate(req.params.id, body, { new: true, runValidators: true });
    if (!inspection) return res.status(404).json({ success: false, message: 'Not found' });
    res.json({ success: true, data: inspection });
  } catch (err) {
    res.status(400).json({ success: false, message: err.message });
  }
};

// ─── DELETE Inspection ────────────────────────────────────────────────────────
exports.deleteInspection = async (req, res) => {
  try {
    const inspection = await Inspection.findById(req.params.id);
    if (!inspection) return res.status(404).json({ success: false, message: 'Not found' });
    // Delete uploaded images
    (inspection.imageCards || []).forEach(card => {
      if (card.filename) {
        const filePath = path.join(__dirname, '../public/uploads', card.filename);
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      }
    });
    await inspection.deleteOne();
    res.json({ success: true, message: 'Deleted successfully' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
};

// ─── UPLOAD Image ─────────────────────────────────────────────────────────────
exports.uploadImage = async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'No file uploaded' });
    const url = `/uploads/${req.file.filename}`;
    res.json({
      success: true,
      filename: req.file.filename,
      originalName: req.file.originalname,
      mimetype: req.file.mimetype,
      url
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
};

// ─── DOWNLOAD EXCEL ───────────────────────────────────────────────────────────
exports.downloadExcel = async (req, res) => {
  try {
    const inspection = await Inspection.findById(req.params.id);
    if (!inspection) return res.status(404).json({ success: false, message: 'Not found' });

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Inspection Report');

    const BLACK = 'FF000000', YELLOW = 'FFFFD700', AMBER = 'FFFFF3CD';
    const GREEN = 'FF90EE90', RED = 'FFFF9999', LGRAY = 'FFF5F5F5', WHITE = 'FFFFFFFF';
    const thinBorder = { style: 'thin', color: { argb: BLACK } };
    const medBorder  = { style: 'medium', color: { argb: BLACK } };
    const allThin = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
    const allMed  = { top: medBorder, bottom: medBorder, left: medBorder, right: medBorder };

    function solidFill(argb) { return { type: 'pattern', pattern: 'solid', fgColor: { argb } }; }
    function fnt(bold=false, size=11, color=BLACK) { return { name: 'Arial', bold, size, color: { argb: color } }; }
    function aln(h='left', v='center', wrap=false) { return { horizontal: h, vertical: v, wrapText: wrap }; }

    function sc(r, c, val, opts={}) {
      const cell = ws.getCell(r, c);
      if (val !== undefined && val !== null) cell.value = val;
      cell.font      = fnt(opts.bold||false, opts.size||11, opts.color||BLACK);
      cell.alignment = aln(opts.h||'left', opts.v||'center', opts.wrap||false);
      cell.fill      = solidFill(opts.bg||WHITE);
      cell.border    = allThin;
    }

    function mg(r1, c1, r2, c2, val, opts={}) {
      ws.mergeCells(r1, c1, r2, c2);
      sc(r1, c1, val, { h:'center', v:'center', ...opts });
      for (let r=r1; r<=r2; r++)
        for (let c=c1; c<=c2; c++)
          ws.getCell(r,c).border = allThin;
    }

    function sectionHeader(r, text) {
      ws.getRow(r).height = 20;
      mg(r, 1, r, 7, text, { bold:true, size:12, bg:YELLOW, h:'center' });
    }

    [26,26,20,18,18,16,16].forEach((w,i) => { ws.getColumn(i+1).width = w; });

    let row = 1;
    const d = inspection;

    // TITLE
    ws.getRow(row).height = 34;
    mg(row,1,row,7,'TEXTILE INSPECTION HUB',{bold:true,size:18,bg:WHITE,h:'center'});
    ws.getCell(row,1).border = allMed;
    row++; ws.getRow(row).height = 6; row++;

    // HEADER INFO
    ws.getRow(row).height = 20;
    mg(row,1,row+3,1,'INSPECTION REPORT',{bold:true,size:13,bg:AMBER,h:'center',wrap:true});
    const infoData = [
      [row,   'Buyer',           d.buyer||'',    'DPI', d.dpi||'',     'FRI', d.fri||''],
      [row+1, 'PO / Style #',    d.poStyle||'',  'Factory', d.factory||'', '', ''],
      [row+2, 'Inspection Date', d.inspDate||'', 'Unit', d.unit||'',   '', ''],
      [row+3, 'Re-Inspection',   d.reInspDate||'','', '',               '', ''],
    ];
    infoData.forEach(([r,l1,v1,l2,v2,l3,v3]) => {
      ws.getRow(r).height = 20;
      sc(r,2,l1,{bold:true,bg:AMBER,h:'left'});
      sc(r,3,v1,{bg:WHITE,h:'left'});
      sc(r,4,l2,{bold:true,bg:AMBER,h:'left'});
      sc(r,5,v2,{bg:WHITE,h:'left'});
      sc(r,6,l3,{bold:true,bg:AMBER,h:'left'});
      sc(r,7,v3,{bg:WHITE,h:'left'});
    });
    row += 4; ws.getRow(row).height = 6; row++;

    // ARTICLE INFO
    const artRows = [
      ['Article:',                  d.article||'',        'AQL - Critical Def:', d.aqlCritical||'N/A'],
      ['Colour:',                   d.colour||'',         'AQL - Major Defect:', d.aqlMajor||'2.5'],
      ['Order Quantity (Pcs):',     d.orderQtyPcs||'',    'AQL - Minor Defect:', d.aqlMinor||''],
      ['Offered Quantity (Pcs):',   d.offeredQtyPcs||'',  'Re-Inspection:',      d.reInspYes?'YES':'NO'],
      ['Order Qty (Cartons):',      d.orderQtyCtns||'',   'Approved Sample:',    d.approvedYes?'YES':'NO'],
      ['Offered Qty (Cartons):',    d.offeredQtyCtns||'', 'Inspector Name:',     d.inspectorName||''],
      ['Units Per Carton:',         d.unitsPerCarton||'', 'Inspection Location:',d.inspLocation||''],
    ];
    artRows.forEach(([l1,v1,l2,v2]) => {
      ws.getRow(row).height = 18;
      sc(row,1,l1,{bold:true,bg:AMBER,h:'left'});
      mg(row,2,row,3,v1,{bg:WHITE,h:'left'});
      sc(row,4,l2,{bold:true,bg:AMBER,h:'left'});
      mg(row,5,row,7,v2,{bg:WHITE,h:'left'});
      row++;
    });
    ws.getRow(row).height = 6; row++;

    // OVERALL RESULT
    sectionHeader(row,'INSPECTION OVERALL RESULT'); row++;
    ws.getRow(row).height = 24;
    const passText = d.result==='PASS'    ? '✓  PASS'    : 'PASS';
    const failText = d.result==='FAIL'    ? '✓  FAIL, due to beyond AQL and/or remark' : 'FAIL, due to beyond AQL and/or remark';
    const pendText = d.result==='PENDING' ? '✓  PENDING' : 'PENDING';
    mg(row,1,row,2,passText,{bold:true,size:12,bg:GREEN,h:'center'});
    mg(row,3,row,5,failText,{bold:true,bg:RED,h:'center',wrap:true});
    mg(row,6,row,7,pendText,{bold:true,size:12,bg:WHITE,h:'center'}); row++;
    ws.getRow(row).height = 18;
    mg(row,1,row,7,d.resultNote||'',{h:'left'}); row++;
    ws.getRow(row).height = 6; row++;

    // SAMPLE SIZE
    ws.getRow(row).height = 22;
    sc(row,1,'SAMPLE SIZE',{bold:true,bg:AMBER,h:'center'});
    mg(row,2,row,4,`NO. OF PIECES:   ${d.samplePcs||'315'}   pcs`,{bold:true,bg:WHITE,h:'center'});
    mg(row,5,row,7,`SELECTED CARTONS:   ${d.selectedCartons||'56'}`,{bold:true,bg:WHITE,h:'center'});
    row++; ws.getRow(row).height = 6; row++;

    // CATEGORY CHECKLIST
    sectionHeader(row,'CATEGORY CHECKLIST'); row++;
    ws.getRow(row).height = 18;
    mg(row,1,row,3,'Category',{bold:true,bg:YELLOW,h:'center'});
    sc(row,4,'OK',{bold:true,bg:YELLOW,h:'center'});
    sc(row,5,'NOT OK',{bold:true,bg:YELLOW,h:'center'});
    mg(row,6,row,7,'Attachment',{bold:true,bg:YELLOW,h:'center'}); row++;
    (d.categories||[]).forEach(cat => {
      ws.getRow(row).height = 18;
      mg(row,1,row,3,cat.name||'',{bg:WHITE,h:'left'});
      sc(row,4,cat.ok?'✓':'',{h:'center',bg:WHITE,bold:cat.ok,color:cat.ok?'FF006600':BLACK});
      sc(row,5,cat.notok?'✗':'',{h:'center',bg:WHITE,bold:cat.notok,color:cat.notok?'FFCC0000':BLACK});
      mg(row,6,row,7,cat.attachment||'',{bg:WHITE,h:'left'}); row++;
    });
    ws.getRow(row).height = 6; row++;

    // DEFECT LIST
    sectionHeader(row,'DEFECT LIST'); row++;
    ws.getRow(row).height = 18;
    sc(row,1,'S.NO',{bold:true,bg:YELLOW,h:'center'});
    mg(row,2,row,4,'Defect Name',{bold:true,bg:YELLOW,h:'center'});
    sc(row,5,'Critical',{bold:true,bg:YELLOW,h:'center'});
    sc(row,6,'Major',{bold:true,bg:YELLOW,h:'center'});
    sc(row,7,'Minor',{bold:true,bg:YELLOW,h:'center'}); row++;
    const dStart = row;
    (d.defects||[]).forEach((def, i) => {
      ws.getRow(row).height = 18;
      sc(row,1,i+1,{h:'center'});
      mg(row,2,row,4,def.name||'',{bg:WHITE,h:'left'});
      sc(row,5,def.critical||'',{h:'center',bg:WHITE});
      sc(row,6,def.major||'',{h:'center',bg:WHITE});
      sc(row,7,def.minor||'',{h:'center',bg:WHITE}); row++;
    });
    const dEnd = row - 1;
    ws.getRow(row).height = 18;
    mg(row,1,row,4,'Total Found Defects:',{bold:true,bg:AMBER,h:'right'});
    sc(row,5,{formula:`SUM(E${dStart}:E${dEnd})`},{h:'center',bold:true,bg:AMBER});
    sc(row,6,{formula:`SUM(F${dStart}:F${dEnd})`},{h:'center',bold:true,bg:AMBER});
    sc(row,7,{formula:`SUM(G${dStart}:G${dEnd})`},{h:'center',bold:true,bg:AMBER}); row++;
    ws.getRow(row).height = 18;
    mg(row,1,row,4,'Allowed Defects According to AQL:',{bold:true,h:'right'});
    sc(row,5,d.aqlAllowCritical||'N/A',{h:'center'});
    sc(row,6,d.aqlAllowMajor||'14',{h:'center'});
    sc(row,7,d.aqlAllowMinor||'21',{h:'center'}); row++;
    ws.getRow(row).height = 6; row++;

    // CARTON DETAILS
    sectionHeader(row,'ARTICLE/COLOR WISE SELECTED CARTON DETAILS'); row++;
    ws.getRow(row).height = 18;
    sc(row,1,'Article/Color Name',{bold:true,bg:YELLOW,h:'center'});
    mg(row,2,row,3,'Quantity Checked',{bold:true,bg:YELLOW,h:'center'});
    mg(row,4,row,5,'Size Ratio Checked',{bold:true,bg:YELLOW,h:'center'});
    sc(row,6,'Selected Carton',{bold:true,bg:YELLOW,h:'center'});
    sc(row,7,'Condition',{bold:true,bg:YELLOW,h:'center'}); row++;
    (d.cartons||[]).forEach(ctn => {
      ws.getRow(row).height = 36;
      sc(row,1,ctn.article||'',{bg:WHITE});
      mg(row,2,row,3,ctn.qty||'',{bg:WHITE,h:'left',wrap:true});
      mg(row,4,row,5,ctn.ratio||'',{bg:WHITE,h:'left'});
      sc(row,6,ctn.selected||'',{h:'center',bg:WHITE});
      sc(row,7,ctn.condition||'',{h:'center',bg:WHITE}); row++;
    });
    ws.getRow(row).height = 6; row++;

    // INSPECTOR REMARKS
    sectionHeader(row,'INSPECTOR REMARKS'); row++;
    ws.getRow(row).height = 80;
    mg(row,1,row,7,d.inspectorRemarks||'',{bg:WHITE,h:'left',v:'top',wrap:true}); row++;
    ws.getRow(row).height = 6; row++;

    // PICTORIAL DETAILS - with real images
    sectionHeader(row,'PICTORIAL DETAILS'); row++;
    ws.getRow(row).height = 16;
    mg(row,1,row,3,'Image (Left)',{bold:true,bg:YELLOW,h:'center'});
    mg(row,4,row,7,'Image (Right)',{bold:true,bg:YELLOW,h:'center'}); row++;

    const cards = d.imageCards || [];
    for (let i = 0; i < cards.length; i += 2) {
      const rowH = 130;
      ws.getRow(row).height = rowH;

      // Caption label row
      const cap1 = cards[i].caption || '';
      const cap2 = cards[i+1] ? (cards[i+1].caption || '') : '';

      // Caption row
      ws.getRow(row).height = 18;
      mg(row,1,row,3, cap1, {bold:true,bg:AMBER,h:'center',v:'center'});
      if (cards[i+1]) {
        mg(row,4,row,7, cap2, {bold:true,bg:AMBER,h:'center',v:'center'});
      } else {
        mg(row,4,row,7,'',{bg:WHITE});
      }
      row++;

      // Image row
      ws.getRow(row).height = 165;

      // Left image
      if (cards[i].filename) {
        const imgPath = path.join(__dirname, '../public/uploads', cards[i].filename);
        if (fs.existsSync(imgPath)) {
          try {
            const ext = path.extname(cards[i].filename).toLowerCase().replace('.','');
            const imageType = ext === 'jpg' ? 'jpeg' : ext;
            const imageId = workbook.addImage({ filename: imgPath, extension: imageType });
            ws.addImage(imageId, { tl: { col: 0, row: row - 1 }, br: { col: 3, row: row } });
          } catch(e) { /* skip if image embed fails */ }
        }
      }
      mg(row,1,row,3,'',{bg:LGRAY});

      // Right image
      if (cards[i+1]) {
        if (cards[i+1].filename) {
          const imgPath2 = path.join(__dirname, '../public/uploads', cards[i+1].filename);
          if (fs.existsSync(imgPath2)) {
            try {
              const ext2 = path.extname(cards[i+1].filename).toLowerCase().replace('.','');
              const imageType2 = ext2 === 'jpg' ? 'jpeg' : ext2;
              const imageId2 = workbook.addImage({ filename: imgPath2, extension: imageType2 });
              ws.addImage(imageId2, { tl: { col: 3, row: row - 1 }, br: { col: 7, row: row } });
            } catch(e) { /* skip */ }
          }
        }
        mg(row,4,row,7,'',{bg:LGRAY});
      } else {
        mg(row,4,row,7,'',{bg:WHITE});
      }
      row++;

      // Label row
      ws.getRow(row).height = 16;
      const lbl1 = cards[i].label || '';
      const lbl2 = cards[i+1] ? (cards[i+1].label || '') : '';
      mg(row,1,row,3, lbl1, {bg:WHITE,h:'center',v:'center'});
      if (cards[i+1]) {
        mg(row,4,row,7, lbl2, {bg:WHITE,h:'center',v:'center'});
      } else {
        mg(row,4,row,7,'',{bg:WHITE});
      }
      row++;
    }

    // Send file
    const filename = `Inspection_${d.buyer||'Report'}_${Date.now()}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    const buffer = await workbook.xlsx.writeBuffer();
    res.send(buffer);

  } catch (err) {
    console.error('Excel export error:', err);
    res.status(500).json({ success: false, message: err.message });
  }
};

// ─── Helper: parse body ───────────────────────────────────────────────────────
function parseBody(body) {
  const parsed = { ...body };
  // Parse JSON strings from form
  if (typeof parsed.categories === 'string') {
    try { parsed.categories = JSON.parse(parsed.categories); } catch(e) { parsed.categories = []; }
  }
  if (typeof parsed.defects === 'string') {
    try { parsed.defects = JSON.parse(parsed.defects); } catch(e) { parsed.defects = []; }
  }
  if (typeof parsed.cartons === 'string') {
    try { parsed.cartons = JSON.parse(parsed.cartons); } catch(e) { parsed.cartons = []; }
  }
  if (typeof parsed.imageCards === 'string') {
    try { parsed.imageCards = JSON.parse(parsed.imageCards); } catch(e) { parsed.imageCards = []; }
  }
  // Booleans
  ['reInspYes','reInspNo','approvedYes','approvedNo'].forEach(k => {
    parsed[k] = parsed[k] === 'true' || parsed[k] === true;
  });
  return parsed;
}
