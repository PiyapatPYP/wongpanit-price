/**
 * วงษ์พาณิชย์ — ระบบราคารับซื้อ (Google Apps Script)
 * วิธีใช้: ดูไฟล์ SETUP.md
 */
const PIN = '1234';          // ← เปลี่ยนเป็นรหัสของร้าน (ใช้ตอนประกาศราคา / เพิ่มสินค้า)
const HISTORY_DAYS = 100;    // จำนวนวันย้อนหลังที่ส่งไปแสดงกราฟ

const S_CAT = 'หมวดหมู่', S_ITEM = 'สินค้า', S_HIST = 'ประวัติราคา';
const H_CAT = ['รหัสหมวด','ชื่อหมวดหมู่','ลำดับ'];
const H_ITEM = ['รหัสสินค้า','รหัสหมวด','ชื่อสินค้า','หน่วย','หมายเหตุ','ราคาทั่วไป','บวกเพิ่ม A','บวกเพิ่ม B','ลำดับ','แก้ไขล่าสุด'];
const H_HIST = ['วันเวลา','รหัสประกาศ','รหัสสินค้า','ชื่อสินค้า','หน่วย','ราคาทั่วไป','บวกเพิ่ม A','บวกเพิ่ม B','ราคาลูกค้า A','ราคาลูกค้า B','หมายเหตุ'];

function doGet(e) {
  return json(loadAll());
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad request' }); }
  if (String(b.pin) !== String(PIN)) return json({ ok: false, error: 'pin' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.action === 'publish') return json(publish(b));
    if (b.action === 'addProduct') return json(addProduct(b));
    return json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** รันครั้งเดียวตอนตั้งค่า: สร้างแท็บและใส่สินค้าเริ่มต้น */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const get = n => ss.getSheetByName(n) || ss.insertSheet(n);
  const sc = get(S_CAT), si = get(S_ITEM), sh = get(S_HIST);
  if (si.getLastRow() > 1) throw new Error('มีข้อมูลอยู่แล้ว — ถ้าต้องการเริ่มใหม่ ให้ลบแท็บ หมวดหมู่ / สินค้า / ประวัติราคา ก่อน');
  [[sc, H_CAT], [si, H_ITEM], [sh, H_HIST]].forEach(([s, h]) => {
    s.clear();
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#F2EDE7');
    s.setFrozenRows(1);
  });
  const now = new Date();
  sc.getRange(2, 1, SEED.cats.length, 3).setValues(SEED.cats.map((c, i) => [c[0], c[1], i + 1]));
  si.getRange(2, 1, SEED.items.length, 10).setValues(SEED.items.map((it, i) => [it.id, it.cat, it.name, it.unit, it.note, it.base, it.mA, it.mB, i + 1, now]));
  sh.getRange(2, 1, SEED.items.length, 11).setValues(SEED.items.map(it => [now, 'setup', it.id, it.name, it.unit, it.base, it.mA, it.mB, round(it.base + it.mA), round(it.base + it.mB), it.note]));
  si.getRange('F:H').setNumberFormat('0.00');
  sh.getRange('F:J').setNumberFormat('0.00');
  si.getRange('J:J').setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
  ['Sheet1', 'แผ่น1', 'ชีต1'].forEach(n => { const d = ss.getSheetByName(n); if (d && ss.getSheets().length > 3) ss.deleteSheet(d); });
  Logger.log('ตั้งค่าเสร็จแล้ว: ' + SEED.items.length + ' สินค้า, ' + SEED.cats.length + ' หมวดหมู่');
}

function loadAll() {
  const cats = rows(S_CAT).filter(r => r[0] !== '').sort((a, b) => (a[2] || 0) - (b[2] || 0)).map(r => [String(r[0]), String(r[1])]);
  const items = rows(S_ITEM).filter(r => r[0] !== '').sort((a, b) => (a[8] || 0) - (b[8] || 0)).map(r => ({
    id: String(r[0]), cat: String(r[1]), name: String(r[2]), unit: String(r[3] || 'กก.'), note: String(r[4] || ''),
    base: num(r[5]), mA: num(r[6]), mB: num(r[7])
  }));
  const cutoff = Date.now() - HISTORY_DAYS * 864e5, last2 = {}, out = [], seen = {};
  rows(S_HIST).forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const h = { t: r[0].toISOString(), pid: String(r[1]), id: String(r[2]), base: num(r[5]), mA: num(r[6]), mB: num(r[7]) };
    const l = last2[h.id] = last2[h.id] || []; l.push(h); if (l.length > 2) l.shift();
    if (r[0].getTime() >= cutoff) { out.push(h); seen[h.pid + '|' + h.id] = 1; }
  });
  Object.keys(last2).forEach(k => last2[k].forEach(h => { const key = h.pid + '|' + h.id; if (!seen[key]) { out.push(h); seen[key] = 1; } }));
  out.sort((a, b) => a.t < b.t ? -1 : 1);
  return { ok: true, cats, items, history: out, serverTime: new Date().toISOString() };
}

function publish(b) {
  const ss = SpreadsheetApp.getActive(), si = ss.getSheetByName(S_ITEM), sh = ss.getSheetByName(S_HIST);
  const data = si.getRange(2, 1, si.getLastRow() - 1, 10).getValues();
  const idx = {}; data.forEach((r, i) => idx[String(r[0])] = i);
  const now = new Date(), hist = [];
  (b.changes || []).forEach(c => {
    const i = idx[c.id]; if (i === undefined) return;
    const r = data[i];
    r[3] = c.unit; r[4] = c.note; r[5] = c.base; r[6] = c.mA; r[7] = c.mB; r[9] = now;
    hist.push([now, b.pid, c.id, r[2], c.unit, c.base, c.mA, c.mB, round(c.base + c.mA), round(c.base + c.mB), c.note]);
  });
  si.getRange(2, 1, data.length, 10).setValues(data);
  if (hist.length) sh.getRange(sh.getLastRow() + 1, 1, hist.length, 11).setValues(hist);
  return { ok: true, t: now.toISOString(), count: hist.length };
}

function addProduct(b) {
  const ss = SpreadsheetApp.getActive(), now = new Date(), it = b.item;
  if (b.cat) { const sc = ss.getSheetByName(S_CAT); sc.appendRow([b.cat.id, b.cat.name, sc.getLastRow()]); }
  const si = ss.getSheetByName(S_ITEM), sh = ss.getSheetByName(S_HIST);
  si.appendRow([it.id, it.cat, it.name, it.unit, it.note, it.base, it.mA, it.mB, si.getLastRow(), now]);
  sh.appendRow([now, b.pid, it.id, it.name, it.unit, it.base, it.mA, it.mB, round(it.base + it.mA), round(it.base + it.mB), it.note]);
  return { ok: true, t: now.toISOString() };
}

function rows(name) {
  const s = SpreadsheetApp.getActive().getSheetByName(name);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues();
}
function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : Math.round(n * 100) / 100; }
function round(v) { return Math.round(v * 100) / 100; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const SEED = {"cats":[["metal","เหล็ก"],["alu","อลูมิเนียม"],["battery","แบตเตอรี่ / มอเตอร์"],["paper","กระดาษ"],["glass","แก้ว"],["copper","ทองแดง"],["steel","สแตนเลส"],["oil","น้ำมันพืชใช้แล้ว"],["lead","ตะกั่ว"],["brass","ทองเหลือง"],["plastic","พลาสติก"],["ewaste","ขยะอิเล็กทรอนิกส์"]],"items":[{"id":"i0","cat":"metal","name":"เหล็กหนา ตัดสั้น 80 ซม.","unit":"กก.","note":"ราคานี้ นน.ตั้งแต่ 200 กก.ขึ้นไป (น้อยกว่า 200 กก. = 8 บาท)","base":8.2,"mA":0,"mB":0},{"id":"i1","cat":"metal","name":"เหล็กหนายาว / เศษเหล็กเล็กๆ","unit":"กก.","note":"ยาวเกิน 1 เมตร","base":7.5,"mA":0,"mB":0},{"id":"i2","cat":"metal","name":"เหล็กหนาชิ้นใหญ่ตัดยาก","unit":"กก.","note":"","base":7,"mA":0,"mB":0},{"id":"i3","cat":"metal","name":"เหล็กบาง / เมทัลชีท","unit":"กก.","note":"ลวดล้อเผา 4 บ.","base":7,"mA":0.1,"mB":0},{"id":"i4","cat":"metal","name":"เหล็กรวม","unit":"กก.","note":"","base":7.2,"mA":0,"mB":0},{"id":"i5","cat":"metal","name":"เหล็กหล่อชิ้นเล็ก / จักรเจาะแล้ว","unit":"กก.","note":"","base":8.4,"mA":0.1,"mB":0},{"id":"i6","cat":"metal","name":"จักรรีดยาง","unit":"กก.","note":"","base":7.5,"mA":0.2,"mB":0},{"id":"i7","cat":"metal","name":"หล่อใหญ่ / หล่อเครื่อง","unit":"กก.","note":"","base":6.9,"mA":0.1,"mB":0},{"id":"i8","cat":"metal","name":"โช้คอัพ","unit":"กก.","note":"","base":5,"mA":0,"mB":0},{"id":"i9","cat":"metal","name":"สังกะสีแผ่น","unit":"กก.","note":"","base":3.1,"mA":0.1,"mB":0},{"id":"i10","cat":"metal","name":"กระป๋องเหล็ก / กระป๋องสังกะสี","unit":"กก.","note":"","base":4.3,"mA":0.1,"mB":0},{"id":"i11","cat":"alu","name":"อลูมิเนียมบาง/อ่อน สะอาด","unit":"กก.","note":"","base":73,"mA":1,"mB":0},{"id":"i12","cat":"alu","name":"อลูมิเนียมบาง/อ่อน ติดสี ติดสติ๊กเกอร์","unit":"กก.","note":"","base":71,"mA":1,"mB":0},{"id":"i13","cat":"alu","name":"อลูมิเนียมไส้ในมอเตอร์, สเปรย์","unit":"กก.","note":"","base":71,"mA":1,"mB":0},{"id":"i14","cat":"alu","name":"อลูมิเนียมฉลาก สะอาด","unit":"กก.","note":"","base":90,"mA":1,"mB":0},{"id":"i15","cat":"alu","name":"อลูมิเนียมฉลากติดสี / สติ๊กเกอร์ (-2%)","unit":"กก.","note":"","base":87,"mA":1,"mB":0},{"id":"i16","cat":"alu","name":"อลูมิเนียมหนา / แข็ง / หนาเครื่อง","unit":"กก.","note":"","base":67,"mA":1,"mB":0},{"id":"i17","cat":"alu","name":"อลูมิเนียมล้อแม็กซ์","unit":"กก.","note":"","base":91,"mA":1,"mB":0},{"id":"i18","cat":"alu","name":"อลูมิเนียมกระป๋อง (ไม่ปนกระป๋องสเปรย์)","unit":"กก.","note":"","base":72,"mA":1,"mB":0},{"id":"i19","cat":"alu","name":"อลูมิเนียมสายไฟ (เส้น)","unit":"กก.","note":"","base":97,"mA":1,"mB":0},{"id":"i20","cat":"alu","name":"อลูมิเนียมหม้อน้ำไม่ติดเหล็ก","unit":"กก.","note":"","base":51,"mA":1,"mB":0},{"id":"i21","cat":"alu","name":"อลูมิเนียมกระทะผัด / กระทะไฟฟ้า / ก้นหม้อ","unit":"กก.","note":"","base":44,"mA":1,"mB":0},{"id":"i22","cat":"alu","name":"อลูมิเนียมผ้าเบรกแกะ, ลูกสูบ, คาบูลู","unit":"กก.","note":"","base":67,"mA":1,"mB":0},{"id":"i23","cat":"alu","name":"อลูมิเนียมผ้าเบรกไม่แกะ","unit":"กก.","note":"","base":46,"mA":1,"mB":0},{"id":"i24","cat":"alu","name":"อลูมิเนียมมู่ลี่กันสาด","unit":"กก.","note":"","base":34,"mA":1,"mB":0},{"id":"i25","cat":"alu","name":"อลูมิเนียมมุ้งลวด","unit":"กก.","note":"","base":28,"mA":1,"mB":0},{"id":"i26","cat":"alu","name":"อลูมิเนียมฝาจุกแกะ","unit":"กก.","note":"","base":53,"mA":1,"mB":0},{"id":"i27","cat":"alu","name":"อลูมิเนียมฝาจุกไม่แกะ","unit":"กก.","note":"","base":18,"mA":1,"mB":0},{"id":"i28","cat":"alu","name":"อลูมิเนียมหม้อน้ำแอร์แกะ","unit":"กก.","note":"","base":47,"mA":1,"mB":0},{"id":"i29","cat":"alu","name":"อัลลอยด์","unit":"กก.","note":"","base":51,"mA":1,"mB":0},{"id":"i30","cat":"battery","name":"แบตเตอรี่ใหญ่ (เทน้ำออกหมด)","unit":"กก.","note":"","base":22.9,"mA":0.2,"mB":0},{"id":"i31","cat":"battery","name":"แบตเตอรี่ดำ / แบตแห้ง","unit":"กก.","note":"","base":17.6,"mA":0.2,"mB":0},{"id":"i32","cat":"battery","name":"แบตเตอรี่มอเตอร์ไซค์","unit":"กก.","note":"แบตจีน งดรับ","base":17.9,"mA":0.2,"mB":0},{"id":"i33","cat":"paper","name":"กระดาษลัง","unit":"กก.","note":"ราคานี้ 200 กก.ขึ้นไป","base":3.9,"mA":0,"mB":0},{"id":"i34","cat":"paper","name":"กระดาษจั๊บ / กระดาษสี","unit":"กก.","note":"","base":2.2,"mA":0.1,"mB":0},{"id":"i35","cat":"paper","name":"กระดาษขาวดำ ไม่ระบายสี ไม่ปริ้นสี","unit":"กก.","note":"","base":5.3,"mA":0.1,"mB":0},{"id":"i36","cat":"glass","name":"แก้วสีชา","unit":"กก.","note":"","base":0.4,"mA":0,"mB":0},{"id":"i37","cat":"glass","name":"แก้วสีขาว","unit":"กก.","note":"","base":1.4,"mA":0,"mB":0},{"id":"i38","cat":"glass","name":"แก้วสีเขียว","unit":"กก.","note":"","base":1.4,"mA":0,"mB":0},{"id":"i39","cat":"copper","name":"ทองแดงปอกสวย เบอร์ 1","unit":"กก.","note":"","base":461,"mA":0,"mB":0},{"id":"i40","cat":"copper","name":"ทองแดงปอกช็อต เบอร์ 2","unit":"กก.","note":"","base":451,"mA":0,"mB":0},{"id":"i41","cat":"copper","name":"ทองแดงใหญ่ ไม่เผา","unit":"กก.","note":"","base":433,"mA":0,"mB":0},{"id":"i42","cat":"copper","name":"ทองแดงเล็ก มอเตอร์ ไม่เผา","unit":"กก.","note":"","base":426,"mA":0,"mB":0},{"id":"i43","cat":"copper","name":"ทองแดงเคลือบ","unit":"กก.","note":"","base":418,"mA":0,"mB":0},{"id":"i44","cat":"copper","name":"หม้อน้ำทองแดงไม่ติดเหล็ก เป็นแผงสวย","unit":"กก.","note":"ไม่สวย, ท่อไม่ครบ -5 บาท","base":230,"mA":0,"mB":0},{"id":"i45","cat":"copper","name":"ทองแดงเผา","unit":"กก.","note":"","base":421,"mA":5,"mB":0},{"id":"i46","cat":"steel","name":"สแตนเลส 304","unit":"กก.","note":"","base":36,"mA":1,"mB":0},{"id":"i47","cat":"steel","name":"สแตนเลส 304 มีรอยเชื่อม","unit":"กก.","note":"","base":33,"mA":1,"mB":0},{"id":"i48","cat":"steel","name":"สแตนเลส 304 ชิ้นใหญ่ ไม่เชื่อม","unit":"กก.","note":"","base":32,"mA":1,"mB":0},{"id":"i49","cat":"steel","name":"สแตนเลสเทียม","unit":"กก.","note":"","base":9,"mA":0,"mB":0},{"id":"i50","cat":"oil","name":"น้ำมันพืช (ไม่ผสมน้ำ และน้ำมันชนิดอื่น)","unit":"กก.","note":"","base":30,"mA":0,"mB":0},{"id":"i51","cat":"lead","name":"ตะกั่วแข็ง","unit":"กก.","note":"","base":70,"mA":1,"mB":0},{"id":"i52","cat":"lead","name":"ตะกั่วอ่อน","unit":"กก.","note":"","base":45,"mA":1,"mB":0},{"id":"i53","cat":"lead","name":"แผ่น CD","unit":"กก.","note":"","base":10,"mA":0.2,"mB":0},{"id":"i54","cat":"brass","name":"ทองเหลืองหนา","unit":"กก.","note":"","base":282,"mA":0,"mB":0},{"id":"i55","cat":"brass","name":"ทองเหลืองบาง","unit":"กก.","note":"","base":266,"mA":0,"mB":0},{"id":"i56","cat":"brass","name":"ทองเหลืองอูดติด","unit":"กก.","note":"","base":271,"mA":0,"mB":0},{"id":"i57","cat":"brass","name":"หม้อน้ำทองเหลือง ไม่ติดเหล็ก","unit":"กก.","note":"","base":246,"mA":0,"mB":0},{"id":"i58","cat":"plastic","name":"เพทใส สะอาด","unit":"กก.","note":"","base":7.05,"mA":0.2,"mB":0},{"id":"i59","cat":"plastic","name":"เพทสกรีน","unit":"กก.","note":"","base":3.3,"mA":0.2,"mB":0},{"id":"i60","cat":"plastic","name":"พลาสติกรวม ไม่มีเพทเจือปน","unit":"กก.","note":"","base":3.4,"mA":0.1,"mB":0},{"id":"i61","cat":"plastic","name":"พลาสติกดำ","unit":"กก.","note":"","base":1.2,"mA":0.1,"mB":0},{"id":"i62","cat":"plastic","name":"ท่อ PVC ฟ้า","unit":"กก.","note":"","base":2,"mA":0.2,"mB":0},{"id":"i63","cat":"plastic","name":"ข้อต่อ PVC ฟ้า","unit":"กก.","note":"","base":1,"mA":0.2,"mB":0},{"id":"i64","cat":"plastic","name":"สายยางอ่อน","unit":"กก.","note":"","base":1.5,"mA":0.1,"mB":0},{"id":"i65","cat":"plastic","name":"รองเท้าบู๊ท PVC","unit":"กก.","note":"","base":8,"mA":0.1,"mB":0},{"id":"i66","cat":"plastic","name":"รองเท้าบู๊ทสี / ตราคบ / สั้น","unit":"กก.","note":"","base":1,"mA":0,"mB":0},{"id":"i67","cat":"plastic","name":"เป่าใส 2","unit":"กก.","note":"","base":11,"mA":0,"mB":0},{"id":"i68","cat":"plastic","name":"เป่าขุ่น (ไม่รวมขวดน้ำมันเครื่อง)","unit":"กก.","note":"","base":4,"mA":0,"mB":0},{"id":"i69","cat":"plastic","name":"น้ำเกลือ แกะพลาสติกหุ้มและฝา","unit":"กก.","note":"","base":4.7,"mA":0,"mB":0},{"id":"i70","cat":"ewaste","name":"ทีวีจอแก้ว 21 นิ้วขึ้นไป","unit":"ชิ้น","note":"","base":50,"mA":0,"mB":0},{"id":"i71","cat":"ewaste","name":"ทีวีจอแก้วเล็ก","unit":"ชิ้น","note":"","base":40,"mA":0,"mB":0},{"id":"i72","cat":"ewaste","name":"ทีวีจอแบน","unit":"ชิ้น","note":"","base":30,"mA":0,"mB":0},{"id":"i73","cat":"ewaste","name":"หน้าจอคอมพิวเตอร์","unit":"ชิ้น","note":"","base":20,"mA":0,"mB":0},{"id":"i74","cat":"ewaste","name":"เครื่องซักผ้า มีมอเตอร์","unit":"ชิ้น","note":"","base":80,"mA":0,"mB":0},{"id":"i75","cat":"ewaste","name":"ตู้เย็น มีมอเตอร์","unit":"ชิ้น","note":"","base":80,"mA":0,"mB":0},{"id":"i76","cat":"ewaste","name":"แผงวงจร รวมสี","unit":"กก.","note":"","base":5,"mA":1,"mB":0},{"id":"i77","cat":"ewaste","name":"แผงวงจรเขียว มีชิป","unit":"กก.","note":"","base":15,"mA":1,"mB":0}]};
