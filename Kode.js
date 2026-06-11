const SOURCE_SPREADSHEET_ID = '1RfS8ld3uUiC_5zmZNKqVgOV7zcCrVxd57L2a0AsNUJc';
const OUTPUT_NAME = 'Instrumen Penjaminan Mutu Pelatihan';
const ENABLE_FILE_UPLOAD_QUESTIONS = true;
const EXISTING_MANDIRI_FORM_ID = '1uHzZrUlV2MuCAxA6DwmzI0q-IoAZJhlY5FUHOKEWH44';
const EXISTING_VERIFIKASI_FORM_ID = '';
const EXISTING_REKAP_SPREADSHEET_ID = '1nJGiUrdr-QCb9GoltjmxMS9BK_i1eGwwx7VjDTLJ8Z8';
const EVIDENCE_BATCH_SIZE = 5;

function setupInstrumenPenjaminanMutu() {
  const source = loadSourceData_();
  const spreadsheet = SpreadsheetApp.create(OUTPUT_NAME + ' - Rekap');

  const masterSheet = resetSheet_(spreadsheet, 'Master Indikator');
  const rekapSheet = resetSheet_(spreadsheet, 'Rekap');
  const detailSheet = resetSheet_(spreadsheet, 'Detail Jawaban');
  const dashboardSheet = resetSheet_(spreadsheet, 'Dashboard');

  writeMaster_(masterSheet, source.indicators);
  writeRekap_(rekapSheet, source.metadata, source.indicators);
  writeDetailJawaban_(detailSheet);
  writeDashboard_(dashboardSheet);

  const mandiriForm = buildMandiriForm_(source, spreadsheet);
  installSubmitTriggers_(spreadsheet, mandiriForm);

  const infoSheet = resetSheet_(spreadsheet, 'Panduan');
  writePanduan_(infoSheet, source, mandiriForm, spreadsheet);

  SpreadsheetApp.flush();

  Logger.log('Google Sheet Rekap: ' + spreadsheet.getUrl());
  Logger.log('Form Penilaian Mandiri Penyelenggara: ' + mandiriForm.getEditUrl());
}

function handleMandiriSubmit(event) {
  upsertMandiriSubmission_(event);
  const assessmentContext = createTrainingAssessmentSheet_(event);
  syncOrganizerProgressSheet_(event, assessmentContext);
}

function handleVerifikasiSubmit(event) {
  syncFormResponse_(event, 'VERIFIKASI');
}

function handlePenilaianSheetEdit(event) {
  if (!event || !event.range) {
    return;
  }

  const sheet = event.range.getSheet();
  const sheetName = sheet.getName();
  const ignoredSheets = ['Rekap', 'Detail Jawaban', 'Dashboard', 'Master Indikator', 'Panduan', 'Form Responses 1'];
  if (ignoredSheets.indexOf(sheetName) !== -1 || event.range.getRow() <= 10) {
    return;
  }

  const editedColumn = event.range.getColumn();
  if (editedColumn < 10 || editedColumn > 12) {
    return;
  }

  const context = buildAssessmentContextFromSheet_(sheet);
  if (context) {
    syncOrganizerProgressContext_(context, '', false);
  }
}

function resyncMandiriResponsesToSummaryModel() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('Isi EXISTING_MANDIRI_FORM_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const form = FormApp.openById(mandiriFormId);
  const responses = form.getResponses();
  responses.forEach(function(response) {
    const event = { response: response };
    upsertMandiriSubmission_(event);
    const assessmentContext = createTrainingAssessmentSheet_(event);
    syncOrganizerProgressSheet_(event, assessmentContext);
  });

  Logger.log('Resync selesai. Total respons diproses ulang: ' + responses.length);
}

function polishExistingForms() {
  const source = loadSourceData_();
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  const verifikasiFormId = EXISTING_VERIFIKASI_FORM_ID || PropertiesService.getScriptProperties().getProperty('VERIFIKASI_FORM_ID');

  if (mandiriFormId) {
    const mandiriForm = openFormSafely_(mandiriFormId, 'Form Mandiri');
    if (mandiriForm) {
      polishFormHelpText_(mandiriForm, source, 'MANDIRI');
      Logger.log('Form Mandiri diperbarui: ' + mandiriForm.getEditUrl());
    }
  }

  if (verifikasiFormId) {
    const verifikasiForm = openFormSafely_(verifikasiFormId, 'Form Verifikasi');
    if (verifikasiForm) {
      polishFormHelpText_(verifikasiForm, source, 'VERIFIKASI');
      Logger.log('Form Verifikasi diperbarui: ' + verifikasiForm.getEditUrl());
    }
  }
}

function upgradeExistingFormQuestionTypes() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');

  if (mandiriFormId) {
    const mandiriForm = openFormSafely_(mandiriFormId, 'Form Mandiri');
    if (mandiriForm) {
      upgradeIdentityQuestionTypes_(mandiriForm);
      polishFormHelpText_(mandiriForm, loadSourceData_(), 'MANDIRI');
      Logger.log('Tipe pertanyaan Form Mandiri diperbarui: ' + mandiriForm.getEditUrl());
    }
  }
}

function styleExistingMandiriQuestionTitles() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('Isi EXISTING_MANDIRI_FORM_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  form.getItems().forEach(function(item) {
    const title = item.getTitle();
    const match = title.match(/^\[(I\d{3})\]\s*(.+)$/);
    if (!match) {
      return;
    }

    const code = match[1];
    const field = normalizeIndicatorField_(match[2]);
    if (field === 'Nilai Mandiri') {
      setItemTitle_(item, '[' + code + '] 𝘕𝘪𝘭𝘢𝘪 𝘔𝘢𝘯𝘥𝘪𝘳𝘪');
    } else if (field === 'Bukti Dukung' || field === 'Lampirkan Link Dokumen') {
      setItemTitle_(item, '[' + code + '] 𝘓𝘢𝘮𝘱𝘪𝘳𝘬𝘢𝘯 𝘓𝘪𝘯𝘬 𝘋𝘰𝘬𝘶𝘮𝘦𝘯');
    } else if (field === 'Keterangan') {
      setItemTitle_(item, '[' + code + '] 𝘒𝘦𝘵𝘦𝘳𝘢𝘯𝘨𝘢𝘯');
    } else if (field === 'Upload Soft Copy' || field === 'Upload File') {
      setItemTitle_(item, '[' + code + '] 𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦');
    }
  });

  Logger.log('Judul pertanyaan Form Mandiri diperbarui: ' + form.getEditUrl());
}

function addEvidenceUploadInputsToExistingMandiriForm() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('Isi EXISTING_MANDIRI_FORM_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  const source = loadSourceData_();
  source.indicators.forEach(function(indicator) {
    ensureEvidenceInputsForIndicator_(form, indicator.code);
  });

  Logger.log('Input bukti dukung Form Mandiri diperbarui: ' + form.getEditUrl());
}

function addEvidenceUploadInputsBatch() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('Isi EXISTING_MANDIRI_FORM_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  const source = loadSourceData_();
  const properties = PropertiesService.getScriptProperties();
  const startIndex = Number(properties.getProperty('EVIDENCE_BATCH_INDEX') || '0');
  const endIndex = Math.min(startIndex + EVIDENCE_BATCH_SIZE, source.indicators.length);
  const itemMap = buildFormItemMap_(form);

  for (let i = startIndex; i < endIndex; i++) {
    ensureEvidenceInputsForIndicator_(form, source.indicators[i].code, itemMap);
  }

  if (endIndex >= source.indicators.length) {
    properties.deleteProperty('EVIDENCE_BATCH_INDEX');
    Logger.log('Selesai. Semua input bukti dukung sudah diproses.');
  } else {
    properties.setProperty('EVIDENCE_BATCH_INDEX', String(endIndex));
    Logger.log('Batch selesai sampai indikator ke-' + endIndex + ' dari ' + source.indicators.length + '. Jalankan addEvidenceUploadInputsBatch lagi untuk melanjutkan.');
  }
}

function resetEvidenceUploadBatchProgress() {
  PropertiesService.getScriptProperties().deleteProperty('EVIDENCE_BATCH_INDEX');
  Logger.log('Progress batch input bukti dukung direset. Jalankan addEvidenceUploadInputsBatch dari awal.');
}

function ensureEvidenceInputsForIndicator_(form, code, itemMap) {
  itemMap = itemMap || buildFormItemMap_(form);
  let linkItem = null;
  let uploadItem = null;

  (itemMap[code] || []).forEach(function(entry) {
    const item = entry.item;
    const field = entry.field;
    if (field === 'Bukti Dukung' || field === 'Lampirkan Link Dokumen') {
      linkItem = item;
    } else if (field === 'Upload Soft Copy' || field === 'Upload File') {
      uploadItem = item;
    }
  });

  if (linkItem) {
    setItemTitle_(linkItem, '[' + code + '] 𝘓𝘢𝘮𝘱𝘪𝘳𝘬𝘢𝘯 𝘓𝘪𝘯𝘬 𝘋𝘰𝘬𝘶𝘮𝘦𝘯');
    setItemHelpText_(linkItem, evidenceLinkHelpText_());
  }

  if (uploadItem) {
    setItemTitle_(uploadItem, '[' + code + '] 𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦');
    setItemHelpText_(uploadItem, evidenceUploadHelpText_());
    return;
  }

  try {
    const newUploadItem = createUploadEvidenceItem_(form, code);

    if (linkItem) {
      form.moveItem(newUploadItem.getIndex(), linkItem.getIndex() + 1);
    }
  } catch (error) {
    Logger.log('Upload File gagal ditambahkan untuk ' + code + ': ' + error.message);
    Logger.log('Jika akun/domain tidak mendukung upload file, gunakan Lampirkan Link Dokumen sebagai fallback.');
  }
}

function buildFormItemMap_(form) {
  const map = {};
  form.getItems().forEach(function(item) {
    const title = item.getTitle();
    const match = title.match(/^\[(I\d{3})\]\s*(.+)$/);
    if (!match) {
      return;
    }

    const code = match[1];
    if (!map[code]) {
      map[code] = [];
    }

    map[code].push({
      item: item,
      field: normalizeIndicatorField_(match[2])
    });
  });
  return map;
}

function createUploadEvidenceItem_(form, code) {
  const title = '[' + code + '] 𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦';
  if (typeof form.addFileUploadItem === 'function') {
    return form.addFileUploadItem()
      .setTitle(title)
      .setHelpText(evidenceUploadHelpText_())
      .setRequired(false);
  }

  return form.addTextItem()
    .setTitle(title)
    .setHelpText(evidenceUploadFallbackHelpText_())
    .setRequired(false);
}

function migrateExistingSheetToSingleFormModel() {
  const spreadsheetId = EXISTING_REKAP_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Isi EXISTING_REKAP_SPREADSHEET_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const rekapSheet = spreadsheet.getSheetByName('Rekap');
  if (!rekapSheet) {
    throw new Error('Tab Rekap tidak ditemukan.');
  }

  rekapSheet.getRange('I9').setValue('Nilai Tim Penilai');
  rekapSheet.getRange('J9').setValue('Catatan Tim Penilai');

  const lastRow = Math.max(rekapSheet.getLastRow(), 10);
  if (lastRow >= 10) {
    const rowCount = lastRow - 9;
    rekapSheet.getRange(10, 9, rowCount, 2).clearContent();
    rekapSheet.getRange(10, 9, rowCount, 2).setBackground('#fff9e6');
    rekapSheet.getRange(10, 11, rowCount, 1).setFormulaR1C1('=IF(RC[-2]<>"",RC[-2],RC[-5])');
  }

  const verifikasiSheet = spreadsheet.getSheetByName('Verifikasi Mutu');
  if (verifikasiSheet) {
    verifikasiSheet.setName('Verifikasi Mutu - Tidak Dipakai');
    verifikasiSheet.hideSheet();
  }

  const panduanSheet = spreadsheet.getSheetByName('Panduan') || spreadsheet.insertSheet('Panduan');
  panduanSheet.clear();
  writeSingleFormPanduan_(panduanSheet, spreadsheet);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'handleVerifikasiSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('Sheet dimigrasikan ke model 1 Form + penilaian tim di Sheet: ' + spreadsheet.getUrl());
}

function migrateExistingRekapToSummaryModel() {
  const spreadsheetId = EXISTING_REKAP_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Isi EXISTING_REKAP_SPREADSHEET_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  backupSheetIfPresent_(spreadsheet, 'Rekap');
  backupSheetIfPresent_(spreadsheet, 'Detail Jawaban');
  backupSheetIfPresent_(spreadsheet, 'Dashboard');
  writeRekap_(resetSheet_(spreadsheet, 'Rekap'), {}, []);
  writeDetailJawaban_(resetSheet_(spreadsheet, 'Detail Jawaban'));
  writeDashboard_(resetSheet_(spreadsheet, 'Dashboard'));
  SpreadsheetApp.flush();

  Logger.log('Tab Rekap, Detail Jawaban, dan Dashboard sudah dimigrasikan ke model ringkasan: ' + spreadsheet.getUrl());
}

function updateExistingRekapColumnsForSimpleModel() {
  const spreadsheetId = EXISTING_REKAP_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Isi EXISTING_REKAP_SPREADSHEET_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('Rekap');
  if (!sheet) {
    throw new Error('Tab Rekap tidak ditemukan.');
  }

  backupSheetIfPresent_(spreadsheet, 'Rekap');

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const columnsToDelete = [];

  headers.forEach(function(header, index) {
    if (header === 'Persentase Capaian' || header === 'Kategori Mutu') {
      columnsToDelete.push(index + 1);
    }
  });

  columnsToDelete.sort(function(a, b) {
    return b - a;
  }).forEach(function(column) {
    sheet.deleteColumn(column);
  });

  const updatedLastColumn = sheet.getLastColumn();
  const updatedHeaders = sheet.getRange(1, 1, 1, updatedLastColumn).getValues()[0];
  updatedHeaders.forEach(function(header, index) {
    if (header === 'Jumlah Terpenuhi') {
      sheet.getRange(1, index + 1).setValue('Jumlah Terisi');
    } else if (header === 'Jumlah Tidak Terpenuhi') {
      sheet.getRange(1, index + 1).setValue('Jumlah Tidak Terisi');
    }
  });

  const finalHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const timestampColumn = finalHeader.indexOf('Terakhir Diubah') + 1;
  if (timestampColumn > 0) {
    sheet.getRange(2, timestampColumn, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  formatHeader_(sheet, 1, sheet.getLastColumn());
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, sheet.getLastColumn());

  writeDashboard_(spreadsheet.getSheetByName('Dashboard') || spreadsheet.insertSheet('Dashboard'));
  SpreadsheetApp.flush();

  Logger.log('Kolom Rekap disederhanakan: Persentase Capaian dan Kategori Mutu dihapus; header jumlah diperbarui.');
}

function backupSheetIfPresent_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) {
    return;
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const backupName = buildSafeSheetName_(sheetName + ' Backup ' + timestamp);
  const backup = sheet.copyTo(spreadsheet).setName(backupName);
  backup.hideSheet();
}

function removeMetodePemeriksaanFromExistingAssets() {
  const spreadsheetId = EXISTING_REKAP_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Isi EXISTING_REKAP_SPREADSHEET_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  spreadsheet.getSheets().forEach(function(sheet) {
    removeMethodColumnsFromSheet_(sheet);
    repairNilaiAkhirFormula_(sheet);
  });

  const panduanSheet = spreadsheet.getSheetByName('Panduan');
  if (panduanSheet) {
    panduanSheet.clear();
    writeSingleFormPanduan_(panduanSheet, spreadsheet);
  }

  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (mandiriFormId) {
    const source = loadSourceData_();
    const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
    if (form) {
      polishFormHelpText_(form, source, 'MANDIRI');
      styleExistingMandiriQuestionTitles();
      Logger.log('Informasi metode pemeriksaan di Form Mandiri diperbarui: ' + form.getEditUrl());
    }
  }

  Logger.log('Metode Pemeriksaan dihapus dari Sheet dan Form yang dapat diakses: ' + spreadsheet.getUrl());
}

function removeMethodColumnsFromSheet_(sheet) {
  const lastRow = Math.min(sheet.getLastRow(), 20);
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    return;
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const columnsToDelete = [];
  values.forEach(function(row) {
    row.forEach(function(value, index) {
      if (String(value || '').trim().toLowerCase() === 'metode pemeriksaan') {
        columnsToDelete.push(index + 1);
      }
    });
  });

  Array.from(new Set(columnsToDelete)).sort(function(a, b) {
    return b - a;
  }).forEach(function(column) {
    sheet.deleteColumn(column);
  });
}

function repairNilaiAkhirFormula_(sheet) {
  const header = findHeaderRowAndColumns_(sheet);
  if (!header) {
    return;
  }

  const mandiriCol = header.columns['Nilai Mandiri'];
  const timCol = header.columns['Nilai Tim Penilai'];
  const akhirCol = header.columns['Nilai Akhir'];
  if (!mandiriCol || !timCol || !akhirCol) {
    return;
  }

  const firstDataRow = header.row + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) {
    return;
  }

  for (let row = firstDataRow; row <= lastRow; row++) {
    const firstCell = sheet.getRange(row, 1).getDisplayValue();
    if (!firstCell || String(firstCell).indexOf('Jumlah') === 0) {
      continue;
    }

    const formula = '=IF(' + columnToLetter_(timCol) + row + '<>"",' + columnToLetter_(timCol) + row + ',' + columnToLetter_(mandiriCol) + row + ')';
    sheet.getRange(row, akhirCol).setFormula(formula);
  }
}

function findHeaderRowAndColumns_(sheet) {
  const lastRow = Math.min(sheet.getLastRow(), 20);
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    return null;
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    const columns = {};
    values[r].forEach(function(value, index) {
      const key = String(value || '').trim();
      if (key) {
        columns[key] = index + 1;
      }
    });

    if (columns['Nilai Mandiri'] && columns['Nilai Akhir']) {
      return { row: r + 1, columns: columns };
    }
  }

  return null;
}

function columnToLetter_(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function openFormSafely_(formIdOrUrl, label) {
  try {
    const formId = extractFormId_(formIdOrUrl);
    return FormApp.openById(formId);
  } catch (error) {
    try {
      const value = String(formIdOrUrl || '').trim();
      if (value.indexOf('https://docs.google.com/forms/') === 0) {
        return FormApp.openByUrl(value);
      }
    } catch (urlError) {
      Logger.log(label + ' juga gagal dibuka sebagai URL: ' + urlError.message);
    }
    Logger.log(label + ' dilewati. ID/URL tidak valid atau akun tidak punya akses edit: ' + formIdOrUrl);
    Logger.log('Detail error: ' + error.message);
    return null;
  }
}

function extractFormId_(formIdOrUrl) {
  const value = String(formIdOrUrl || '').trim();
  const match = value.match(/\/forms\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : value;
}

function syncFormResponse_(event, role) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId || !event || !event.response) {
    return;
  }

  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Rekap');
  const responseMap = {};

  event.response.getItemResponses().forEach(function(itemResponse) {
    const item = itemResponse.getItem();
    const title = item.getTitle();
    const match = title.match(/^\[(I\d{3})\]\s*(.+)$/);
    if (!match) {
      return;
    }

    const code = match[1];
    const field = normalizeResponseField_(match[2], item);
    if (!responseMap[code]) {
      responseMap[code] = {};
    }
    responseMap[code][field] = formatResponseValue_(itemResponse.getResponse(), item);
  });

  Object.keys(responseMap).forEach(function(code) {
    const row = findIndicatorRow_(sheet, code);
    if (!row) {
      return;
    }

    const values = responseMap[code];
    if (role === 'MANDIRI') {
      sheet.getRange(row, 6).setValue(values['Nilai Mandiri'] || '');
      sheet.getRange(row, 7).setValue(combineEvidenceValues_(values));
      sheet.getRange(row, 8).setValue(values['Keterangan'] || '');
    } else {
      sheet.getRange(row, 9).setValue(values['Nilai Tim Penilai'] || '');
      sheet.getRange(row, 10).setValue(values['Catatan Tim Penilai'] || '');
    }
  });
}

function combineEvidenceValues_(values) {
  const link = values['Lampirkan Link Dokumen'] || values['Bukti Dukung'] || '';
  const upload = values['Upload File'] || values['Upload Soft Copy'] || '';
  const parts = [];
  if (link) {
    parts.push('Link: ' + link);
  }
  if (upload) {
    parts.push('Upload: ' + upload);
  }
  return parts.join('\n');
}

function upsertMandiriSubmission_(event) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID') || EXISTING_REKAP_SPREADSHEET_ID;
  if (!spreadsheetId || !event || !event.response) {
    return;
  }

  const source = loadSourceData_();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const parsed = parseMandiriResponse_(event.response);
  const trainingName = parsed.identity['Nama Pelatihan'] || source.metadata.namaPelatihan || 'Penilaian Mandiri';
  const normalizedName = normalizeTrainingName_(trainingName);
  const trainingId = buildTrainingId_(normalizedName);
  const detailSheetName = buildSafeSheetName_(trainingName);
  const timestamp = event.response.getTimestamp ? event.response.getTimestamp() : new Date();
  const summary = summarizeMandiriSubmission_(source.indicators, parsed.indicators);

  const rekapSheet = ensureRekapSummarySheet_(spreadsheet);
  upsertRekapSummaryRow_(rekapSheet, {
    trainingId: trainingId,
    trainingName: trainingName,
    normalizedName: normalizedName,
    identity: parsed.identity,
    timestamp: timestamp,
    detailSheetName: detailSheetName,
    summary: summary
  });

  const detailSheet = ensureDetailJawabanSheet_(spreadsheet);
  replaceDetailRows_(detailSheet, {
    trainingId: trainingId,
    trainingName: trainingName,
    identity: parsed.identity,
    indicators: source.indicators,
    responseIndicators: parsed.indicators
  });

  const dashboardSheet = spreadsheet.getSheetByName('Dashboard') || spreadsheet.insertSheet('Dashboard');
  if (dashboardSheet.getLastRow() === 0) {
    writeDashboard_(dashboardSheet);
  }
}

function summarizeMandiriSubmission_(indicators, responseIndicators) {
  let filled = 0;
  let empty = 0;
  let notApplicable = 0;
  let evidenceLinks = 0;
  let evidenceUploads = 0;
  let notes = 0;

  indicators.forEach(function(item) {
    const submitted = responseIndicators[item.code] || {};
    const score = String(submitted['Nilai Mandiri'] || '').trim();
    const link = submitted['Lampirkan Link Dokumen'] || submitted['Bukti Dukung'] || '';
    const upload = submitted['Upload File'] || submitted['Upload Soft Copy'] || '';
    const note = submitted['Keterangan'] || '';

    if (score) {
      filled++;
    } else {
      empty++;
    }

    if (score === '-') {
      notApplicable++;
    }

    if (link) {
      evidenceLinks++;
    }
    if (upload) {
      evidenceUploads++;
    }
    if (note) {
      notes++;
    }
  });

  return {
    totalIndicators: indicators.length,
    filled: filled,
    empty: empty,
    notApplicable: notApplicable,
    evidenceLinks: evidenceLinks,
    evidenceUploads: evidenceUploads,
    notes: notes
  };
}

function categorizeAchievement_(achievement) {
  if (achievement === '') {
    return '';
  }
  if (achievement >= 0.9) {
    return 'Sangat Baik';
  }
  if (achievement >= 0.75) {
    return 'Baik';
  }
  if (achievement >= 0.6) {
    return 'Cukup';
  }
  return 'Perlu Tindak Lanjut';
}

function ensureRekapSummarySheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Rekap') || spreadsheet.insertSheet('Rekap');
  const expectedHeader = rekapSummaryHeader_();
  const currentHeader = sheet.getLastRow() ? sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), expectedHeader.length)).getValues()[0] : [];
  if (currentHeader.join('|') !== expectedHeader.join('|')) {
    sheet.clear();
    writeRekap_(sheet, {}, []);
  }
  return sheet;
}

function ensureDetailJawabanSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Detail Jawaban') || spreadsheet.insertSheet('Detail Jawaban');
  const expectedHeader = detailJawabanHeader_();
  const currentHeader = sheet.getLastRow() ? sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), expectedHeader.length)).getValues()[0] : [];
  if (currentHeader[0] !== expectedHeader[0] || currentHeader[2] !== expectedHeader[2]) {
    sheet.clear();
    writeDetailJawaban_(sheet);
  }
  return sheet;
}

function upsertRekapSummaryRow_(sheet, payload) {
  const header = rekapSummaryHeader_();
  const row = [
    payload.trainingId,
    payload.trainingName,
    payload.normalizedName,
    payload.identity['Metode'] || '',
    payload.identity['Tanggal Mulai Pelatihan'] || payload.identity['Tanggal Pelatihan'] || '',
    payload.identity['Tanggal Selesai Pelatihan'] || '',
    payload.identity['Tempat Penyelenggaraan'] || '',
    payload.identity['Unit Penyelenggara'] || '',
    payload.identity['Tanggal Kegiatan Penjaminan Mutu'] || '',
    'OK',
    'Belum Dinilai',
    payload.summary.totalIndicators,
    payload.summary.filled,
    payload.summary.empty,
    payload.summary.notApplicable,
    payload.summary.evidenceLinks,
    payload.summary.evidenceUploads,
    payload.summary.notes,
    payload.timestamp,
    payload.detailSheetName
  ];

  const targetRow = findRowByValue_(sheet, 1, payload.trainingId) || sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1, 1, header.length).setValues([row]);
  sheet.getRange(targetRow, 19).setNumberFormat('yyyy-mm-dd hh:mm');
}

function replaceDetailRows_(sheet, payload) {
  deleteRowsByValue_(sheet, 1, payload.trainingId);

  const rows = payload.indicators.map(function(item) {
    const submitted = payload.responseIndicators[item.code] || {};
    const nilaiMandiri = submitted['Nilai Mandiri'] || '';
    const nilaiAkhir = nilaiMandiri;
    return [
      payload.trainingId,
      payload.trainingName,
      item.code,
      item.no,
      item.dimension,
      item.indicator,
      item.scale,
      nilaiMandiri,
      submitted['Lampirkan Link Dokumen'] || submitted['Bukti Dukung'] || '',
      submitted['Upload File'] || submitted['Upload Soft Copy'] || '',
      submitted['Keterangan'] || '',
      '',
      '',
      nilaiAkhir,
      indicatorStatus_(nilaiAkhir)
    ];
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, detailJawabanHeader_().length).setValues(rows);
  }
}

function indicatorStatus_(value) {
  const normalized = String(value || '').trim();
  if (normalized === '1') {
    return 'Terpenuhi';
  }
  if (normalized === '0') {
    return 'Tidak Terpenuhi';
  }
  if (normalized === '-') {
    return 'Tidak Berlaku';
  }
  return 'Belum Diisi';
}

function normalizeTrainingName_(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/[^\w\s]/g, ' ')
    .replace(/\b(pelatihan|diklat|training|workshop|bimtek)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildTrainingId_(normalizedName) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, normalizedName || 'pelatihan');
  const hex = digest.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('').substring(0, 10).toUpperCase();
  return 'PEL-' + hex;
}

function findRowByValue_(sheet, column, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) {
      return i + 2;
    }
  }
  return null;
}

function deleteRowsByValue_(sheet, column, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === value) {
      sheet.deleteRow(i + 2);
    }
  }
}

function createTrainingAssessmentSheet_(event) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID') || EXISTING_REKAP_SPREADSHEET_ID;
  if (!spreadsheetId || !event || !event.response) {
    return null;
  }

  const source = loadSourceData_();
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const parsed = parseMandiriResponse_(event.response);
  const trainingName = parsed.identity['Nama Pelatihan'] || source.metadata.namaPelatihan || 'Penilaian Mandiri';
  const normalizedName = normalizeTrainingName_(trainingName);
  const trainingId = buildTrainingId_(normalizedName);
  const sheetName = buildSafeSheetName_(trainingName);
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  const reviewMap = existingSheet ? readTrainingReviewValues_(existingSheet) : {};
  const sheet = resetSheet_(spreadsheet, sheetName);

  writeTrainingAssessmentSheet_(sheet, parsed.identity, source.indicators, parsed.indicators, reviewMap);
  Logger.log('Sheet penilaian dibuat/diperbarui: ' + spreadsheet.getUrl() + '#gid=' + sheet.getSheetId());

  return {
    spreadsheet: spreadsheet,
    sheet: sheet,
    sheetName: sheetName,
    trainingId: trainingId,
    trainingName: trainingName,
    identity: parsed.identity,
    indicators: source.indicators,
    responseIndicators: parsed.indicators,
    reviewMap: reviewMap
  };
}

function parseMandiriResponse_(response) {
  const identity = {};
  const indicators = {};

  response.getItemResponses().forEach(function(itemResponse) {
    const item = itemResponse.getItem();
    const title = item.getTitle();
    const value = formatResponseValue_(itemResponse.getResponse(), item);
    const match = title.match(/^\[(I\d{3})\]\s*(.+)$/);

    if (match) {
      const code = match[1];
      const field = normalizeResponseField_(match[2], item);
      if (!indicators[code]) {
        indicators[code] = {};
      }
      indicators[code][field] = value;
    } else {
      identity[title] = value;
    }
  });

  return { identity: identity, indicators: indicators };
}

function normalizeIndicatorField_(field) {
  // MEMBANGUN MAP KARAKTER ITALIC SECARA KOMPREHENSIF (MEMPERBAIKI BUG)
  const italicMap = {};
  // Range 1: Mathematical Sans-Serif Italic (A-Z: 0x1D608 s/d 0x1D621, a-z: 0x1D622 s/d 0x1D63B)
  for (let i = 0; i < 26; i++) {
    italicMap[0x1d608 + i] = String.fromCharCode(65 + i);
    italicMap[0x1d622 + i] = String.fromCharCode(97 + i);
  }
  // Range 2: Mathematical Italic Serif Standar (A-Z: 0x1D434 s/d 0x1D44D, a-z: 0x1D44E s/d 0x1D467)
  for (let i = 0; i < 26; i++) {
    italicMap[0x1d434 + i] = String.fromCharCode(65 + i);
    italicMap[0x1d44e + i] = String.fromCharCode(97 + i);
  }

  const normalized = Array.from(String(field || '')).map(function(char) {
    return italicMap[char.codePointAt(0)] || char;
  }).join('');

  return normalized.replace(/^\[(I\d{3})\]\s*/, '').trim();
}

function normalizeResponseField_(field, item) {
  if (item && item.getType && item.getType() === FormApp.ItemType.FILE_UPLOAD) {
    return 'Upload File';
  }

  const normalized = normalizeIndicatorField_(field);
  if (normalized === 'Bukti Dukung') {
    return 'Lampirkan Link Dokumen';
  }
  return normalized;
}

function writeTrainingAssessmentSheet_(sheet, identity, indicators, responseIndicators, reviewMap) {
  reviewMap = reviewMap || {};
  const identityRows = [
    ['Nama Pelatihan', identity['Nama Pelatihan'] || ''],
    ['Angkatan', identity['Angkatan'] || ''],
    ['Metode', identity['Metode'] || ''],
    ['Tanggal Mulai Pelatihan', identity['Tanggal Mulai Pelatihan'] || ''],
    ['Tanggal Selesai Pelatihan', identity['Tanggal Selesai Pelatihan'] || ''],
    ['Tempat Penyelenggaraan', identity['Tempat Penyelenggaraan'] || ''],
    ['Unit Penyelenggara', identity['Unit Penyelenggara'] || ''],
    ['Tanggal Kegiatan Penjaminan Mutu', identity['Tanggal Kegiatan Penjaminan Mutu'] || ''],
    ['', '']
  ];

  sheet.getRange(1, 1, identityRows.length, 2).setValues(identityRows);
  sheet.getRange(1, 1, identityRows.length - 1, 1).setFontWeight('bold');
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold');

  const header = [
    'Kode Indikator',
    'No',
    'Dimensi',
    'Indikator',
    'Skala',
    'Nilai Mandiri',
      'Link Dokumen',
      'Upload File',
    'Keterangan Penyelenggara',
    'Nilai Tim Penilai',
    'Catatan Tim Penilai',
    'Status Penilaian',
    'Nilai Akhir'
  ];
  const startRow = 10;
  sheet.getRange(startRow, 1, 1, header.length).setValues([header]);

  const rows = indicators.map(function(item) {
    const submitted = responseIndicators[item.code] || {};
    const review = reviewMap[item.code] || {};
    return [
      item.code,
      item.no,
      item.dimension,
      item.indicator,
      item.scale,
      submitted['Nilai Mandiri'] || '',
      submitted['Lampirkan Link Dokumen'] || submitted['Bukti Dukung'] || '',
      submitted['Upload File'] || submitted['Upload Soft Copy'] || '',
      submitted['Keterangan'] || '',
      review.nilaiTimPenilai || '',
      review.catatanTimPenilai || '',
      review.statusPenilaian || 'Belum Dinilai',
      ''
    ];
  });

  sheet.getRange(startRow + 1, 1, rows.length, header.length).setValues(rows);
  sheet.getRange(startRow + 1, 10, rows.length, 3).setBackground('#fff9e6');
  applyDropdownValidation_(sheet.getRange(startRow + 1, 12, rows.length, 1), statusPenilaianOptions_());
  applyStatusPenilaianColors_(sheet.getRange(startRow + 1, 12, rows.length, 1));
  sheet.getRange(startRow + 1, 13, rows.length, 1).setFormulaR1C1('=IF(RC[-3]<>"",RC[-3],RC[-7])');

  const totalRow = startRow + rows.length + 2;
  sheet.getRange(totalRow, 1, 3, 2).setValues([
    ['Jumlah Nilai Akhir', '=SUM(M' + (startRow + 1) + ':M' + (startRow + rows.length) + ')'],
    ['Jumlah Indikator Dinilai', '=COUNT(M' + (startRow + 1) + ':M' + (startRow + rows.length) + ')'],
    ['Nilai Mutu', '=IF(B' + (totalRow + 1) + '=0,"",B' + totalRow + '/B' + (totalRow + 1) + '*100)']
  ]);
  sheet.getRange(totalRow, 1, 3, 1).setFontWeight('bold');

  formatHeader_(sheet, startRow, header.length);
  sheet.setFrozenRows(startRow);
  sheet.autoResizeColumns(1, header.length);
}

function statusPenilaianOptions_() {
  return [
    'Belum Dinilai',
    'Sudah Sesuai',
    'Perlu Tindak Lanjut'
  ];
}

function applyDropdownValidation_(range, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function applyStatusPenilaianColors_(range) {
  const sheet = range.getSheet();
  const existingRules = sheet.getConditionalFormatRules().filter(function(rule) {
    return !rule.getRanges().some(function(ruleRange) {
      return ruleRange.getA1Notation() === range.getA1Notation();
    });
  });

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Belum Dinilai')
      .setBackground('#fff2cc')
      .setFontColor('#7f6000')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Perlu Tindak Lanjut')
      .setBackground('#fce4d6')
      .setFontColor('#c65911')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sudah Sesuai')
      .setBackground('#d9ead3')
      .setFontColor('#274e13')
      .setRanges([range])
      .build()
  ];

  sheet.setConditionalFormatRules(existingRules.concat(rules));
}

function readTrainingReviewValues_(sheet) {
  const map = {};
  const headerRow = 10;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= headerRow || lastColumn === 0) {
    return map;
  }

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const codeCol = headers.indexOf('Kode Indikator') + 1;
  const nilaiCol = headers.indexOf('Nilai Tim Penilai') + 1;
  const catatanCol = headers.indexOf('Catatan Tim Penilai') + 1;
  const statusCol = headers.indexOf('Status Penilaian') + 1;
  if (!codeCol) {
    return map;
  }

  const values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastColumn).getValues();
  values.forEach(function(row) {
    const code = row[codeCol - 1];
    if (!code) {
      return;
    }
    map[code] = {
      nilaiTimPenilai: nilaiCol ? row[nilaiCol - 1] : '',
      catatanTimPenilai: catatanCol ? row[catatanCol - 1] : '',
      statusPenilaian: statusCol ? row[statusCol - 1] : ''
    };
  });

  return map;
}

function buildAssessmentContextFromSheet_(sheet) {
  const source = loadSourceData_();
  const identity = {
    'Nama Pelatihan': sheet.getRange('B1').getValue(),
    'Angkatan': sheet.getRange('B2').getValue(),
    'Metode': sheet.getRange('B3').getValue(),
    'Tanggal Mulai Pelatihan': sheet.getRange('B4').getValue(),
    'Tanggal Selesai Pelatihan': sheet.getRange('B5').getValue(),
    'Tempat Penyelenggaraan': sheet.getRange('B6').getValue(),
    'Unit Penyelenggara': sheet.getRange('B7').getValue(),
    'Tanggal Kegiatan Penjaminan Mutu': sheet.getRange('B8').getValue()
  };
  const trainingName = identity['Nama Pelatihan'] || sheet.getName();
  const responseIndicators = {};
  const headerRow = 10;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= headerRow) {
    return null;
  }

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const codeCol = headers.indexOf('Kode Indikator') + 1;
  const nilaiMandiriCol = headers.indexOf('Nilai Mandiri') + 1;
  const linkCol = headers.indexOf('Link Dokumen') + 1;
  const uploadCol = headers.indexOf('Upload File') + 1;
  const ketCol = headers.indexOf('Keterangan Penyelenggara') + 1;
  if (!codeCol) {
    return null;
  }

  const values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastColumn).getValues();
  values.forEach(function(row) {
    const code = row[codeCol - 1];
    if (!code) {
      return;
    }
    responseIndicators[code] = {
      'Nilai Mandiri': nilaiMandiriCol ? row[nilaiMandiriCol - 1] : '',
      'Lampirkan Link Dokumen': linkCol ? row[linkCol - 1] : '',
      'Upload File': uploadCol ? row[uploadCol - 1] : '',
      'Keterangan': ketCol ? row[ketCol - 1] : ''
    };
  });

  const normalizedName = normalizeTrainingName_(trainingName);
  return {
    sheet: sheet,
    sheetName: sheet.getName(),
    trainingId: buildTrainingId_(normalizedName),
    trainingName: trainingName,
    identity: identity,
    indicators: source.indicators,
    responseIndicators: responseIndicators,
    reviewMap: readTrainingReviewValues_(sheet)
  };
}

function syncOrganizerProgressSheet_(event, assessmentContext) {
  if (!event || !event.response || !assessmentContext) {
    return;
  }

  const email = getOrganizerEmail_(event.response, assessmentContext.identity);
  syncOrganizerProgressContext_(assessmentContext, email, true);
}

function syncOrganizerProgressContext_(assessmentContext, email, shouldSendEmail) {
  const organizerSpreadsheet = getOrCreateOrganizerSpreadsheet_(assessmentContext.trainingId, assessmentContext.trainingName, email);
  const progressSheet = writeOrganizerProgressWorkbook_(organizerSpreadsheet, assessmentContext);

  if (shouldSendEmail && email) {
    sendOrganizerProgressEmail_(email, assessmentContext.trainingName, buildSheetTabUrl_(organizerSpreadsheet, progressSheet));
  } else if (shouldSendEmail) {
    Logger.log('Email penyelenggara tidak tersedia; link sheet progres tidak dikirim. URL: ' + buildSheetTabUrl_(organizerSpreadsheet, progressSheet));
  }
}

function getOrCreateOrganizerSpreadsheet_(trainingId, trainingName, email) {
  const properties = PropertiesService.getScriptProperties();
  const propertyKey = 'ORGANIZER_PROGRESS_SPREADSHEET_ID_' + trainingId;
  const existingId = properties.getProperty(propertyKey);

  if (existingId) {
    try {
      const spreadsheet = SpreadsheetApp.openById(existingId);
      grantOrganizerAccess_(spreadsheet.getId(), email);
      return spreadsheet;
    } catch (error) {
      Logger.log('Sheet progres penyelenggara lama tidak dapat dibuka, membuat baru: ' + error.message);
    }
  }

  const spreadsheet = SpreadsheetApp.create('Progress Penilaian Penyelenggara - ' + trainingName);
  properties.setProperty(propertyKey, spreadsheet.getId());
  grantOrganizerAccess_(spreadsheet.getId(), email);

  return spreadsheet;
}

function grantOrganizerAccess_(spreadsheetId, email) {
  if (!email) {
    return;
  }

  try {
    DriveApp.getFileById(spreadsheetId).addViewer(email);
  } catch (error) {
    Logger.log('Gagal memberi akses viewer ke ' + email + ': ' + error.message);
  }
}

function writeOrganizerProgressWorkbook_(spreadsheet, context) {
  const summarySheet = resetSheet_(spreadsheet, 'Ringkasan');
  const progressSheet = resetSheet_(spreadsheet, 'Progress Indikator');
  removeDefaultBlankSheet_(spreadsheet);
  const timestamp = new Date();
  const progress = summarizeOrganizerProgress_(context.indicators, context.responseIndicators, context.reviewMap);

  const summaryRows = [
    ['Progress Penilaian Penyelenggara'],
    [''],
    ['ID Pelatihan', context.trainingId],
    ['Nama Pelatihan', context.trainingName],
    ['Metode', context.identity['Metode'] || ''],
    ['Tanggal Mulai', context.identity['Tanggal Mulai Pelatihan'] || context.identity['Tanggal Pelatihan'] || ''],
    ['Tanggal Selesai', context.identity['Tanggal Selesai Pelatihan'] || ''],
    ['Tempat', context.identity['Tempat Penyelenggaraan'] || ''],
    ['Unit Penyelenggara', context.identity['Unit Penyelenggara'] || ''],
    ['Terakhir Diperbarui', timestamp],
    [''],
    ['Jumlah Indikator', progress.total],
    ['Sudah Dinilai', progress.done],
    ['Belum Dinilai', progress.pending],
    ['Perlu Tindak Lanjut', progress.needsFollowUp]
  ];

  summarySheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows.map(function(row) {
    return row.length === 1 ? [row[0], ''] : row;
  }));
  summarySheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  summarySheet.getRange(10, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  summarySheet.autoResizeColumns(1, 2);

  const header = [
    'Kode Indikator',
    'No',
    'Dimensi',
    'Indikator',
    'Nilai Mandiri',
    'Ada Link Dokumen',
    'Ada File Upload',
    'Keterangan Penyelenggara',
    'Status Penilaian',
    'Catatan Tim Penilai',
    'Terakhir Diperbarui'
  ];
  progressSheet.getRange(1, 1, 1, header.length).setValues([header]);

  const rows = context.indicators.map(function(item) {
    const submitted = context.responseIndicators[item.code] || {};
    const review = context.reviewMap[item.code] || {};
    const link = submitted['Lampirkan Link Dokumen'] || submitted['Bukti Dukung'] || '';
    const upload = submitted['Upload File'] || submitted['Upload Soft Copy'] || '';
    return [
      item.code,
      item.no,
      item.dimension,
      item.indicator,
      submitted['Nilai Mandiri'] || '',
      link ? 'Ada' : 'Kosong',
      upload ? 'Ada' : 'Kosong',
      submitted['Keterangan'] || '',
      review.statusPenilaian || 'Belum Dinilai',
      review.catatanTimPenilai || '',
      timestamp
    ];
  });

  if (rows.length) {
    progressSheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    applyDropdownValidation_(progressSheet.getRange(2, 9, rows.length, 1), statusPenilaianOptions_());
    applyStatusPenilaianColors_(progressSheet.getRange(2, 9, rows.length, 1));
    progressSheet.getRange(2, 11, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  formatHeader_(progressSheet, 1, header.length);
  progressSheet.setFrozenRows(1);
  progressSheet.autoResizeColumns(1, header.length);
  spreadsheet.setActiveSheet(progressSheet);
  return progressSheet;
}

function buildSheetTabUrl_(spreadsheet, sheet) {
  return spreadsheet.getUrl() + '#gid=' + sheet.getSheetId();
}

function removeDefaultBlankSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Sheet1');
  if (!sheet || spreadsheet.getSheets().length <= 1) {
    return;
  }

  if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && sheet.getLastColumn() === 1 && sheet.getRange(1, 1).getValue() === '')) {
    spreadsheet.deleteSheet(sheet);
  }
}

function summarizeOrganizerProgress_(indicators, responseIndicators, reviewMap) {
  let done = 0;
  let needsFollowUp = 0;

  indicators.forEach(function(item) {
    const review = reviewMap[item.code] || {};
    const status = String(review.statusPenilaian || '').toLowerCase();
    if (status && status !== 'belum dinilai') {
      done++;
    }
    if (status === 'perlu tindak lanjut') {
      needsFollowUp++;
    }
  });

  return {
    total: indicators.length,
    done: done,
    pending: indicators.length - done,
    needsFollowUp: needsFollowUp
  };
}

function getOrganizerEmail_(response, identity) {
  if (response && response.getRespondentEmail) {
    const respondentEmail = response.getRespondentEmail();
    if (respondentEmail) {
      return respondentEmail;
    }
  }

  const possibleKeys = ['Email', 'Alamat Email', 'Alamat email', 'Email Penyelenggara'];
  for (let i = 0; i < possibleKeys.length; i++) {
    const value = identity[possibleKeys[i]];
    if (value && String(value).indexOf('@') !== -1) {
      return String(value).trim();
    }
  }

  return '';
}

function sendOrganizerProgressEmail_(email, trainingName, spreadsheetUrl) {
  const subject = 'Link Progress Penilaian - ' + trainingName;
  const body = [
    'Yth. Penyelenggara,',
    '',
    'Berikut link Google Sheet khusus penyelenggara untuk memantau progres penilaian indikator:',
    spreadsheetUrl,
    '',
    'Sheet ini diperbarui otomatis setiap kali respons Form Penilaian Mandiri dikirim atau diedit.',
    '',
    'Salam,',
    'Sistem Penjaminan Mutu Pelatihan'
  ].join('\n');

  MailApp.sendEmail(email, subject, body);
}

function buildSafeSheetName_(name) {
  const cleaned = String(name || 'Penilaian Mandiri')
    .replace(/[\\/?*[\]:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 90);
  return cleaned || 'Penilaian Mandiri';
}

function loadSourceData_() {
  const csvUrl = 'https://docs.google.com/spreadsheets/d/' + SOURCE_SPREADSHEET_ID + '/export?format=csv';
  const csv = UrlFetchApp.fetch(csvUrl).getContentText();
  const rows = Utilities.parseCsv(csv);

  const metadata = {
    namaPelatihan: valueAt_(rows, 3, 2),
    angkatan: valueAt_(rows, 4, 2),
    metode: valueAt_(rows, 5, 2),
    tanggalPelatihan: valueAt_(rows, 6, 2),
    tempat: valueAt_(rows, 7, 2),
    unit: valueAt_(rows, 8, 2),
    tanggalPm: valueAt_(rows, 9, 2),
  };

  const indicators = [];
  let currentDimension = '';
  let fallbackNo = 1;

  for (let i = 12; i < rows.length; i++) {
    const row = rows[i];
    const indicator = clean_(row[2]);
    const method = clean_(row[3]);
    const scale = clean_(row[4]);

    if (!indicator || indicator.toLowerCase() === 'jumlah') {
      continue;
    }

    const dimension = clean_(row[1]);
    if (dimension) {
      currentDimension = dimension;
    }

    const no = clean_(row[0]) || String(fallbackNo);
    indicators.push({
      code: Utilities.formatString('I%03d', fallbackNo),
      no: no,
      dimension: currentDimension,
      indicator: indicator,
      method: method,
      scale: scale || '1. ada / 0. tidak',
    });
    fallbackNo++;
  }

  return { metadata: metadata, indicators: indicators };
}

function buildMandiriForm_(source, spreadsheet) {
  const form = FormApp.create('Form Penilaian Mandiri Penyelenggara');
  form.setDescription(
    'Form ini diisi oleh Penyelenggara sebagai penilaian mandiri sebelum proses verifikasi mutu.\n\n' +
    'Isi nilai, bukti dukung, dan keterangan untuk setiap indikator. Nilai menggunakan angka 0 sampai 1. ' +
    'Jika bukti tidak tersedia atau indikator tidak berlaku, isi alasannya pada Keterangan.'
  );
  form.setCollectEmail(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  addIdentitySection_(form, source.metadata);
  addMandiriIndicatorItems_(form, source.indicators);

  return form;
}

function buildVerifikasiForm_(source, spreadsheet) {
  const form = FormApp.create('Form Verifikasi Mutu');
  form.setDescription(
    'Form ini diisi oleh Verifikator setelah memeriksa hasil penilaian mandiri, bukti dukung, dan keterangan Penyelenggara pada Google Sheet Rekap: ' +
    spreadsheet.getUrl()
  );
  form.setCollectEmail(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());

  addIdentitySection_(form, source.metadata);
  form.addSectionHeaderItem()
    .setTitle('Acuan Verifikasi')
    .setHelpText('Sebelum mengisi, buka tab Rekap pada Google Sheet untuk melihat nilai mandiri, bukti dukung, dan keterangan Penyelenggara.');
  addVerifikasiIndicatorItems_(form, source.indicators);

  return form;
}

function installSubmitTriggers_(spreadsheet, mandiriForm) {
  PropertiesService.getScriptProperties().setProperties({
    REKAP_SPREADSHEET_ID: spreadsheet.getId(),
    MANDIRI_FORM_ID: mandiriForm.getId()
  });

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === 'handleMandiriSubmit') {
      ScriptApp.deleteTrigger(trigger);
    } else if (handler === 'handlePenilaianSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleMandiriSubmit')
    .forForm(mandiriForm)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('handlePenilaianSheetEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}

function installPenilaianEditTriggerForExistingSpreadsheet() {
  const spreadsheetId = EXISTING_REKAP_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Isi EXISTING_REKAP_SPREADSHEET_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'handlePenilaianSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handlePenilaianSheetEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  Logger.log('Trigger edit status penilaian terpasang untuk spreadsheet: ' + spreadsheet.getUrl());
}

function applyStatusPenilaianDropdownsToExistingSheets() {
  const spreadsheetId = EXISTING_REKAP_SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('REKAP_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Isi EXISTING_REKAP_SPREADSHEET_ID atau jalankan setupInstrumenPenjaminanMutu terlebih dahulu.');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let count = 0;
  spreadsheet.getSheets().forEach(function(sheet) {
    if (applyStatusDropdownToAssessmentSheet_(sheet)) {
      count++;
    }
  });

  Logger.log('Dropdown Status Penilaian diterapkan pada ' + count + ' tab penilaian.');
}

function applyStatusDropdownToAssessmentSheet_(sheet) {
  const headerRow = 10;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= headerRow || lastColumn === 0) {
    return false;
  }

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
  const statusColumn = headers.indexOf('Status Penilaian') + 1;
  if (!statusColumn) {
    return false;
  }

  const statusRange = sheet.getRange(headerRow + 1, statusColumn, lastRow - headerRow, 1);
  applyDropdownValidation_(statusRange, statusPenilaianOptions_());
  applyStatusPenilaianColors_(statusRange);
  return true;
}

function addIdentitySection_(form, metadata) {
  form.addSectionHeaderItem()
    .setTitle('Identitas Pelatihan')
    .setHelpText(identitySectionHelpText_());
  form.addTextItem().setTitle('Nama Pelatihan').setRequired(true).setHelpText(identityHelpText_('Nama Pelatihan', metadata.namaPelatihan));
  form.addTextItem().setTitle('Angkatan').setRequired(true).setHelpText(identityHelpText_('Angkatan', metadata.angkatan));
  form.addMultipleChoiceItem()
    .setTitle('Metode')
    .setChoiceValues(['1 - Daring', '2 - Luring', '3 - Blended'])
    .setRequired(true)
    .setHelpText(identityHelpText_('Metode', metadata.metode));
  form.addDateItem()
    .setTitle('Tanggal Mulai Pelatihan')
    .setRequired(true)
    .setHelpText('Pilih tanggal hari pertama pelaksanaan pelatihan. Contoh sumber: ' + metadata.tanggalPelatihan);
  form.addDateItem()
    .setTitle('Tanggal Selesai Pelatihan')
    .setRequired(true)
    .setHelpText('Pilih tanggal hari terakhir pelaksanaan pelatihan. Contoh sumber: ' + metadata.tanggalPelatihan);
  form.addTextItem().setTitle('Tempat Penyelenggaraan').setRequired(true).setHelpText(identityHelpText_('Tempat Penyelenggaraan', metadata.tempat));
  form.addTextItem().setTitle('Unit Penyelenggara').setRequired(true).setHelpText(identityHelpText_('Unit Penyelenggara', metadata.unit));
  form.addTextItem().setTitle('Tanggal Kegiatan Penjaminan Mutu').setRequired(true).setHelpText(identityHelpText_('Tanggal Kegiatan Penjaminan Mutu', metadata.tanggalPm));
}

function addMandiriIndicatorItems_(form, indicators) {
  let currentDimension = '';
  indicators.forEach(function(item) {
    if (item.dimension !== currentDimension) {
      currentDimension = item.dimension;
      form.addPageBreakItem().setTitle(currentDimension);
    }

    const title = itemTitle_(item);
    form.addSectionHeaderItem()
      .setTitle(title)
      .setHelpText(indicatorHelpText_(item, 'MANDIRI'));

    form.addTextItem()
      .setTitle('[' + item.code + '] 𝘕𝘪𝘭𝘢𝘪 𝘔𝘢𝘯𝘥𝘪𝘳𝘪')
      .setHelpText(scoreHelpText_('MANDIRI'))
      .setRequired(true);

    form.addTextItem()
      .setTitle('[' + item.code + '] 𝘓𝘢𝘮𝘱𝘪𝘳𝘬𝘢𝘯 𝘓𝘪𝘯𝘬 𝘋𝘰𝘬𝘶𝘮𝘦𝘯')
      .setHelpText(evidenceLinkHelpText_())
      .setRequired(false);

    form.addFileUploadItem()
      .setTitle('[' + item.code + '] 𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦')
      .setHelpText(evidenceUploadHelpText_())
      .setRequired(false);

    form.addParagraphTextItem()
      .setTitle('[' + item.code + '] 𝘒𝘦𝘵𝘦𝘳𝘢𝘯𝘨𝘢𝘯')
      .setHelpText(noteHelpText_('MANDIRI'))
      .setRequired(false);
  });
}

function addVerifikasiIndicatorItems_(form, indicators) {
  let currentDimension = '';
  indicators.forEach(function(item) {
    if (item.dimension !== currentDimension) {
      currentDimension = item.dimension;
      form.addPageBreakItem().setTitle(currentDimension);
    }

    form.addSectionHeaderItem()
      .setTitle(itemTitle_(item))
      .setHelpText(indicatorHelpText_(item, 'VERIFIKASI'));

    form.addTextItem()
      .setTitle('[' + item.code + '] Nilai Verifikator')
      .setHelpText(scoreHelpText_('VERIFIKASI'))
      .setRequired(true);

    form.addParagraphTextItem()
      .setTitle('[' + item.code + '] Catatan Verifikator')
      .setHelpText(noteHelpText_('VERIFIKASI'))
      .setRequired(false);
  });
}

function polishFormHelpText_(form, source, role) {
  const indicatorMap = {};
  source.indicators.forEach(function(item) {
    indicatorMap[item.code] = item;
  });

  if (role === 'MANDIRI') {
    form.setDescription(
      'Form ini diisi oleh Penyelenggara sebagai penilaian mandiri sebelum proses verifikasi mutu.\n\n' +
      'Isi nilai, bukti dukung, dan keterangan untuk setiap indikator. Nilai menggunakan angka 0 sampai 1. ' +
      'Jika bukti tidak tersedia atau indikator tidak berlaku, isi alasannya pada Keterangan.'
    );
  } else {
    form.setDescription(
      'Form ini diisi oleh Verifikator setelah memeriksa hasil penilaian mandiri, bukti dukung, dan keterangan Penyelenggara.\n\n' +
      'Isi nilai verifikasi dan catatan bila ada koreksi, klarifikasi, atau bukti belum memadai.'
    );
  }

  form.getItems().forEach(function(item) {
    const title = item.getTitle();
    const codeMatch = title.match(/^\[(I\d{3})\]/);
    const indicator = codeMatch ? indicatorMap[codeMatch[1]] : null;

    if (title === 'Identitas Pelatihan' && item.getType() === FormApp.ItemType.SECTION_HEADER) {
      item.asSectionHeaderItem().setHelpText(identitySectionHelpText_());
    } else if (title === 'Nama Pelatihan') {
      setItemHelpText_(item, identityHelpText_('Nama Pelatihan', source.metadata.namaPelatihan));
    } else if (title === 'Angkatan') {
      setItemHelpText_(item, identityHelpText_('Angkatan', source.metadata.angkatan));
    } else if (title === 'Metode') {
      setItemHelpText_(item, identityHelpText_('Metode', source.metadata.metode));
    } else if (title === 'Tanggal Pelatihan') {
      setItemHelpText_(item, identityHelpText_('Tanggal Pelatihan', source.metadata.tanggalPelatihan));
    } else if (title === 'Tempat Penyelenggaraan') {
      setItemHelpText_(item, identityHelpText_('Tempat Penyelenggaraan', source.metadata.tempat));
    } else if (title === 'Unit Penyelenggara') {
      setItemHelpText_(item, identityHelpText_('Unit Penyelenggara', source.metadata.unit));
    } else if (title === 'Tanggal Kegiatan Penjaminan Mutu') {
      setItemHelpText_(item, identityHelpText_('Tanggal Kegiatan Penjaminan Mutu', source.metadata.tanggalPm));
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Nilai Mandiri') !== -1) {
      setItemHelpText_(item, scoreHelpText_('MANDIRI'));
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Lampirkan Link Dokumen') !== -1) {
      setItemHelpText_(item, evidenceLinkHelpText_());
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Upload File') !== -1) {
      setItemHelpText_(item, evidenceUploadHelpText_());
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Bukti Dukung') !== -1) {
      setItemHelpText_(item, evidenceLinkHelpText_());
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Keterangan') !== -1) {
      setItemHelpText_(item, noteHelpText_('MANDIRI'));
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Nilai Verifikator') !== -1) {
      setItemHelpText_(item, scoreHelpText_('VERIFIKASI'));
    } else if (indicator && normalizeIndicatorField_(title).indexOf('Catatan Verifikator') !== -1) {
      setItemHelpText_(item, noteHelpText_('VERIFIKASI'));
    } else if (indicator && item.getType() === FormApp.ItemType.SECTION_HEADER) {
      item.asSectionHeaderItem().setHelpText(indicatorHelpText_(indicator, role));
    }
  });
}

function setItemHelpText_(item, helpText) {
  const type = item.getType();
  if (type === FormApp.ItemType.TEXT) {
    item.asTextItem().setHelpText(helpText);
  } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
    item.asParagraphTextItem().setHelpText(helpText);
  } else if (type === FormApp.ItemType.SECTION_HEADER) {
    item.asSectionHeaderItem().setHelpText(helpText);
  } else if (type === FormApp.ItemType.FILE_UPLOAD) {
    item.asFileUploadItem().setHelpText(helpText);
  } else if (type === FormApp.ItemType.PAGE_BREAK) {
    item.asPageBreakItem().setHelpText(helpText);
  }
}

function setItemTitle_(item, title) {
  const type = item.getType();
  if (type === FormApp.ItemType.TEXT) {
    item.asTextItem().setTitle(title);
  } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
    item.asParagraphTextItem().setTitle(title);
  } else if (type === FormApp.ItemType.FILE_UPLOAD) {
    item.asFileUploadItem().setTitle(title);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setTitle(title);
  } else if (type === FormApp.ItemType.DATE) {
    item.asDateItem().setTitle(title);
  } else if (type === FormApp.ItemType.SECTION_HEADER) {
    item.asSectionHeaderItem().setTitle(title);
  }
}

function upgradeIdentityQuestionTypes_(form) {
  replaceTextWithMultipleChoice_(form, 'Metode', [
    '1 - Daring',
    '2 - Luring',
    '3 - Blended'
  ], 'Pilih metode penyelenggaraan pelatihan.');

  replaceTextWithDateRange_(form, 'Tanggal Pelatihan');
}

function replaceTextWithMultipleChoice_(form, title, choices, helpText) {
  const item = findItemByTitle_(form, title);
  if (!item || item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
    return;
  }

  const index = item.getIndex();
  const required = isRequiredQuestion_(item);
  const newItem = form.addMultipleChoiceItem()
    .setTitle(title)
    .setChoiceValues(choices)
    .setHelpText(helpText)
    .setRequired(required);
  form.moveItem(newItem.getIndex(), index);
  form.deleteItem(item);
}

function replaceTextWithDateRange_(form, title) {
  const item = findItemByTitle_(form, title);
  if (!item) {
    return;
  }

  const index = item.getIndex();
  const required = isRequiredQuestion_(item);
  const startItem = form.addDateItem()
    .setTitle('Tanggal Mulai Pelatihan')
    .setHelpText('Pilih tanggal hari pertama pelaksanaan pelatihan.')
    .setRequired(required);
  form.moveItem(startItem.getIndex(), index);

  const endItem = form.addDateItem()
    .setTitle('Tanggal Selesai Pelatihan')
    .setHelpText('Pilih tanggal hari terakhir pelaksanaan pelatihan.')
    .setRequired(required);
  form.moveItem(endItem.getIndex(), index + 1);

  form.deleteItem(item);
}

function findItemByTitle_(form, title) {
  const items = form.getItems();
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle && items[i].getTitle() === title) {
      return items[i];
    }
  }
  return null;
}

function isRequiredQuestion_(item) {
  const type = item.getType();
  if (type === FormApp.ItemType.TEXT) {
    return item.asTextItem().isRequired();
  }
  if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
    return item.asParagraphTextItem().isRequired();
  }
  if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    return item.asMultipleChoiceItem().isRequired();
  }
  if (type === FormApp.ItemType.DATE) {
    return item.asDateItem().isRequired();
  }
  return false;
}

function identityHelpText_(field, value) {
  const examples = {
    'Nama Pelatihan': ['𝗖𝗼𝗻𝘁𝗼𝗵', value, '', '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬', '- Gunakan nama pelatihan resmi sesuai dokumen penyelenggaraan.'].join('\n'),
    'Angkatan': ['𝗖𝗼𝗻𝘁𝗼𝗵', value, '', '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬', '- Isi nomor angkatan atau gelombang pelatihan.'].join('\n'),
    'Metode': ['𝗣𝗶𝗹𝗶𝗵𝗮𝗻 𝗠𝗲𝘁𝗼𝗱𝗲', '- 1 - Daring', '- 2 - Luring', '- 3 - Blended', '', '𝘊𝘰𝘯𝘵𝘰𝘩 𝘥𝘢𝘳𝘪 𝘴𝘶𝘮𝘣𝘦𝘳', value].join('\n'),
    'Tanggal Pelatihan': ['𝗖𝗼𝗻𝘁𝗼𝗵', value, '', '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬', '- Gunakan rentang tanggal pelaksanaan pelatihan.'].join('\n'),
    'Tempat Penyelenggaraan': ['𝗖𝗼𝗻𝘁𝗼𝗵', value, '', '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬', '- Isi lokasi utama penyelenggaraan pelatihan.'].join('\n'),
    'Unit Penyelenggara': ['𝗖𝗼𝗻𝘁𝗼𝗵', value, '', '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬', '- Isi unit kerja atau tim yang bertanggung jawab.'].join('\n'),
    'Tanggal Kegiatan Penjaminan Mutu': ['𝗖𝗼𝗻𝘁𝗼𝗵', value, '', '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬', '- Isi tanggal pelaksanaan kegiatan penjaminan mutu.'].join('\n')
  };
  return examples[field] || ('Contoh: ' + value);
}

function identitySectionHelpText_() {
  return [
    '𝗜𝗗𝗘𝗡𝗧𝗜𝗧𝗔𝗦 𝗣𝗘𝗟𝗔𝗧𝗜𝗛𝗔𝗡',
    'Bagian ini diisi untuk memastikan hasil penilaian tersambung ke pelatihan yang tepat.',
    '',
    '𝘗𝘦𝘵𝘶𝘯𝘫𝘶𝘬',
    '- Isi sesuai dokumen resmi penyelenggaraan.',
    '- Gunakan format yang konsisten agar tab hasil penilaian mudah ditemukan.',
    '- Nama Pelatihan akan dipakai sebagai nama sheet hasil penilaian.'
  ].join('\n');
}

function indicatorHelpText_(item, role) {
  const prefix = role === 'MANDIRI'
    ? 'Baca indikator berikut, lalu isi nilai mandiri berdasarkan bukti yang tersedia.'
    : 'Periksa hasil penilaian mandiri, bukti dukung, dan keterangan Penyelenggara sebelum memberi nilai verifikasi.';

  return [
    '𝗥𝗜𝗡𝗚𝗞𝗔𝗦𝗔𝗡 𝗜𝗡𝗗𝗜𝗞𝗔𝗧𝗢𝗥',
    prefix,
    '',
    '𝗗𝗶𝗺𝗲𝗻𝘀𝗶',
    '- ' + item.dimension,
    '',
    '𝗜𝗻𝗱𝗶𝗸𝗮𝘁𝗼𝗿',
    '- ' + item.indicator,
    '',
    '𝗦𝗸𝗮𝗹𝗮 𝗽𝗲𝗻𝗶𝗹𝗮𝗶𝗮𝗻',
    '- ' + item.scale,
    '',
    '𝗖𝗼𝗻𝘁𝗼𝗵 𝗯𝘂𝗸𝘁𝗶 𝘆𝗮𝗻𝗴 𝗿𝗲𝗹𝗲𝘃𝗮𝗻',
    '- Link dokumen, gojags, kalender, evaluasi, file soft copy, atau keterangan hasil pengamatan lapangan.'
  ].join('\n');
}

function scoreHelpText_(role) {
  const label = role === 'MANDIRI' ? 'mandiri' : 'verifikasi';
  return [
    '𝗖𝗔𝗥𝗔 𝗠𝗘𝗡𝗚𝗜𝗦𝗜 𝗡𝗜𝗟𝗔𝗜',
    'Isi nilai ' + label + ' dengan angka 0 sampai 1.',
    '',
    '𝗣𝗶𝗹𝗶𝗵𝗮𝗻 𝗻𝗶𝗹𝗮𝗶',
    '- 1 = bukti tersedia dan indikator terpenuhi.',
    '- 0,5 = bukti tersedia sebagian atau indikator terpenuhi sebagian.',
    '- 0 = bukti tidak tersedia atau indikator tidak terpenuhi.',
    '- - = indikator tidak berlaku.',
    '',
    '𝘊𝘢𝘵𝘢𝘵𝘢𝘯',
    '- Jika memakai nilai kurang dari 1 atau tanda -, jelaskan alasannya pada keterangan/catatan.'
  ].join('\n');
}

function evidenceLinkHelpText_() {
  return [
    '𝗟𝗜𝗡𝗞 𝗗𝗢𝗞𝗨𝗠𝗘𝗡',
    'Isi tautan dokumen atau folder bukti dukung yang relevan dengan indikator.',
    '',
    '𝗖𝗼𝗻𝘁𝗼𝗵',
    '- Link Google Drive, gojags, kalender, evaluasi, LMS, atau folder dokumen.',
    '',
    '𝘊𝘢𝘵𝘢𝘵𝘢𝘯',
    '- Pastikan link dapat diakses oleh Tim Penilai.',
    '- Jika tidak ada link, gunakan upload file atau jelaskan alasannya pada Keterangan.'
  ].join('\n');
}

function evidenceUploadHelpText_() {
  return [
    '𝗨𝗣𝗟𝗢𝗔𝗗 𝗙𝗜𝗟𝗘',
    'Unggah file bukti dukung yang relevan dengan indikator.',
    '',
    '𝗖𝗼𝗻𝘁𝗼𝗵',
    '- PDF, Word, Excel, gambar, atau dokumen pendukung lainnya.',
    '',
    '𝘊𝘢𝘵𝘢𝘵𝘢𝘯',
    '- Gunakan jika bukti belum tersedia dalam bentuk link.',
    '- Jika tidak dapat mengunggah file, lampirkan link dokumen atau jelaskan alasannya pada Keterangan.'
  ].join('\n');
}

function evidenceUploadFallbackHelpText_() {
  return [
    '𝗨𝗣𝗟𝗢𝗔𝗗 𝗙𝗜𝗟𝗘',
    'Fitur upload file native tidak tersedia melalui Apps Script pada akun/project ini.',
    '',
    '𝗖𝗮𝗿𝗮 𝗺𝗲𝗻𝗴𝗶𝘀𝗶',
    '- Upload file ke Google Drive atau media penyimpanan resmi.',
    '- Pastikan akses file dapat dibuka oleh Tim Penilai.',
    '- Tempel link file yang sudah di-upload pada isian ini.',
    '',
    '𝘊𝘰𝘯𝘵𝘰𝘩',
    '- Link PDF, Word, Excel, gambar, atau folder bukti dukung.'
  ].join('\n');
}

function evidenceExample_(method) {
  const normalized = String(method || '').toLowerCase();
  const examples = [];
  if (normalized.indexOf('gojag') !== -1) {
    examples.push('link gojags yang memuat dokumen terkait');
  }
  if (normalized.indexOf('kalender') !== -1) {
    examples.push('link kalender pelatihan atau tangkapan informasi jadwal');
  }
  if (normalized.indexOf('evaluasi') !== -1) {
    examples.push('link hasil evaluasi atau rekap evaluasi');
  }
  if (normalized.indexOf('dokumen') !== -1 || normalized.indexOf('lms') !== -1) {
    examples.push('link dokumen, LMS, soft copy, atau folder Drive');
  }
  if (normalized.indexOf('pengamatan') !== -1) {
    examples.push('ringkasan hasil pengamatan kondisi lapangan pada Keterangan');
  }
  return examples.length ? examples.join('; ') : 'link dokumen, soft copy, atau keterangan pemeriksaan yang relevan';
}

function noteHelpText_(role) {
  if (role === 'MANDIRI') {
    return [
      '𝗞𝗘𝗧𝗘𝗥𝗔𝗡𝗚𝗔𝗡 𝗣𝗘𝗡𝗬𝗘𝗟𝗘𝗡𝗚𝗚𝗔𝗥𝗔',
      'Isi keterangan bila bukti dukung kosong, indikator tidak berlaku, atau nilai belum penuh.',
      '',
      '𝗖𝗼𝗻𝘁𝗼𝗵 𝗸𝗲𝘁𝗲𝗿𝗮𝗻𝗴𝗮𝗻',
      '- KAK tersedia, tetapi belum memuat kualifikasi pengajar.',
      '- Ruang komputer tidak digunakan pada pelatihan ini.',
      '- Bukti dukung masih dalam proses finalisasi.'
    ].join('\n');
  }

  return [
    '𝗖𝗔𝗧𝗔𝗧𝗔𝗡 𝗩𝗘𝗥𝗜𝗙𝗜𝗞𝗔𝗧𝗢𝗥',
    'Isi catatan bila ada koreksi, klarifikasi, atau bukti belum memadai.',
    '',
    '𝗖𝗼𝗻𝘁𝗼𝗵 𝗰𝗮𝘁𝗮𝘁𝗮𝗻',
    '- Nilai disesuaikan karena bukti hanya menunjukkan sebagian dokumen.',
    '- Perlu melengkapi daftar hadir rapat.',
    '- Indikator tidak berlaku untuk metode pelatihan ini.'
  ].join('\n');
}

function writeMaster_(sheet, indicators) {
  const values = [[
    'Kode Indikator',
    'No',
    'Dimensi',
    'Indikator',
    'Skala'
  ]];

  indicators.forEach(function(item) {
    values.push([item.code, item.no, item.dimension, item.indicator, item.scale]);
  });

  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  formatHeader_(sheet, 1, values[0].length);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, values[0].length);
}

function writeRekap_(sheet, metadata, indicators) {
  const header = rekapSummaryHeader_();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  formatHeader_(sheet, 1, header.length);
  sheet.setFrozenRows(1);
  sheet.getRange(2, 19, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.autoResizeColumns(1, header.length);
}

function writeDetailJawaban_(sheet) {
  const header = detailJawabanHeader_();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  formatHeader_(sheet, 1, header.length);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);
}

function writeDashboard_(sheet) {
  sheet.clear();
  const rows = [
    ['Dashboard Penjaminan Mutu'],
    [''],
    ['Total Pelatihan', '=COUNTA(Rekap!A2:A)'],
    ['Sudah Dinilai', '=COUNTIF(Rekap!K2:K,"Sudah Dinilai")'],
    ['Belum Dinilai', '=COUNTIF(Rekap!K2:K,"Belum Dinilai")'],
    ['Ditolak/Duplikat', '=COUNTIF(Rekap!J2:J,"DITOLAK*")'],
    ['Total Indikator', '=SUM(Rekap!L2:L)'],
    ['Total Terisi', '=SUM(Rekap!M2:M)'],
    ['Total Tidak Terisi', '=SUM(Rekap!N2:N)'],
    ['Total Tidak Berlaku', '=SUM(Rekap!O2:O)'],
    ['Total Bukti Link', '=SUM(Rekap!P2:P)'],
    ['Total Bukti File', '=SUM(Rekap!Q2:Q)'],
    ['Total Catatan', '=SUM(Rekap!R2:R)']
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows.map(function(row) {
    return row.length === 1 ? [row[0], ''] : row;
  }));
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.autoResizeColumns(1, 2);
}

function rekapSummaryHeader_() {
  return [
    'ID Pelatihan',
    'Nama Pelatihan',
    'Nama Pelatihan Normalisasi',
    'Metode',
    'Tanggal Mulai',
    'Tanggal Selesai',
    'Tempat',
    'Unit Penyelenggara',
    'Tanggal Penjaminan Mutu',
    'Status Validasi',
    'Status Penilaian',
    'Jumlah Indikator',
    'Jumlah Terisi',
    'Jumlah Tidak Terisi',
    'Jumlah Tidak Berlaku',
    'Jumlah Bukti Link',
    'Jumlah Bukti File',
    'Jumlah Catatan',
    'Terakhir Diubah',
    'Nama Tab Detail'
  ];
}

function detailJawabanHeader_() {
  return [
    'ID Pelatihan',
    'Nama Pelatihan',
    'Kode Indikator',
    'No',
    'Dimensi',
    'Indikator',
    'Skala',
    'Nilai Mandiri',
    'Link Dokumen',
    'Upload File',
    'Keterangan Penyelenggara',
    'Nilai Penilai',
    'Catatan Penilai',
    'Nilai Akhir',
    'Status Indikator'
  ];
}

function writePanduan_(sheet, source, mandiriForm, spreadsheet) {
  const rows = [
    ['Instrumen Penjaminan Mutu Pelatihan'],
    [''],
    ['Google Sheet Rekap', spreadsheet.getUrl()],
    ['Form Penilaian Mandiri Penyelenggara', mandiriForm.getPublishedUrl()],
    ['Edit Form Penilaian Mandiri Penyelenggara', mandiriForm.getEditUrl()],
    [''],
    ['Alur'],
    ['1. Penyelenggara mengisi Form Penilaian Mandiri Penyelenggara.'],
    ['2. Penyelenggara mengisi nilai, bukti dukung, dan keterangan jika diperlukan.'],
    ['3. Tab Rekap hanya berisi ringkasan satu baris per pelatihan.'],
    ['4. Tab Detail Jawaban berisi jawaban indikator dalam format long-table untuk filter, pivot, dan audit.'],
    ['5. Setiap submit Form Mandiri otomatis membuat/memperbarui tab kerja dengan nama pelatihan.'],
    ['6. Tim Penilai mengisi kolom Nilai Tim Penilai dan Catatan Tim Penilai pada tab pelatihan atau Detail Jawaban.'],
    [''],
    ['Aturan Nilai'],
    ['Nilai menggunakan angka 0 sampai 1. Gunakan - untuk indikator tidak berlaku dan jelaskan di keterangan/catatan.'],
    [''],
    ['Bukti Dukung'],
    ['Bukti dukung dapat berupa link gojags, kalender, evaluasi, dokumen, soft copy, atau keterangan hasil pengamatan kondisi lapangan.'],
    [''],
    ['Catatan Teknis'],
    ['Trigger Apps Script akan memperbarui Rekap, Detail Jawaban, Dashboard, dan tab kerja per pelatihan setiap kali Form Mandiri disubmit.'],
    ['Jumlah Indikator', source.indicators.length]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows.map(function(row) {
    return row.length === 1 ? [row[0], ''] : row;
  }));
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.autoResizeColumns(1, 2);
}

function writeSingleFormPanduan_(sheet, spreadsheet) {
  const rows = [
    ['Instrumen Penjaminan Mutu Pelatihan'],
    [''],
    ['Google Sheet Rekap', spreadsheet.getUrl()],
    [''],
    ['Alur'],
    ['1. Penyelenggara mengisi Form Penilaian Mandiri Penyelenggara.'],
    ['2. Hasil penilaian mandiri masuk ke tab Rekap sebagai ringkasan satu baris per pelatihan.'],
    ['3. Jawaban indikator lengkap masuk ke tab Detail Jawaban dalam format long-table.'],
    ['4. Tim Penilai membuka tab dengan nama pelatihan atau Detail Jawaban untuk menilai.'],
    ['5. Nilai akhir menggunakan Nilai Tim Penilai jika tersedia; jika belum, memakai Nilai Mandiri.'],
    [''],
    ['Aturan Nilai'],
    ['Nilai menggunakan angka 0 sampai 1. Gunakan - untuk indikator tidak berlaku dan jelaskan di keterangan/catatan.'],
    [''],
    ['Bukti Dukung'],
    ['Bukti dukung dapat berupa link gojags, kalender, evaluasi, dokumen, soft copy, atau keterangan hasil pengamatan kondisi lapangan.'],
    [''],
    ['Catatan Teknis'],
    ['Model ini hanya memakai satu Google Form. Rekap digunakan untuk agregat, Detail Jawaban untuk audit indikator, dan tab pelatihan untuk lembar kerja penilai.']
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows.map(function(row) {
    return row.length === 1 ? [row[0], ''] : row;
  }));
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
  sheet.autoResizeColumns(1, 2);
}

function resetSheet_(spreadsheet, name) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  sheet.clear();
  return sheet;
}

function formatHeader_(sheet, row, columnCount) {
  sheet.getRange(row, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#1f4e79')
    .setFontColor('#ffffff')
    .setWrap(true);
}

function itemTitle_(item) {
  return '[' + item.code + '] ' + item.no + '. ' + item.indicator;
}

function findIndicatorRow_(sheet, code) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) {
    return null;
  }

  const codes = sheet.getRange(10, 1, lastRow - 9, 1).getValues();
  for (let i = 0; i < codes.length; i++) {
    if (codes[i][0] === code) {
      return i + 10;
    }
  }
  return null;
}

function formatResponseValue_(value, item) {
  if (item && item.getType && item.getType() === FormApp.ItemType.FILE_UPLOAD) {
    return formatFileUploadResponse_(value);
  }

  if (Array.isArray(value)) {
    return value.join(', ');
  }
  return value || '';
}

function formatFileUploadResponse_(value) {
  const values = Array.isArray(value) ? value : (value ? [value] : []);
  return values.map(function(fileValue) {
    const text = String(fileValue || '').trim();
    if (!text) {
      return '';
    }
    if (/^https?:\/\//i.test(text)) {
      return text;
    }
    return 'https://drive.google.com/open?id=' + encodeURIComponent(text);
  }).filter(String).join('\n');
}

function valueAt_(rows, rowIndex, colIndex) {
  return clean_((rows[rowIndex] || [])[colIndex]);
}

function clean_(value) {
  return String(value || '').trim();
}

function replaceBuktiDukungWithFileUpload() {
  // Ambil ID Form dari konstanta yang sudah ada
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Pastikan EXISTING_MANDIRI_FORM_ID sudah terisi.');
  }

  // Buka form dengan aman menggunakan fungsi bawaan Anda
  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  const items = form.getItems();
  let count = 0;

  // Kita lakukan looping dari belakang (reverse) agar saat menghapus item, index item lain tidak bergeser
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i];
    const title = item.getTitle();
    
    // Cari judul pertanyaan yang diawali dengan kode Indikator, misal: [I001]
    const match = title.match(/^\[(I\d{3})\]\s*(.+)$/);
    if (!match) {
      continue;
    }

    const code = match[1];
    // Gunakan fungsi bawaan Anda untuk menormalkan format text (italic/bold)
    const field = normalizeIndicatorField_(match[2]);

    // Jika pertanyaan adalah Bukti Dukung atau Link Dokumen
    if (field === 'Bukti Dukung' || field === 'Lampirkan Link Dokumen') {
      const oldIndex = item.getIndex();

      try {
        // Buat pertanyaan Upload File baru
        const uploadItem = form.addFileUploadItem()
          .setTitle('[' + code + '] 𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦 𝘉𝘶𝘬𝘵𝘪 𝘋𝘶𝘬𝘶𝘯𝘨')
          .setHelpText('Unggah file bukti dukung yang relevan dengan indikator ini (PDF, Word, Excel, Gambar, dll).')
          .setRequired(false);

        // Pindahkan ke posisi pertanyaan lama
        form.moveItem(uploadItem.getIndex(), oldIndex);

        // Hapus pertanyaan teks/link lama
        form.deleteItem(item);
        
        count++;
        Logger.log('Berhasil mengganti input menjadi Upload File untuk indikator: ' + code);
      } catch (error) {
        // Error handling jika akun GWorkspace membatasi pembuatan Upload File via Script
        Logger.log('Gagal membuat form upload untuk ' + code + '. Error: ' + error.message);
      }
    }
  }

  Logger.log('Selesai. Total pertanyaan yang diganti: ' + count);
  Logger.log('Cek Form Anda di sini: ' + form.getEditUrl());
}

function deleteAllFileUploadItems() {
  // Ambil ID Form dari konstanta atau properti script
  const formId = EXISTING_MANDIRI_FORM_ID;
  if (!formId) {
    Logger.log('Form ID tidak ditemukan di properti script.');
    return;
  }
  const form = FormApp.openById(formId);
  const items = form.getItems();
  let count = 0;

  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i];
    const title = item.getTitle() || '';

    // Cek apakah judul mengandung pola [Ixxx] dan teks "𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦"
    const match = title.match(/^\[I\d{3}\]/);
    if (match && title.indexOf('𝘜𝘱𝘭𝘰𝘢𝘥 𝘍𝘪𝘭𝘦') !== -1) {
      form.deleteItem(item);
      count++;
      Logger.log('Berhasil menghapus: ' + title);
    }
  }
  Logger.log('Berhasil menghapus ' + count + ' item upload file.');
}

function renameBuktiDukungSequential() {
  // Ambil ID Form dari konstanta atau properti script
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Pastikan EXISTING_MANDIRI_FORM_ID sudah terisi.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  const items = form.getItems();
  let counter = 1;

  Logger.log('Mulai proses rename...');

  // Looping dari atas ke bawah (berurutan)
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const title = item.getTitle() || '';
    
    // Cari pertanyaan yang mengandung kata "Bukti Dukung"
    // Pastikan itu adalah pertanyaan bertipe Upload File
    if (title.indexOf('Bukti Dukung') !== -1 && item.getType() === FormApp.ItemType.FILE_UPLOAD) {
      
      // Format angka menjadi 3 digit (contoh: 1 menjadi 001)
      const formattedNumber = Utilities.formatString('%03d', counter);
      const newTitle = '[I' + formattedNumber + '] Bukti Dukung';
      
      // Update judul pertanyaan langsung menggunakan objek item dasar (tanpa asFileUploadItem)
      item.setTitle(newTitle);
      
      Logger.log('Berhasil mengubah judul menjadi: ' + newTitle);
      
      counter++;
    }
  }

  Logger.log('Selesai! Total ' + (counter - 1) + ' pertanyaan Bukti Dukung berhasil di-rename.');
}

function repositionBuktiDukungSequential() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Pastikan EXISTING_MANDIRI_FORM_ID sudah terisi.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  Logger.log('Mulai proses memposisikan ulang...');
  const items = form.getItems();
  const map = {};

  // 1. Pindai semua item dan simpan referensi objeknya ke dalam memori
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const title = item.getTitle() || '';
    
    const match = title.match(/^\[(I\d{3})\]/);
    if (match) {
      const code = match[1]; // contoh: "I001"
      if (!map[code]) {
        map[code] = {};
      }
      
      // Kelompokkan item berdasarkan kode indikatornya
      if (title.indexOf('Bukti Dukung') !== -1 && item.getType() === FormApp.ItemType.FILE_UPLOAD) {
        map[code].uploadItem = item;
      } else if (title.indexOf('𝘕𝘪𝘭𝘢𝘪 𝘔𝘢𝘯𝘥𝘪𝘳𝘪') !== -1 || title.indexOf('Nilai Mandiri') !== -1) {
        map[code].nilaiItem = item;
      }
    }
  }

  let count = 0;

  // 2. Loop indikator I001 sampai I054 dan pindahkan item ke posisi yang tepat
  for (let i = 1; i <= 54; i++) {
    const code = 'I' + Utilities.formatString('%03d', i);
    const data = map[code];
    
    if (data && data.uploadItem && data.nilaiItem) {
      // Ambil index posisi live terbaru
      // Harus diambil langsung karena indeks bisa bergeser tiap kali ada item pindah
      const uploadIndex = data.uploadItem.getIndex();
      const nilaiIndex = data.nilaiItem.getIndex();
      
      // Tentukan target posisi: tepat SETELAH pertanyaan "Nilai Mandiri"
      let targetIndex = nilaiIndex + 1;
      
      // Koreksi kalkulasi pergeseran indeks Google Forms jika item ditarik dari bawah ke atas
      if (uploadIndex < nilaiIndex) {
        targetIndex = nilaiIndex;
      }
      
      // Pindahkan item jika posisinya belum sesuai
      if (uploadIndex !== targetIndex) {
        form.moveItem(data.uploadItem, targetIndex);
        Logger.log('Berhasil memposisikan: [' + code + '] Bukti Dukung ke tempat yang sesuai.');
        count++;
      }
    }
  }

  Logger.log('Selesai! Total ' + count + ' pertanyaan telah sukses dipindahkan ke posisinya masing-masing.');
}

function readFullFormDetails() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Pastikan EXISTING_MANDIRI_FORM_ID sudah terisi.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  Logger.log('Membaca data form dan menyusun JSON...');

  // Siapkan objek utama untuk menampung data
  const formData = {
    metadata: {
      judul: form.getTitle(),
      deskripsi: form.getDescription(),
      id: form.getId(),
      urlEdit: form.getEditUrl(),
      urlPublik: form.getPublishedUrl(),
      kumpulkanEmail: form.collectsEmail(),
      totalItem: 0
    },
    pertanyaan: []
  };
  
  const items = form.getItems();
  formData.metadata.totalItem = items.length;
  
  items.forEach(function(item, index) {
    const type = item.getType();
    const title = item.getTitle() || '(Tanpa Judul)';
    const helpText = item.getHelpText() || '';
    let isRequired = false;
    let choices = [];

    // Deteksi apakah pertanyaan bersifat wajib diisi dan ambil data spesifik tipe
    try {
      if (type === FormApp.ItemType.TEXT) isRequired = item.asTextItem().isRequired();
      else if (type === FormApp.ItemType.PARAGRAPH_TEXT) isRequired = item.asParagraphTextItem().isRequired();
      else if (type === FormApp.ItemType.DATE) isRequired = item.asDateItem().isRequired();
      else if (type === FormApp.ItemType.FILE_UPLOAD) isRequired = item.asFileUploadItem().isRequired();
      else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        const mcItem = item.asMultipleChoiceItem();
        isRequired = mcItem.isRequired();
        choices = mcItem.getChoices().map(function(c) { return c.getValue(); });
      }
      else if (type === FormApp.ItemType.CHECKBOX) {
        const cbItem = item.asCheckboxItem();
        isRequired = cbItem.isRequired();
        choices = cbItem.getChoices().map(function(c) { return c.getValue(); });
      }
      else if (type === FormApp.ItemType.LIST) {
        const listItem = item.asListItem();
        isRequired = listItem.isRequired();
        choices = listItem.getChoices().map(function(c) { return c.getValue(); });
      }
    } catch (e) {
      // Abaikan error jika item tidak mendukung method isRequired (contoh: Page Break, Section Header)
    }

    // Susun objek pertanyaan
    const questionData = {
      indeks: index,
      tipeItem: String(type),
      judul: title,
      deskripsi: helpText
    };

    if (type !== FormApp.ItemType.PAGE_BREAK && type !== FormApp.ItemType.SECTION_HEADER && type !== FormApp.ItemType.IMAGE && type !== FormApp.ItemType.VIDEO) {
      questionData.wajibDiisi = isRequired;
      if (choices.length > 0) {
        questionData.pilihan = choices;
      }
    }

    // Masukkan ke array pertanyaan
    formData.pertanyaan.push(questionData);
  });

  // 1. Ubah objek JavaScript menjadi string JSON yang rapi (indentasi 2 spasi)
  const jsonString = JSON.stringify(formData, null, 2);

  // 2. Buat file .json di Google Drive (root folder Anda)
  const fileName = 'Form_Data_' + form.getId() + '.json';
  const file = DriveApp.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);

  // 3. Tampilkan link di log agar mudah diakses/didownload
  Logger.log('=================================');
  Logger.log('File JSON berhasil dibuat dan disimpan ke Google Drive Anda!');
  Logger.log('Nama File : ' + file.getName());
  Logger.log('Link File : ' + file.getUrl());
  Logger.log('=================================');
}

function updateBuktiDukungDescriptions() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Pastikan EXISTING_MANDIRI_FORM_ID sudah terisi.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  Logger.log('Mulai memperbarui deskripsi Bukti Dukung...');
  const items = form.getItems();
  let count = 0;
  let currentIndicatorText = '';

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const title = item.getTitle() || '';
    const type = item.getType();

    // Lacak konteks indikator saat ini dengan membaca Section Header
    if (type === FormApp.ItemType.SECTION_HEADER) {
      const match = title.match(/^\[I\d{3}\]\s+(.+)$/);
      if (match) {
        currentIndicatorText = match[1];
      }
    }

    // Jika ini adalah pertanyaan Bukti Dukung tipe File Upload
    if (type === FormApp.ItemType.FILE_UPLOAD && title.indexOf('Bukti Dukung') !== -1) {
      if (currentIndicatorText !== '') {
        const helpText = generateSmartHelpText_(currentIndicatorText);
        item.setHelpText(helpText);
        count++;
        Logger.log('Updated deskripsi untuk: ' + title);
      }
    }
  }

  Logger.log('Selesai! Total ' + count + ' deskripsi pertanyaan Bukti Dukung telah diperbarui dengan konteks indikator.');
}

// Fungsi pembantu untuk menebak tipe file dan merangkai deskripsi informatif
function generateSmartHelpText_(indicatorText) {
  let docType = "file/dokumen pendukung";
  let lowerText = indicatorText.toLowerCase();

  // Klasifikasi pintar jenis dokumen berdasarkan kata kunci pada teks indikator
  if (lowerText.indexOf("ruang") !== -1 || lowerText.indexOf("aula") !== -1 || lowerText.indexOf("fasilitas") !== -1 || lowerText.indexOf("wisma") !== -1 || lowerText.indexOf("asrama") !== -1 || lowerText.indexOf("kamar mandi") !== -1 || lowerText.indexOf("meja") !== -1 || lowerText.indexOf("kursi") !== -1 || lowerText.indexOf("papan tulis") !== -1 || lowerText.indexOf("lcd") !== -1 || lowerText.indexOf("pc/laptop") !== -1 || lowerText.indexOf("jaringan") !== -1) {
    docType = "foto dokumentasi fisik/tangkapan layar";
  } else if (lowerText.indexOf("sk ") !== -1 || lowerText.indexOf("surat") !== -1 || lowerText.indexOf("sk/") !== -1) {
    docType = "salinan SK/surat resmi";
  } else if (lowerText.indexOf("undangan") !== -1) {
    docType = "salinan undangan";
  } else if (lowerText.indexOf("daftar hadir") !== -1 || lowerText.indexOf("kehadiran") !== -1) {
    docType = "scan/rekap daftar hadir (presensi)";
  } else if (lowerText.indexOf("notulen") !== -1 || lowerText.indexOf("berita acara") !== -1) {
    docType = "salinan notulensi/berita acara";
  } else if (lowerText.indexOf("evaluasi") !== -1 || lowerText.indexOf("rekaptulasi") !== -1 || lowerText.indexOf("laporan") !== -1) {
    docType = "file rekapitulasi/laporan/hasil evaluasi";
  } else if (lowerText.indexOf("sertifikat") !== -1) {
    docType = "salinan sertifikat";
  } else if (lowerText.indexOf("cv") !== -1) {
    docType = "dokumen Curriculum Vitae (CV)";
  } else if (lowerText.indexOf("kak") !== -1 || lowerText.indexOf("kurikulum") !== -1 || lowerText.indexOf("jadwal") !== -1 || lowerText.indexOf("pedoman") !== -1 || lowerText.indexOf("modul") !== -1) {
    docType = "dokumen perencanaan/pembelajaran";
  }

  // Hapus nomor urut di awal teks indikator jika ada (misal: "1. Terdapat di..." jadi "Terdapat di...")
  let cleanIndicator = indicatorText.replace(/^\d+\.\s*/, '');

  return "Unggah " + docType + " untuk kelengkapan: " + cleanIndicator + ".\n\n" +
         "⚠️ PENTING: Cukup pilih SALAH SATU cara pelaporan.\n" +
         "• Jika mengunggah file di sini ➔ KOSONGKAN isian 'Lampirkan Link Dokumen'.\n" +
         "• Jika menggunakan 'Lampirkan Link Dokumen' ➔ KOSONGKAN isian upload file ini.";
}

function updateKeteranganAndNilaiDescriptions() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Pastikan EXISTING_MANDIRI_FORM_ID sudah terisi.');
  }

  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    return;
  }

  Logger.log('Mulai memperbarui deskripsi Keterangan dan Nilai Mandiri...');
  const items = form.getItems();
  let countKeterangan = 0;
  let countNilai = 0;

  // Teks Bantuan Baru untuk Keterangan
  const helpTextKeterangan = "𝗞𝗘𝗧𝗘𝗥𝗔𝗡𝗚𝗔𝗡 𝗣𝗘𝗡𝗬𝗘𝗟𝗘𝗡𝗚𝗚𝗔𝗥𝗔\nHanya diisi sebagai catatan jika dokumen bukti dukung ada (dinilai 1) namun belum sempurna/tidak lengkap, atau jika indikator tidak berlaku (dinilai -).\n\n𝗖𝗼𝗻𝘁𝗼𝗵 𝗰𝗮𝘁𝗮𝘁𝗮𝗻\n- KAK tersedia, tetapi belum memuat kualifikasi pengajar.\n- Daftar hadir ada, tapi belum ditandatangani semua peserta.\n- Ruang komputer tidak digunakan pada pelatihan ini (indikator tidak berlaku).";

  // Teks Bantuan Baru untuk Nilai Mandiri (Skala Biner 1 atau 0)
  const helpTextNilai = "𝗖𝗔𝗥𝗔 𝗠𝗘𝗡𝗚𝗜𝗦𝗜 𝗡𝗜𝗟𝗔𝗜\nIsi nilai mandiri dengan angka bulat 1 atau 0.\n\n𝗣𝗶𝗹𝗶𝗵𝗮𝗻 𝗻𝗶𝗹𝗮𝗶\n- 1 = bukti tersedia (meskipun belum sempurna).\n- 0 = bukti tidak tersedia sama sekali.\n- - = indikator tidak berlaku.\n\n𝘊𝘢𝘵𝘢𝘵𝘢𝘯\n- Jika bukti ada tapi tidak lengkap, tetap beri nilai 1 dan tulis kekurangannya di isian Keterangan.\n- Jika menggunakan tanda -, jelaskan alasannya pada isian Keterangan.";

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const title = item.getTitle() || '';

    // Cari dan update pertanyaan Keterangan
    if (title.indexOf('Keterangan') !== -1 || title.indexOf('𝘒𝘦𝘵𝘦𝘳𝘢𝘯𝘨𝘢𝘯') !== -1) {
      item.setHelpText(helpTextKeterangan);
      countKeterangan++;
      Logger.log('Updated deskripsi Keterangan: ' + title);
    }

    // Cari dan update pertanyaan Nilai Mandiri
    if (title.indexOf('Nilai Mandiri') !== -1 || title.indexOf('𝘕𝘪𝘭𝘢𝘪 𝘔𝘢𝘯𝘥𝘪𝘳𝘪') !== -1) {
      item.setHelpText(helpTextNilai);
      countNilai++;
      Logger.log('Updated deskripsi Nilai: ' + title);
    }
  }

  Logger.log('Selesai! Total ' + countKeterangan + ' deskripsi Keterangan dan ' + countNilai + ' deskripsi Nilai Mandiri telah diperbarui.');
}

function generateLembarKerjaPembagianTugas() {
  const spreadsheetId = '1iEaiReEt6d1290rRWpm6zw2PR13P89yYkf75mdbH3ZY';
  
  let spreadsheet;
  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      spreadsheet = SpreadsheetApp.create(OUTPUT_NAME + ' - Pembagian Tugas Tim');
    }
  } else {
    spreadsheet = SpreadsheetApp.create(OUTPUT_NAME + ' - Pembagian Tugas Tim');
  }

  const sheetName = 'Pembagian Tugas Bukti Dukung';
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear(); // Bersihkan jika sudah ada
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  Logger.log('Menyusun data lembar kerja...');

  // Data 54 Indikator beserta saran referensi buktinya
  const rows = [
    ['Dimensi', 'Kode', 'Indikator', 'Saran Referensi Detail Bukti (Panduan Operasional)', 'PIC / Penanggung Jawab', 'Status Pengerjaan'],
    ['PERENCANAAN PELATIHAN', 'I001', 'Terdapat di Kalender Pelatihan', 'Buktikan bahwa pelatihan ini ada dalam jadwal resmi. Contoh: Tangkapan layar web kalender BPS atau lampiran SK Kalender Pelatihan Tahunan.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I002', 'KAK (Kerangka Acuan Kerja)', 'Mohon lampirkan salinan lengkap dokumen Kerangka Acuan Kerja (KAK) atau TOR pelatihan yang telah disetujui dan disahkan.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I003', 'Undangan rapat persiapan', 'Sertakan file surat undangan resmi untuk rapat persiapan pelatihan yang ditujukan kepada panitia internal maupun pengajar.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I004', 'Daftar hadir rapat persiapan', 'Lampirkan hasil scan atau rekap absensi kehadiran rapat persiapan yang memuat tanda tangan peserta rapat.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I005', 'notulen rapat persiapan', 'Unggah dokumen notulensi atau berita acara yang merangkum poin-poin hasil rapat persiapan pelatihan.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I006', 'Mitigasi Resiko', 'Sertakan dokumen identifikasi atau matriks mitigasi risiko yang telah disusun untuk mengantisipasi kendala selama pelatihan.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I007', 'Surat izin penyelenggaraan', 'Jika instansi pembina berada di luar BPS, pastikan Anda melampirkan salinan SK atau surat izin penyelenggaraan resmi.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I008', 'Undangan rapat penyamaan persepsi pengajar', 'Lampirkan file surat undangan yang dikirimkan kepada para pengajar untuk mengikuti rapat penyamaan persepsi.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I009', 'Daftar hadir rapat penyamaan persepsi pengajar', 'Unggah scan absensi atau daftar hadir yang membuktikan keikutsertaan pengajar dalam rapat penyamaan persepsi.', '', ''],
    ['PERENCANAAN PELATIHAN', 'I010', 'notulen rapat penyamaan persepsi pengajar', 'Sertakan dokumen notulensi yang mencatat hasil diskusi dan kesepakatan dari rapat penyamaan persepsi pengajar.', '', ''],
    ['TENAGA PELATIHAN', 'I011', 'CV Pengajar', 'Mohon kumpulkan dan lampirkan dokumen Curriculum Vitae (CV) atau riwayat hidup dari seluruh pengajar/fasilitator.', '', ''],
    ['TENAGA PELATIHAN', 'I012', 'Sertifikat Workshop/ToT Pengajar', 'Buktikan kompetensi pengajar dengan melampirkan salinan sertifikat ToT, workshop, atau sertifikasi relevan yang mereka miliki.', '', ''],
    ['TENAGA PELATIHAN', 'I013', 'Evaluasi pengajar', 'Sertakan rekap hasil evaluasi kinerja pengajar dari pelatihan sebelumnya, atau form penilaian kompetensi yang digunakan.', '', ''],
    ['TENAGA PELATIHAN', 'I014', 'SK/ST Tim Peyelenggaraan', 'Unggah salinan Surat Keputusan (SK) atau Surat Tugas (ST) resmi yang menetapkan susunan panitia/tim penyelenggara pelatihan.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I015', 'SK Pedoman Penyelenggaraan Pelatihan (dasar hukum)', 'Lampirkan salinan dokumen SK, peraturan, atau pedoman dasar hukum yang menjadi landasan penyelenggaraan pelatihan ini.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I016', 'SK Tim Perancang Kurikulum', 'Sertakan SK penetapan Tim Perancang Kurikulum beserta salinan sertifikat pelatihan kurikulum dari anggota tim tersebut.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I017', 'Kurikulum pelatihan', 'Mohon lampirkan dokumen resmi yang berisi rincian kurikulum, silabus, atau struktur program materi pelatihan.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I018', 'Jadwal Pelatihan', 'Unggah file PDF atau Excel yang berisi rundown acara atau jadwal pelaksanaan pelatihan secara lengkap.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I019', 'RBPMP dan/atau RP', 'Sertakan dokumen Rancang Bangun Pembelajaran Mata Pelatihan (RBPMP) atau Rencana Pembelajaran (RP) yang digunakan.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I020', 'Media pembelajaran', 'Kumpulkan dan lampirkan softcopy modul materi, bahan tayang (presentasi/PPT), atau media ajar pendukung lainnya.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I021', 'Instrumen kehadiran peserta', 'Unggah dokumen kosong (format) atau rekap absensi kehadiran peserta yang digunakan selama pelatihan berlangsung.', '', ''],
    ['PENYELENGGARAAN PELATIHAN', 'I022', 'Kecukupan konsumsi', 'Buktikan kecukupan konsumsi dengan melampirkan foto dokumentasi saat makan/snack atau salinan tanda terima katering.', '', ''],
    ['PEMANTAUAN DAN EVALUASI', 'I023', 'Petugas Piket', 'Sertakan dokumen jadwal piket panitia harian atau Surat Tugas penunjukan petugas piket yang berjaga di kelas.', '', ''],
    ['PEMANTAUAN DAN EVALUASI', 'I024', 'Dokumentasi pembelajaran', 'Kumpulkan dan unggah foto-foto, video, atau tangkapan layar yang menunjukkan aktivitas pembelajaran sedang berlangsung.', '', ''],
    ['PEMANTAUAN DAN EVALUASI', 'I025', 'Evaluasi kepuasan pelatihan oleh peserta', 'Lampirkan dokumen rekapitulasi hasil kuesioner evaluasi peserta terkait sarana prasarana, panitia, dan penyelenggaraan.', '', ''],
    ['PEMANTAUAN DAN EVALUASI', 'I026', 'Evaluasi pengajar oleh peserta', 'Sertakan file rekap hasil kuesioner penilaian atau evaluasi dari peserta terhadap kinerja masing-masing pengajar.', '', ''],
    ['HASIL PENYELENGGARAAN', 'I027', 'Rapat evaluasi kelulusan peserta dihadiri pengajar', 'Buktikan pelaksanaan rapat evaluasi kelulusan dengan melampirkan undangan, daftar hadir, atau notulensinya.', '', ''],
    ['HASIL PENYELENGGARAAN', 'I028', 'Berita Acara Kelulusan', 'Unggah salinan dokumen Berita Acara Kelulusan peserta yang telah disahkan dan ditandatangani.', '', ''],
    ['HASIL PENYELENGGARAAN', 'I029', 'Persyaratan Kualifikasi kelulusan peserta', 'Sertakan dokumen tata tertib atau kriteria persyaratan lulus (bisa berupa tangkapan layar bagian KAK/Pedoman yang memuat hal ini).', '', ''],
    ['HASIL PENYELENGGARAAN', 'I030', 'Rekaptulasi Nilai Akhir Evaluasi Peserta', 'Lampirkan dokumen atau spreadsheet rekapan nilai akhir seluruh peserta (termasuk nilai kehadiran, pre/post-test, dll).', '', ''],
    ['HASIL PENYELENGGARAAN', 'I031', 'Laporan Penyelenggaraan Pelatihan', 'Mohon unggah dokumen final Laporan Penyelenggaraan Pelatihan yang telah selesai disusun.', '', ''],
    ['PRASARANA PELATIHAN', 'I032', 'Aula (100 orang, viewer, sound, AC)', 'Tunjukkan kelayakan fasilitas. Contoh: Foto Aula saat kosong/terisi yang memperlihatkan AC, sound system, dan proyektor.', '', ''],
    ['PRASARANA PELATIHAN', 'I033', 'Ruang Kelas (40 orang, viewer, sound, AC)', 'Lampirkan dokumentasi berupa foto Ruang Kelas yang menampilkan ketersediaan fasilitas pendukungnya secara utuh.', '', ''],
    ['PRASARANA PELATIHAN', 'I034', 'Ruang Diskusi (10 orang, viewer, sound, AC)', 'Sertakan foto fisik Ruang Diskusi atau Breakout Room yang disiapkan untuk kegiatan kerja kelompok.', '', ''],
    ['PRASARANA PELATIHAN', 'I035', 'Ruang kebugaran', 'Buktikan ketersediaan fasilitas kebugaran dengan mengunggah foto ruangan atau alat olahraga yang dapat digunakan peserta.', '', ''],
    ['PRASARANA PELATIHAN', 'I036', 'Ruang Komputer (40 orang, viewer, sound, AC)', 'Lampirkan foto ruangan komputer/laboratorium IT yang menampilkan jajaran PC/Laptop dan fasilitas ruangannya.', '', ''],
    ['PRASARANA PELATIHAN', 'I037', 'Asrama/Penginapan', 'Sertakan foto yang menunjukkan kondisi kelayakan kamar asrama atau penginapan yang diperuntukkan bagi peserta.', '', ''],
    ['PRASARANA PELATIHAN', 'I038', 'Wisma Tenaga Kepelatihanan', 'Unggah foto dokumentasi kondisi kamar/penginapan khusus yang disediakan untuk tenaga pengajar dan panitia.', '', ''],
    ['PRASARANA PELATIHAN', 'I039', 'Perpustakaan/ Digital Library', 'Buktikan ketersediaan literatur dengan melampirkan tangkapan layar portal Digital Library atau foto fisik perpustakaan.', '', ''],
    ['PRASARANA PELATIHAN', 'I040', 'Ruang makan', 'Lampirkan foto fisik kondisi ruang makan, kafetaria, atau area pantry yang disediakan selama pelatihan.', '', ''],
    ['PRASARANA PELATIHAN', 'I041', 'Fasilitas Olah Raga (indoor dan outdoor)', 'Sertakan dokumentasi foto lapangan olahraga atau fasilitas pendukung aktivitas fisik lainnya (indoor/outdoor).', '', ''],
    ['PRASARANA PELATIHAN', 'I042', 'Fasilitas Hiburan', 'Unggah foto area rekreasi atau fasilitas hiburan ringan yang bisa diakses peserta untuk melepas penat.', '', ''],
    ['PRASARANA PELATIHAN', 'I043', 'Unit Kesehatan', 'Buktikan kesiapan P3K dengan melampirkan foto klinik/ruang kesehatan atau kotak P3K beserta kelengkapannya.', '', ''],
    ['PRASARANA PELATIHAN', 'I044', 'Tempat Ibadah', 'Sertakan foto fasilitas tempat ibadah (masjid/mushola/kapel) yang tersedia di area lokasi pelatihan.', '', ''],
    ['PRASARANA PELATIHAN', 'I045', 'Fasilitas Keamanan', 'Lampirkan foto pendukung keamanan, seperti pos penjagaan, CCTV, jalur evakuasi yang jelas, atau ketersediaan APAR.', '', ''],
    ['PRASARANA PELATIHAN', 'I046', 'Kamar Mandi dan WC', 'Unggah foto yang menunjukkan ketersediaan dan kebersihan fasilitas toilet atau kamar mandi bagi peserta.', '', ''],
    ['SARANA PELATIHAN', 'I047', 'Meja', 'Sertakan foto dokumentasi yang menunjukkan kondisi dan kelayakan meja bagi peserta maupun pengajar di ruang kelas.', '', ''],
    ['SARANA PELATIHAN', 'I048', 'Kursi', 'Lampirkan foto untuk membuktikan ketersediaan dan kelayakan kursi belajar untuk peserta dan pengajar.', '', ''],
    ['SARANA PELATIHAN', 'I049', 'Papan Tulis', 'Unggah foto yang memperlihatkan keberadaan papan tulis (whiteboard/glassboard) di dalam area ruang kelas.', '', ''],
    ['SARANA PELATIHAN', 'I050', 'Flip chart', 'Buktikan ketersediaan sarana pendukung diskusi dengan foto standing flip chart beserta persediaan kertasnya.', '', ''],
    ['SARANA PELATIHAN', 'I051', 'LCD Proyektor/Televisi', 'Sertakan foto proyektor atau Smart TV di ruang kelas, idealnya saat sedang menyala/berfungsi dengan baik.', '', ''],
    ['SARANA PELATIHAN', 'I052', 'PC/Laptop', 'Lampirkan foto dokumentasi perangkat PC/Laptop di meja peserta, terutama untuk pelatihan yang membutuhkan praktik komputer.', '', ''],
    ['SARANA PELATIHAN', 'I053', 'Jaringan Wi-fi', 'Buktikan kualitas koneksi dengan melampirkan tangkapan layar ikon sinyal Wi-Fi atau hasil speedtest internet dari lokasi.', '', ''],
    ['SARANA PELATIHAN', 'I054', 'Fasilitas Pembelajaran dengan Pemanfaatan Teknologi', 'Unggah tangkapan layar yang menunjukkan penggunaan teknologi pendukung, seperti portal LMS, absensi online, atau aplikasi interaktif.', '', '']
  ];

  // Masukkan data ke Sheet
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // Styling Header
  const headerRange = sheet.getRange(1, 1, 1, rows[0].length);
  headerRange.setFontWeight('bold')
             .setBackground('#1f4e79')
             .setFontColor('#ffffff')
             .setWrap(true);

  // Styling Kolom
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(4, 400); // Lebarkan kolom Saran Referensi
  sheet.setColumnWidth(5, 200); // Kolom PIC
  sheet.setColumnWidth(6, 150); // Kolom Status

  // Aktifkan Wrap Text untuk Saran Referensi
  sheet.getRange(2, 4, rows.length - 1, 1).setWrap(true);

  // Buat Dropdown (Data Validation) untuk Status Pengerjaan
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Belum Mulai', 'Sedang Dikerjakan', 'Selesai', 'Tidak Berlaku'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, rows.length - 1, 1).setDataValidation(statusRule);

  Logger.log('=============================================');
  Logger.log('Lembar Kerja Pembagian Tugas Berhasil Dibuat!');
  Logger.log('Klik link berikut untuk membuka Sheet-nya:');
  Logger.log(spreadsheet.getUrl() + '#gid=' + sheet.getSheetId());
  Logger.log('=============================================');
}

// Pola judul untuk identifikasi (Regex). Mencari [Ixxx] diikuti teks khusus italic
const JUDUL_PATTERN = /^\[I\d{3}\].*𝘕𝘪𝘭𝘢𝘪\s*𝘔𝘢𝘯𝘥𝘪𝘳𝘪.*$/;

// Pilihan skor yang diperbolehkan di dalam dropdown
const PILIHAN_SKOR = ['1', '0', '-'];

function ubahInputSkorTeksMenjadiDropdown() {
  let form;
  try {
    form = FormApp.openById(EXISTING_MANDIRI_FORM_ID);
  } catch (e) {
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('GAGAL: ID Form tidak valid atau Anda tidak memiliki akses edit. Mohon cek konfigurasi script.');
    } else {
      Logger.log('GAGAL: ID Form tidak valid atau Anda tidak memiliki akses edit.');
    }
    return;
  }
  
  // Ambil semua item generik dalam form saat ini
  const initialItems = form.getItems();
  let counter = 0;
  
  Logger.log('--- Mulai Proses Update Form: ' + form.getTitle() + ' ---');

  /**
   * Proses item secara MUNDUR (backward loop).
   * Penting agar object referensi item lama tetap valid meskipun indeksnya bergeser.
   */
  for (let i = initialItems.length - 1; i >= 0; i--) {
    let oldGenericItem = initialItems[i]; // Object referensi item lama (generik 'Item')
    let title = oldGenericItem.getTitle();
    let itemType = oldGenericItem.getType();

    // Cek apakah item bertipe TEXT dan judulnya cocok dengan pola skoring
    if (itemType === FormApp.ItemType.TEXT && JUDUL_PATTERN.test(title)) {
      
      Logger.log('Mengonversi target: ' + title);
      
      // 1. Backup data penting dari item teks lama
      let textItem = oldGenericItem.asTextItem();
      let oldHelpText = textItem.getHelpText();
      let oldIsRequired = textItem.isRequired();
      let oldIndex = oldGenericItem.getIndex(); // Posisi urutan saat ini dalam form

      /**
       * 2. Buat item baru bertipe LIST (Dropdown).
       * Item otomatis muncul di paling bawah formulir.
       */
      let newDropdownItem = form.addListItem();

      // 3. Setel properti item baru berdasarkan backup data item lama
      newDropdownItem.setTitle(title)
                     .setHelpText(oldHelpText)
                     .setRequired(oldIsRequired)
                     .setChoiceValues(PILIHAN_SKOR);

      /**
       * 4. FIX PARAH: Ambil referensi 'Item' generik untuk moveItem.
       * Kita tidak bisa menggunakan moveItem(newDropdownItem, ...) karena newDropdownItem 
       * bertipe ListItem. moveItem meminta tipe generik 'Item'.
       * Kita harus ambil ulang semua item, dan ambil yang terakhir (yang baru kita buat).
       */
      let allItemsNow = form.getItems(); // Ambil ulang semua item terbaru
      let newGenericDropdownItem = allItemsNow[allItemsNow.length - 1]; // Referensi generic Item

      // 5. Pindahkan referensi dropdown generik ke posisi item lama.
      form.moveItem(newGenericDropdownItem, oldIndex);
      
      /**
       * 6. Hapus item teks lama.
       * Referensi object 'oldGenericItem' masih valid meskipun indeksnya bergeser
       * setelah pemindahan di langkah 5.
       */
      form.deleteItem(oldGenericItem);

      counter++;
    }
  }
  
  Logger.log('--- Selesai. Total ' + counter + ' item berhasil dikonversi. ---');
  
  // Beri notifikasi ke user via Spreadsheet UI (jika dijalankan dari Spreadsheet)
  if (typeof SpreadsheetApp !== 'undefined') {
    const ui = SpreadsheetApp.getUi();
    if (counter > 0) {
      ui.alert('BERHASIL: Total ' + counter + ' isian skoring telah diubah menjadi dropdown (1, 0, -).\nKeterangan nilai dipertahankan pada deskripsi pertanyaan.');
    } else {
      ui.alert('SELESAI: Tidak ditemukan pertanyaan teks skoring baru yang perlu dikonversi.');
    }
  }
}

/**
 * Fungsi Spesial Apps Script untuk membuat menu kustom di Spreadsheet.
 * Menu ini muncul otomatis saat Spreadsheet dibuka.
 */
function onOpen() {
  if (typeof SpreadsheetApp !== 'undefined') {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚙️ Admin Form')
      .addItem('🚀 Jalankan: Ubah Input Skor ke Dropdown', 'ubahInputSkorTeksMenjadiDropdown')
      .addItem('🏥 Jalankan: Perbaiki & Sembuhkan Form Mandiri', 'perbaikiStrukturFormMandiriOtomatis')
      .addItem('🔍 Jalankan: Audit Kelengkapan Indikator Form', 'auditKelengkapanFormMandiri')
      .addItem('📋 Jalankan: Dump Seluruh Elemen Form (Debug)', 'debugDumpAllFormItems')
      .addToUi();
  }
}

/**
 * FUNGSI DIAGNOSTIK MENDALAM
 * Mengaudit struktur pertanyaan Form Mandiri Live satu per satu.
 * Mengecek apakah ke-54 indikator (dari sumber) masing-masing memiliki 4 komponen wajib.
 */
function auditKelengkapanFormMandiri() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Setel EXISTING_MANDIRI_FORM_ID di bagian atas Kode.js.');
  }
  
  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    Logger.log('❌ GAGAL: Tidak bisa membuka Google Form Mandiri. Periksa hak akses Anda.');
    return;
  }
  
  // 1. Ambil data sumber untuk mendapatkan target indikator
  const source = loadSourceData_();
  const targetIndicators = source.indicators;
  
  // 2. Baca semua item pada Form Live
  const items = form.getItems();
  const liveMap = {}; // Map dari code -> list item info
  
  items.forEach(function(item) {
    const title = item.getTitle();
    // Match kode [I001] sampai [I999]
    const match = title.match(/^\[(I\d{3})\]\s*(.+)$/);
    if (!match) return;
    
    const code = match[1];
    const rawField = match[2];
    const normalizedField = normalizeIndicatorField_(rawField);
    
    if (!liveMap[code]) {
      liveMap[code] = [];
    }
    liveMap[code].push({
      title: title,
      type: item.getType().toString(),
      normalizedField: normalizedField,
      index: item.getIndex()
    });
  });
  
  Logger.log('=== MULAI AUDIT MENDALAM GOOGLE FORM LIVE ===');
  Logger.log('Form Target : ' + form.getTitle());
  Logger.log('ID Form     : ' + form.getId());
  Logger.log('Indikator Target di Sumber: ' + targetIndicators.length);
  
  // Komponen wajib yang harus ada di Form Mandiri
  const KOMPONEN_WAJIB = [
    { key: 'Nilai Mandiri', alias: ['Nilai Mandiri'] },
    { key: 'Lampirkan Link Dokumen', alias: ['Lampirkan Link Dokumen', 'Bukti Dukung'] },
    { key: 'Upload File', alias: ['Upload File', 'Upload Soft Copy'] },
    { key: 'Keterangan', alias: ['Keterangan'] }
  ];
  
  let indicatorOkCount = 0;
  let indicatorBrokenCount = 0;
  let detailedLaporan = [];
  
  targetIndicators.forEach(function(indicator) {
    const code = indicator.code;
    const liveItems = liveMap[code] || [];
    let temuanKomponen = {};
    
    // Map kan item live ke key komponen wajib
    liveItems.forEach(function(item) {
      KOMPONEN_WAJIB.forEach(function(komp) {
        const matchFound = komp.alias.some(function(al) {
          return item.normalizedField.toLowerCase().indexOf(al.toLowerCase()) !== -1;
        });
        if (matchFound) {
          temuanKomponen[komp.key] = item;
        }
      });
    });
    
    let issues = [];
    
    KOMPONEN_WAJIB.forEach(function(komp) {
      if (!temuanKomponen[komp.key]) {
        issues.push('   ❌ HILANG: [' + code + '] ' + komp.key);
      } else {
        const liveItem = temuanKomponen[komp.key];
        // Pengecekan khusus tipe Nilai Mandiri
        if (komp.key === 'Nilai Mandiri' && liveItem.type === 'TEXT') {
          issues.push('   ⚠️ PERINGATAN: ' + komp.key + ' masih tipe ISIAN SINGKAT (Seharusnya Dropdown)');
        }
      }
    });
    
    // Jika tidak ada isu sama sekali
    if (issues.length === 0) {
      indicatorOkCount++;
    } else {
      indicatorBrokenCount++;
      detailedLaporan.push('📌 INDIKATOR [' + code + ']: ' + indicator.indicator);
      if (liveItems.length === 0) {
        detailedLaporan.push('   🚨 FATAL ERROR: Indikator ini tidak ada sama sekali di Google Form Anda!');
      } else {
        issues.forEach(function(issue) {
          detailedLaporan.push(issue);
        });
        detailedLaporan.push('   💡 Info: Ditemukan hanya ' + liveItems.length + ' isian terkait dari seharusnya 4 wajib.');
      }
      detailedLaporan.push('-------------------------------------------');
    }
  });
  
  // Cetak Kesimpulan Audit ke Log
  Logger.log('\n=== HASIL KESIMPULAN AUDIT ===');
  Logger.log('✅ INDIKATOR LENGKAP & SESUAI : ' + indicatorOkCount + ' / ' + targetIndicators.length);
  Logger.log('🚨 INDIKATOR BERMASALAH/KURANG : ' + indicatorBrokenCount + ' / ' + targetIndicators.length);
  Logger.log('==============================\n');
  
  if (detailedLaporan.length > 0) {
    Logger.log('RINCIAN INDIKATOR YANG BERMASALAH:');
    detailedLaporan.forEach(function(line) {
      Logger.log(line);
    });
  } else {
    Logger.log('🎉 SEMPURNA: Semua ' + targetIndicators.length + ' indikator memiliki 4 komponen lengkap dan sesuai standar!');
  }
  
  // Berikan dialog box UI jika berjalan dari spreadsheet (Gunakan Try Catch untuk mencegah error Cloud Context)
  try {
    if (typeof SpreadsheetApp !== 'undefined') {
      const ui = SpreadsheetApp.getUi();
      const statusTxt = indicatorBrokenCount === 0 
        ? '🎉 SEMUA LENGKAP! Seluruh ' + targetIndicators.length + ' indikator memiliki 4 komponen struktur yang sempurna.'
        : '🚨 DITEMUKAN MASALAH: Terdapat ' + indicatorBrokenCount + ' indikator yang strukturnya belum lengkap.\n\nBuka "Extensions > Apps Script > Executions" untuk membaca detail rincian indikator yang hilang.';
      ui.alert('HASIL AUDIT FORM', statusTxt, ui.ButtonSet.OK);
    }
  } catch(e) {
    Logger.log('Info: Pop-up Alert UI dilewati karena script dijalankan di luar antarmuka Sheet.');
  }
}

/**
 * FUNGSI DEBUG UTAMA: Mendump seluruh item yang ada di Google Form live secara presisi.
 * Mencetak urutan Indeks, Judul Asli, dan Tipe Item ke Execution Log.
 */
function debugDumpAllFormItems() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Setel EXISTING_MANDIRI_FORM_ID di Kode.js.');
  }
  
  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    Logger.log('Gagal membuka Google Form untuk melakukan Dump.');
    return;
  }
  
  const items = form.getItems();
  Logger.log('=== DUMP SEMUA ELEMEN GOOGLE FORM SECARA URUTAN ===');
  Logger.log('Nama Form            : ' + form.getTitle());
  Logger.log('Total Elemen Terdeteksi: ' + items.length);
  Logger.log('===================================================');
  
  items.forEach(function(item, index) {
    Logger.log('[' + String(index).padStart(3, '0') + '] Tipe: ' + String(item.getType()).padEnd(18, ' ') + ' | Judul: "' + item.getTitle() + '"');
  });
  
  Logger.log('=================== AKHIR DUMP ====================');
  
  try {
    if (typeof SpreadsheetApp !== 'undefined') {
      SpreadsheetApp.getUi().alert('DUMP SELESAI', 'Total ' + items.length + ' item berhasil dicetak ke Apps Script Execution Log.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch(e) {
    Logger.log('Info: Pop-up alert dilewati.');
  }
}

/**
 * FUNGSI PENYEMBUH OTOMATIS (AUTO-REPAIR & SELF-HEALING)
 * Secara ajaib memulihkan seluruh struktur Google Form Mandiri:
 * 1. Menemukan 'SectionHeader' sebagai jangkar posisi untuk ke-54 Indikator.
 * 2. Mendeteksi apakah dropdown 'Nilai Mandiri' hilang sama sekali. Jika iya, otomatis DIBUATKAN BARU.
 * 3. Memastikan posisi dropdown 'Nilai Mandiri' tepat berada di posisi awal di bawah Header (sebelum Bukti Dukung).
 * 4. Mendeteksi dan MENGHAPUS duplikat isian berlebih/sampah secara otomatis!
 */
function perbaikiStrukturFormMandiriOtomatis() {
  const mandiriFormId = EXISTING_MANDIRI_FORM_ID || PropertiesService.getScriptProperties().getProperty('MANDIRI_FORM_ID');
  if (!mandiriFormId) {
    throw new Error('ID Form Mandiri tidak ditemukan. Setel EXISTING_MANDIRI_FORM_ID di Kode.js.');
  }
  
  const form = openFormSafely_(mandiriFormId, 'Form Mandiri');
  if (!form) {
    Logger.log('❌ GAGAL: Tidak bisa membuka Google Form Mandiri.');
    return;
  }
  
  // Ambil konfigurasi indikator resmi dari sumber
  const source = loadSourceData_();
  const targetCodes = source.indicators.map(function(ind) { return ind.code; });
  
  Logger.log('=======================================================');
  Logger.log('⚡ MEMULAI OPERASI PENYEMBUHAN KILAT (SUPER OPTIMIZED) ⚡');
  Logger.log('=======================================================');
  
  // PENGOPTIMALAN KRUSIAL: HANYA LOAD SEKALI SEUMUR HIDUP UNTUK MENCEGAH TIMEOUT 6 MENIT!
  const allItemsSnapshot = form.getItems();
  Logger.log('Berhasil memuat ' + allItemsSnapshot.length + ' item awal dalam memori.');
  
  // Petakan item-item ke dalam Map Cache untuk akses instan O(1)
  const headerMap = {}; // code -> Item Object
  const nilaiMandiriMap = {}; // code -> Array of Item Objects
  
  allItemsSnapshot.forEach(function(item) {
    const title = item.getTitle();
    const type = item.getType();
    
    // Cek format kode [IXXX]
    const match = title.match(/^\[(I\d{3})\]/);
    if (!match) return;
    const code = match[1];
    
    // 1. Petakan SectionHeader
    if (type === FormApp.ItemType.SECTION_HEADER) {
      headerMap[code] = item;
    }
    
    // 2. Petakan Nilai Mandiri
    const isNilaiMandiri = title.indexOf('Nilai Mandiri') !== -1 || title.indexOf('𝘕𝘪𝘭𝘢𝘪 𝘔𝘢𝘯𝘥𝘪𝘳𝘪') !== -1;
    if (isNilaiMandiri && (type === FormApp.ItemType.LIST || type === FormApp.ItemType.TEXT)) {
      if (!nilaiMandiriMap[code]) nilaiMandiriMap[code] = [];
      nilaiMandiriMap[code].push(item);
    }
  });
  
  let insertedCount = 0;
  let deletedCount = 0;
  let positionedCount = 0;
  
  /**
   * TRIK MATEMATIS BACKWARD LOOP (PROSES MUNDUR):
   * Memproses dari I054 ke I001. 
   * Dengan memproses mundur, penyisipan item di hilir TIDAK AKAN pernah menggeser 
   * indeks item-item yang belum kita proses di hulu! Hasilnya 100% presisi & instan.
   */
  for (let i = targetCodes.length - 1; i >= 0; i--) {
    const code = targetCodes[i];
    const headerItem = headerMap[code];
    
    if (!headerItem) {
      Logger.log('⚠️ PERINGATAN [' + code + ']: Section Header tidak ditemukan di Form. Dilewati.');
      continue;
    }
    
    // Dapatkan indeks live terkini (akan otomatis terupdate jika terjadi pergeseran)
    const currentHeaderIndex = headerItem.getIndex();
    const idealIndex = currentHeaderIndex + 1;
    const foundItems = nilaiMandiriMap[code] || [];
    
    // KONDISI A: NILAI MANDIRI HILANG SAMA SEKALI
    if (foundItems.length === 0) {
      Logger.log('➕ [' + code + ']: Menyisipkan Dropdown Nilai Mandiri baru...');
      
      // Buat elemen dropdown
      const newDropdown = form.addListItem();
      newDropdown.setTitle('[' + code + '] 𝘕𝘪𝘭𝘢𝘪 𝘔𝘢𝘯𝘥𝘪𝘳𝘪')
                 .setHelpText('Pilih skor penilaian mandiri untuk indikator ini (1. ada / 0. tidak).')
                 .setRequired(true)
                 .setChoiceValues(['1', '0', '-']);
      
      // PENGOPTIMALAN KRUSIAL: Ambil item generik via ID (jauh lebih cepat daripada getItems)
      const newGenericItem = form.getItemById(newDropdown.getId());
      form.moveItem(newGenericItem, idealIndex);
      
      insertedCount++;
      Logger.log('   ✅ SUKSES: Ditambahkan tepat di bawah Header.');
    }
    
    // KONDISI B: NILAI MANDIRI HANYA ADA SATU (NORMAL / DIRAPIKAN)
    else if (foundItems.length === 1) {
      const singleItem = foundItems[0];
      const currentItemIndex = singleItem.getIndex(); // Cek indeks live-nya
      
      if (currentItemIndex !== idealIndex) {
        Logger.log('🚚 [' + code + ']: Merapikan posisi isian yang melenceng...');
        form.moveItem(singleItem, idealIndex);
        positionedCount++;
      }
    }
    
    // KONDISI C: TERDAPAT DUPLIKAT SAMPAH GANDA
    else {
      Logger.log('🚨 [' + code + ']: Ditemukan ' + foundItems.length + ' duplikat! Melakukan pembersihan...');
      
      // Sisakan hanya item pertama, hapus sampah duplikat sisanya
      for (let k = foundItems.length - 1; k > 0; k--) {
        form.deleteItem(foundItems[k]);
        deletedCount++;
      }
      
      // Rapikan posisi satu-satunya item yang kita pertahankan ke posisi ideal
      const remainingItem = foundItems[0];
      if (remainingItem.getIndex() !== idealIndex) {
        form.moveItem(remainingItem, idealIndex);
        positionedCount++;
      }
      Logger.log('   ✅ SUKSES: Duplikat dihapus dan sisanya dirapikan.');
    }
  }
  
  Logger.log('\n=======================================================');
  Logger.log('🎉 OPERASI PENYEMBUHAN KILAT SELESAI! 🎉');
  Logger.log('=======================================================');
  Logger.log('✅ Baru Dibuat      : ' + insertedCount);
  Logger.log('🚚 Dirapikan Posisi : ' + positionedCount);
  Logger.log('🗑️ Sampah Dihapus    : ' + deletedCount);
  Logger.log('=======================================================\n');
  
  try {
    if (typeof SpreadsheetApp !== 'undefined') {
      const ui = SpreadsheetApp.getUi();
      ui.alert('⚡ PENYEMBUHAN OTOMATIS SELESAI ⚡', 
               'Sistem berhasil memperbaiki Form Mandiri Anda hanya dalam hitungan detik!\n\n' +
               '➕ Ditambahkan Baru: ' + insertedCount + ' Dropdown Nilai Mandiri\n' +
               '🚚 Dirapikan Posisi : ' + positionedCount + ' Dropdown\n' +
               '🗑️ Duplikat Dihapus : ' + deletedCount + ' item sampah\n\n' +
               'Formulir Anda sekarang 100% SEMPURNA & LENGKAP.', 
               ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log('Info: Eksekusi luar antarmuka Google Sheets.');
  }
}
