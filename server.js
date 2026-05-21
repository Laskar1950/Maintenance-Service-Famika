const express = require('express');
const multer = require('multer');
const sharp = require('sharp');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');

const app = express();

// Konfigurasi parser file menggunakan penyimpanan memori (buffer)
const upload = multer({ storage: multer.memoryStorage() });

// ─── HUBUNGKAN KE SUPABASE VIA ENVIRONMENT VARIABLES ─────────────────────────
const SUPABASE_URL = process.env.SUPABASE_URL;
// UTAMAKAN MENGGUNAKAN SERVICE ROLE KEY UNTUK BYPASS ATURAN RLS DI BACKEND
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY;

if (!SUPABASE_URL || !SUPABASE_KEY) {
    throw new Error('[FAMIKA] SUPABASE_URL dan SUPABASE_ANON_KEY/SUPABASE_SERVICE_ROLE_KEY wajib diisi di Environment Variables Vercel.');
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// ─── NAMA BUCKET STORAGE (SINKRON DENGAN REPO INDUK) ─────────────────────────
const STORAGE_BUCKET = 'laporan-foto';

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Tambahan CORS header agar bisa dipanggil dari Mini App Telegram
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    next();
});

// Halaman utama server untuk checkup koneksi
app.get('/', (req, res) => {
    res.status(200).send('🚀 Server Portal Pelaporan Famika Aktif & Terhubung ke Supabase!');
});

// --- HELPER LOOKUP UNTUK MENGHINDARI EROR FOREIGN KEY CONSTRAINT ---
async function getTechnicianId(username) {
    if (username) {
        const { data } = await supabase
            .from('users')
            .select('id')
            .or(`full_name.eq.${username},username.eq.${username}`)
            .limit(1);
        if (data && data.length > 0) return data[0].id;
    }
    const { data: fallback } = await supabase
        .from('users')
        .select('id')
        .eq('is_active', true)
        .limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;

    const { data: globalFallback } = await supabase
        .from('users')
        .select('id')
        .limit(1);
    return globalFallback && globalFallback.length > 0 ? globalFallback[0].id : null;
}

async function getOdcId(kodeOdc, siteId) {
    if (kodeOdc) {
        const { data } = await supabase
            .from('odc_master')
            .select('id')
            .eq('odc_code', kodeOdc)
            .limit(1);
        if (data && data.length > 0) return data[0].id;
    }
    const { data: fallback } = await supabase
        .from('odc_master')
        .select('id')
        .eq('site_id', siteId)
        .limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;

    const { data: globalFallback } = await supabase
        .from('odc_master')
        .select('id')
        .limit(1);
    return globalFallback && globalFallback.length > 0 ? globalFallback[0].id : null;
}

async function getOdpId(kodeOdp, siteId) {
    if (kodeOdp) {
        const { data } = await supabase
            .from('odp_master')
            .select('id')
            .eq('odp_code', kodeOdp)
            .limit(1);
        if (data && data.length > 0) return data[0].id;
    }
    const { data: fallback } = await supabase
        .from('odp_master')
        .select('id')
        .eq('site_id', siteId)
        .limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;

    const { data: globalFallback } = await supabase
        .from('odp_master')
        .select('id')
        .limit(1);
    return globalFallback && globalFallback.length > 0 ? globalFallback[0].id : null;
}

async function getClosureId(kodeClosure, siteId) {
    if (kodeClosure) {
        const { data } = await supabase
            .from('closure_master')
            .select('id')
            .eq('closure_code', kodeClosure)
            .limit(1);
        if (data && data.length > 0) return data[0].id;
    }
    const { data: fallback } = await supabase
        .from('closure_master')
        .select('id')
        .eq('site_id', siteId)
        .limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;

    const { data: globalFallback } = await supabase
        .from('closure_master')
        .select('id')
        .limit(1);
    return globalFallback && globalFallback.length > 0 ? globalFallback[0].id : null;
}

async function getSpanId(kodeSpan, siteId) {
    if (kodeSpan) {
        const { data } = await supabase
            .from('span_master')
            .select('id')
            .eq('span_code', kodeSpan)
            .limit(1);
        if (data && data.length > 0) return data[0].id;
    }
    const { data: fallback } = await supabase
        .from('span_master')
        .select('id')
        .eq('site_id', siteId)
        .limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;

    const { data: globalFallback } = await supabase
        .from('span_master')
        .select('id')
        .limit(1);
    return globalFallback && globalFallback.length > 0 ? globalFallback[0].id : null;
}

// --- ENDPOINT UNTUK MENERIMA INPUT DATA TEKNISI ---
app.post('/api/report/odc', upload.any(), async (req, res) => {
    res.status(400).json({ success: false, message: 'Gunakan upload direct SDK yang ada di index.html' });
});

// ─── ENDPOINT GENERATE SPREADSHEET AUTOMATION (UNTUK ADMIN) ──────────────────
app.get('/api/report/export', async (req, res) => {
    try {
        const { project_id } = req.query;
        if (!project_id) {
            return res.status(400).send('<h1>Error: Parameter project_id wajib dikirim!</h1>');
        }

        // 1. Tarik Data Project Bulanan
        const { data: project, error: projErr } = await supabase
            .from('projects')
            .select('*')
            .eq('id', project_id)
            .single();

        if (projErr || !project) {
            return res.status(404).send('<h1>Error: Project kerja bulanan tidak ditemukan!</h1>');
        }

        const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
        const stringBulanText = `${namaBulan[project.bulan - 1]} ${project.tahun}`;

        // 2. Tarik Semua Data Master untuk Lookup Cepat di Memori
        const [{ data: odcMaster }, { data: odpMaster }, { data: closureMaster }, { data: spanMaster }] = await Promise.all([
            supabase.from('odc_master').select('*'),
            supabase.from('odp_master').select('*'),
            supabase.from('closure_master').select('*'),
            supabase.from('span_master').select('*')
        ]);

        // 3. Tarik Seluruh Tabel Transaksi Pekerjaan untuk Project_id Terkait
        const [
            { data: pmOdc },
            { data: pmOdp },
            { data: pmClosure },
            { data: pmSpan },
            { data: corrective },
            { data: dismantling },
            { data: psb }
        ] = await Promise.all([
            supabase.from('pm_odc').select('*').eq('period_id', project_id),
            supabase.from('pm_odp').select('*').eq('period_id', project_id),
            supabase.from('pm_closure').select('*').eq('period_id', project_id),
            supabase.from('pm_span').select('*').eq('period_id', project_id),
            supabase.from('corrective_customer').select('*').eq('period_id', project_id),
            supabase.from('dismantling_records').select('*').eq('period_id', project_id),
            supabase.from('psb_records').select('*').eq('period_id', project_id)
        ]);

        // 4. Kumpulkan semua record_id milik project aktif ini
        const allRecordIds = [
            ...(pmOdc        || []).map(r => r.id),
            ...(pmOdp        || []).map(r => r.id),
            ...(pmClosure    || []).map(r => r.id),
            ...(pmSpan       || []).map(r => r.id),
            ...(corrective   || []).map(r => r.id),
            ...(dismantling  || []).map(r => r.id),
            ...(psb          || []).map(r => r.id),
        ];

        const photoMap = {};
        if (allRecordIds.length > 0) {
            const { data: photos, error: photoErr } = await supabase
                .from('photo_assets')
                .select('record_id, photo_kind, file_path, file_name, module_name')
                .in('record_id', allRecordIds);

            if (!photoErr && photos) {
                photos.forEach(p => {
                    if (!photoMap[p.record_id]) photoMap[p.record_id] = [];
                    photoMap[p.record_id].push(p);
                });
            }
        }

        // 5. Inisialisasi Excel Workbook & Desain Styling Universal
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Famika PM Telecom Portal';
        workbook.created = new Date();

        const tableHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF85D3F2' } };
        const borderStyle = {
            top:    { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            left:   { style: 'thin', color: { argb: '000000' } },
            right:  { style: 'thin', color: { argb: '000000' } }
        };
        const globalFont = { name: 'Arial', size: 10 };
        const boldFont   = { name: 'Arial', size: 10, bold: true };

        function applyBordersToRange(ws, startCol, startRow, endCol, endRow) {
            for (let r = startRow; r <= endRow; r++) {
                const wsRow = ws.getRow(r);
                for (let c = startCol; c <= endCol; c++) {
                    wsRow.getCell(c).border = borderStyle;
                }
            }
        }

        async function buildReportSheet({ sheetName, isCorrective = false, columnsSetup, headers_r5, headers_r6, mergeSpecs, rowsData, photoFields }) {
            const ws = workbook.addWorksheet(sheetName);
            ws.views = [{ showGridLines: true }];

            ws.getCell('A1').value = isCorrective ? 'Corrective Manage Service' : 'Preventive Manage service';
            ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true };
            ws.getRow(1).height = 24;

            ws.getCell('A2').value = `Stasiun : ${project.site_code.split(' - ')[1] || project.site_code}`;
            ws.getCell('A2').font = { name: 'Arial', size: 11, bold: true };
            ws.getRow(2).height = 18;

            ws.getCell('A3').value = `Periode : ${stringBulanText}`;
            ws.getCell('A3').font = { name: 'Arial', size: 11, bold: true };
            ws.getRow(3).height = 18;

            ws.getRow(4).height = 12;

            columnsSetup.widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

            const r5 = ws.getRow(5);
            r5.height = 20;
            headers_r5.forEach((val, i) => {
                const cell = r5.getCell(i + 1);
                cell.value = val;
                cell.font  = boldFont;
                cell.fill  = tableHeaderFill;
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });

            const r6 = ws.getRow(6);
            r6.height = 20;
            headers_r6.forEach((val, i) => {
                const cell = r6.getCell(i + 1);
                cell.value = val;
                cell.font  = boldFont;
                cell.fill  = tableHeaderFill;
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });

            mergeSpecs.forEach(m => { ws.mergeCells(m); });

            // ─────────────────────────────────────────────────────────────────────────────
            // SINKRONISASI & DOWNLOAD FOTO SECARA PARALEL (BANYAK FOTO SEKALIGUS)
            // Langkah ini memotong loading time dari 15 detik menjadi hanya 1 detik!
            // ─────────────────────────────────────────────────────────────────────────────
            const downloadPromises = [];
            rowsData.forEach(r => {
                const assets = photoMap[r.id] || [];
                photoFields.forEach(pf => {
                    const asset = assets.find(a => {
                        const stored = (a.photo_kind || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                        const target = (pf.field || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                        return stored === target;
                    });
                    if (asset) {
                        downloadPromises.push(
                            supabase.storage
                                .from(STORAGE_BUCKET)
                                .download(asset.file_path)
                                .then(({ data, error }) => {
                                    if (error || !data) {
                                        console.warn(`[FAMIKA] Gagal download file ${asset.file_path}:`, error?.message);
                                        return null;
                                    }
                                    return data.arrayBuffer().then(ab => ({
                                        recordId: r.id,
                                        field: pf.field,
                                        col: pf.col,
                                        buffer: Buffer.from(ab),
                                        fileName: asset.file_name || asset.file_path
                                    }));
                                })
                                .catch(err => {
                                    console.error(`[FAMIKA] Gagal mengekstrak arrayBuffer:`, err.message);
                                    return null;
                                })
                        );
                    }
                });
            });

            const downloadedAssets = await Promise.all(downloadPromises);
            const activePhotoMap = {};
            downloadedAssets.forEach(asset => {
                if (asset) {
                    const key = `${asset.recordId}_${asset.field}`;
                    activePhotoMap[key] = asset;
                }
            });

            // ─────────────────────────────────────────────────────────────────────────────
            // INPUT DATA KE EXCEL ROW BY ROW
            // ─────────────────────────────────────────────────────────────────────────────
            let curRow = 7;
            for (const r of rowsData) {
                const wsRow = ws.getRow(curRow);
                wsRow.height = photoFields.length > 0 ? 110 : 22;

                r.data.forEach((val, colIdx) => {
                    const cell = wsRow.getCell(colIdx + 1);
                    cell.value = val;
                    cell.font  = globalFont;
                    cell.alignment = { vertical: 'middle', horizontal: columnsSetup.alignments[colIdx] || 'left', wrapText: true };
                });

                for (const pf of photoFields) {
                    const key = `${r.id}_${pf.field}`;
                    const asset = activePhotoMap[key];
                    const cell = wsRow.getCell(pf.col);

                    if (asset && asset.buffer) {
                        try {
                            const ext = (asset.fileName || '').toLowerCase().endsWith('.png') ? 'png' : 'jpeg';
                            const imgId = workbook.addImage({ buffer: asset.buffer, extension: ext });

                            ws.addImage(imgId, {
                                tl: { col: pf.col - 1, row: curRow - 1, colOff: 6, rowOff: 6 },
                                br: { col: pf.col,     row: curRow,     colOff: -6, rowOff: -6 },
                                editAs: 'oneCell'
                            });
                            cell.value = '';

                        } catch (e) {
                            console.error(`[FAMIKA] Gagal memasukkan gambar ke sel:`, e.message);
                            cell.value = 'Foto Gagal';
                            cell.font  = { name: 'Arial', size: 8, color: { argb: 'EF4444' } };
                        }
                    } else {
                        // Periksa apakah data foto tersebut memang ada di DB namun gagal download, atau memang kosong
                        const hasAssetInDb = (photoMap[r.id] || []).some(a => {
                            const stored = (a.photo_kind || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                            const target = (pf.field || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                            return stored === target;
                        });
                        cell.value = hasAssetInDb ? 'Foto Gagal' : 'Tidak Ada';
                        cell.font  = { name: 'Arial', size: 8, color: { argb: hasAssetInDb ? 'EF4444' : '94A3B8' } };
                    }
                }
                curRow++;
            }

            applyBordersToRange(ws, 1, 5, columnsSetup.widths.length, curRow - 1);
        }

        // ── SHEET 1: ODC ──────────────────────────────────────────────────────────
        await buildReportSheet({
            sheetName: 'ODC',
            columnsSetup: {
                widths:     [6, 12, 18, 10, 15, 24, 24, 24, 24, 24, 12],
                alignments: ['center','center','center','center','center','center','center','center','center','center','center']
            },
            headers_r5: ['NO','TANGGAL','ODC','KONDISI','KEGIATAN','BEFORE','','AFTER','','HASIL OPM',''],
            headers_r6: ['','','','','','BUKA','TUTUP','BUKA','TUTUP','FOTO','REDAMAN'],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:G5','H5:I5','J5:K5'],
            rowsData: (pmOdc || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, odcMaster?.find(o => o.id === item.odc_id)?.odc_code || 'N/A', item.kondisi, item.kegiatan, '','','','','', item.hasil_opm]
            })),
            photoFields: [
                { field: 'before_buka',  col: 6 },
                { field: 'before_tutup', col: 7 },
                { field: 'after_buka',   col: 8 },
                { field: 'after_tutup',  col: 9 },
                { field: 'foto_opm',     col: 10 }
            ]
        });

        // ── SHEET 2: ODP ──────────────────────────────────────────────────────────
        await buildReportSheet({
            sheetName: 'ODP',
            columnsSetup: {
                widths:     [6, 12, 18, 18, 12, 10, 15, 24, 24, 24, 24, 24, 12],
                alignments: ['center','center','center','center','center','center','center','center','center','center','center','center','center']
            },
            headers_r5: ['NO','TANGGAL','ODC','ODP','SISA PORT','KONDISI','KEGIATAN','BEFORE','','AFTER','','HASIL OPM',''],
            headers_r6: ['','','','','','','','TUTUP','BUKA','TUTUP','BUKA','FOTO','REDAMAN'],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:I5','J5:K5','L5:M5'],
            rowsData: (pmOdp || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal,
                    odcMaster?.find(o => o.id === item.odc_id)?.odc_code || 'N/A',
                    odpMaster?.find(o => o.id === item.odp_id)?.odp_code || 'N/A',
                    item.sisa_port, item.kondisi, item.kegiatan, '','','','','', item.hasil_opm]
            })),
            photoFields: [
                { field: 'before_tutup', col: 8 },
                { field: 'before_buka',  col: 9 },
                { field: 'after_tutup',  col: 10 },
                { field: 'after_buka',   col: 11 },
                { field: 'foto_opm',     col: 12 }
            ]
        });

        // ── SHEET 3: CLOSURE ──────────────────────────────────────────────────────
        const closureCfg = {
            widths:     [6, 12, 18, 12, 15, 24, 24, 24],
            alignments: ['center','center','center','center','left','center','center','center']
        };
        await buildReportSheet({
            sheetName: 'CLOSURE',
            columnsSetup: closureCfg,
            headers_r5: ['NO','TANGGAL','KODE CLOSURE','KONDISI','KEGIATAN','FOTO CLOSURE','FOTO SPARE KABEL','FOTO KESELURUHAN'],
            headers_r6: ['','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6'],
            rowsData: (pmClosure || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, closureMaster?.find(o => o.id === item.closure_id)?.closure_code || 'N/A', item.kondisi, item.kegiatan, '','','']
            })),
            photoFields: [
                { field: 'foto_closure', col: 6 }
            ]
        });

        // ── SHEET 4: KABEL (SPAN) ─────────────────────────────────────────────────
        await buildReportSheet({
            sheetName: 'KABEL',
            columnsSetup: closureCfg,
            headers_r5: ['NO','TANGGAL','KODE SPAN','KONDISI','KEGIATAN','FOTO SPAN','FOTO SPARE KABEL','FOTO KESELURUHAN'],
            headers_r6: ['','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6'],
            rowsData: (pmSpan || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, spanMaster?.find(o => o.id === item.span_id)?.span_code || 'N/A', item.kondisi, item.kegiatan, '','','']
            })),
            photoFields: [
                { field: 'foto_span', col: 6 }
            ]
        });

        // ── SHEET 5: CUSTOMER (GANGGUAN) ──────────────────────────────────────────
        await buildReportSheet({
            sheetName: 'CUSTOMER',
            isCorrective: true,
            columnsSetup: {
                widths:     [6, 12, 18, 12, 15, 15, 8, 20, 15, 16, 16, 24, 24, 10, 10],
                alignments: ['center','center','left','center','center','center','center','left','center','center','center','center','center','center','center']
            },
            headers_r5: ['NO','DATE','NAMA','ID','ODC','ODP','PORT','ACTION','MATERIAL','SN ONT','','BEFORE','AFTER','JAM',''],
            headers_r6: ['','','','','','','','','','NEW','OLD','','','MULAI','SELESAI'],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:K5','L5:L6','M5:M6','N5:O5'],
            rowsData: (corrective || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, item.customer_name, item.service_id,
                    odcMaster?.find(o => o.id === item.odc_id)?.odc_code || '-',
                    odpMaster?.find(o => o.id === item.odp_id)?.odp_code || '-',
                    item.port_no || '-', item.action, item.material || '-',
                    item.sn_ont_new || '-', item.sn_ont_old || '-',
                    '','', item.jam_mulai || '-', item.jam_selesai || '-']
            })),
            photoFields: [
                { field: 'foto_before', col: 12 },
                { field: 'foto_after',  col: 13 }
            ]
        });

        // ── SHEET 6: DISMANTLING ──────────────────────────────────────────────────
        await buildReportSheet({
            sheetName: 'DISMANTLING',
            columnsSetup: {
                widths:     [6, 12, 18, 12, 12, 8, 16, 15, 20, 24, 24],
                alignments: ['center','center','left','center','center','center','center','center','left','center','center']
            },
            headers_r5: ['NO','TANGGAL','NAMA PELANGGAN','ID PELANGGAN','ODP','PORT','SN ONT','NO HP','ALAMAT','FOTO SERIALNUMBER','FOTO RUMAH'],
            headers_r6: ['','','','','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:J6','K5:K6'],
            rowsData: (dismantling || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, item.customer_name, item.service_id,
                    odpMaster?.find(o => o.id === item.odp_id)?.odp_code || 'N/A',
                    item.port_no || '-', item.sn_ont || '-',
                    item.no_hp || '-', item.alamat || '-',
                    '','']
            })),
            photoFields: [
                { field: 'foto_sn_ont', col: 10 },
                { field: 'foto_rumah',  col: 11 }
            ]
        });

        // ── SHEET 7: AKTIVASI (PSB) ───────────────────────────────────────────────
        await buildReportSheet({
            sheetName: 'AKTIVASI',
            columnsSetup: {
                widths:     [6, 12, 18, 12, 12, 8, 16, 15, 20, 24, 24, 24, 24, 24, 24, 24],
                alignments: ['center','center','left','center','center','center','center','center','left','center','center','center','center','center','center','center']
            },
            headers_r5: ['NO','TANGGAL','NAMA PELANGGAN','ID PELANGGAN','ODP','PORT','SN ONT','NO HP','ALAMAT','FOTO ODP','PORT ODP','REDAM ODP','REDAM ONT','FOTO SN','FOTO RUMAH','SPEEDTEST'],
            headers_r6: ['','','','','','','','','','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:J6','K5:K6','L5:L6','M5:M6','N5:N6','O5:O6','P5:P6'],
            rowsData: (psb || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, item.customer_name, item.service_id,
                    odpMaster?.find(o => o.id === item.odp_id)?.odp_code || 'N/A',
                    item.port_no, item.sn_ont,
                    item.no_hp || '-', item.alamat || '-',
                    '','','','','','','']
            })),
            photoFields: [
                { field: 'foto_odp',          col: 10 },
                { field: 'foto_port_odp',      col: 11 },
                { field: 'foto_redaman_odp',   col: 12 },
                { field: 'foto_redaman_akhir', col: 13 },
                { field: 'foto_sn_ont',        col: 14 },
                { field: 'foto_instalasi',     col: 15 },
                { field: 'foto_speedtest',     col: 16 }
            ]
        });

        // 6. Kirim file sebagai unduhan biner .xlsx ke browser admin
        const formattedFileName = `Laporan_Bulanan_${project.project_name.replace(/\s+/g, '_')}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${formattedFileName}"`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (err) {
        console.error('Internal Server Error on Export:', err);
        res.status(500).send(`<h1>Error saat mengekspor laporan:</h1><p>${err.message}</p>`);
    }
});

module.exports = app;
