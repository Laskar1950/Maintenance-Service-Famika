const express = require('express');
const multer = require('multer');
const sharp = require('sharp');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');

const app = express();

const upload = multer({ storage: multer.memoryStorage() });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY;

if (!SUPABASE_URL || !SUPABASE_KEY) {
    throw new Error('[FAMIKA] SUPABASE_URL dan SUPABASE_ANON_KEY/SUPABASE_SERVICE_ROLE_KEY wajib diisi di Environment Variables Vercel.');
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
const STORAGE_BUCKET = 'laporan-foto';

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    if (req.method === 'OPTIONS') return res.sendStatus(200);
    next();
});

app.get('/', (req, res) => {
    res.status(200).send('🚀 Server Portal Pelaporan Famika Aktif & Terhubung ke Supabase!');
});

// ─── MODULE TABLE MAP ─────────────────────────────────────────────────────────
const MODULE_TABLE_MAP = {
    form_odc:        'pm_odc',
    form_odp:        'pm_odp',
    form_closure:    'pm_closure',
    form_span:       'pm_span',
    form_gangguan:   'corrective_customer',
    form_dismantling:'dismantling_records',
    form_psb:        'psb_records',
    odc:             'pm_odc',
    odp:             'pm_odp',
    closure:         'pm_closure',
    span:            'pm_span',
    gangguan:        'corrective_customer',
    dismantling:     'dismantling_records',
    psb:             'psb_records',
};

// ─── HELPER: Download dari Storage lalu konversi ke JPEG via sharp ────────────
async function downloadAsJpegBuffer(filePath) {
    const { data, error } = await supabase.storage
        .from(STORAGE_BUCKET)
        .download(filePath);

    if (error || !data) {
        console.warn(`[FAMIKA] Gagal download: ${filePath} —`, error?.message);
        return null;
    }

    try {
        const ab  = await data.arrayBuffer();
        const raw = Buffer.from(ab);
        const jpegBuf = await sharp(raw)
            .rotate()
            .resize({ width: 800, withoutEnlargement: true })
            .jpeg({ quality: 80 })
            .toBuffer();
        return jpegBuf;
    } catch (e) {
        console.error(`[FAMIKA] sharp gagal konversi ${filePath}:`, e.message);
        return null;
    }
}

// ─── HELPER LOOKUP ─────────────────────────────────────────────────────────────
async function getTechnicianId(username) {
    if (username) {
        const { data } = await supabase.from('users').select('id').or(`full_name.eq.${username},username.eq.${username}`).limit(1);
        if (data && data.length > 0) return data[0].id;
    }
    const { data: fallback } = await supabase.from('users').select('id').eq('is_active', true).limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;
    const { data: gf } = await supabase.from('users').select('id').limit(1);
    return gf && gf.length > 0 ? gf[0].id : null;
}
async function getOdcId(kodeOdc, siteId) {
    if (kodeOdc) { const { data } = await supabase.from('odc_master').select('id').eq('odc_code', kodeOdc).limit(1); if (data && data.length > 0) return data[0].id; }
    const { data: f } = await supabase.from('odc_master').select('id').eq('site_id', siteId).limit(1);
    if (f && f.length > 0) return f[0].id;
    const { data: gf } = await supabase.from('odc_master').select('id').limit(1);
    return gf && gf.length > 0 ? gf[0].id : null;
}
async function getOdpId(kodeOdp, siteId) {
    if (kodeOdp) { const { data } = await supabase.from('odp_master').select('id').eq('odp_code', kodeOdp).limit(1); if (data && data.length > 0) return data[0].id; }
    const { data: f } = await supabase.from('odp_master').select('id').eq('site_id', siteId).limit(1);
    if (f && f.length > 0) return f[0].id;
    const { data: gf } = await supabase.from('odp_master').select('id').limit(1);
    return gf && gf.length > 0 ? gf[0].id : null;
}
async function getClosureId(kodeClosure, siteId) {
    if (kodeClosure) { const { data } = await supabase.from('closure_master').select('id').eq('closure_code', kodeClosure).limit(1); if (data && data.length > 0) return data[0].id; }
    const { data: f } = await supabase.from('closure_master').select('id').eq('site_id', siteId).limit(1);
    if (f && f.length > 0) return f[0].id;
    const { data: gf } = await supabase.from('closure_master').select('id').limit(1);
    return gf && gf.length > 0 ? gf[0].id : null;
}
async function getSpanId(kodeSpan, siteId) {
    if (kodeSpan) { const { data } = await supabase.from('span_master').select('id').eq('span_code', kodeSpan).limit(1); if (data && data.length > 0) return data[0].id; }
    const { data: f } = await supabase.from('span_master').select('id').eq('site_id', siteId).limit(1);
    if (f && f.length > 0) return f[0].id;
    const { data: gf } = await supabase.from('span_master').select('id').limit(1);
    return gf && gf.length > 0 ? gf[0].id : null;
}

app.post('/api/report/odc', upload.any(), async (req, res) => {
    res.status(400).json({ success: false, message: 'Gunakan upload direct SDK yang ada di index.html' });
});

// ─── [NEW] GET LIST DATA PER MODUL ───────────────────────────────────────────
// GET /api/data/:module?project_id=xxx
app.get('/api/data/:module', async (req, res) => {
    try {
        const { module } = req.params;
        const { project_id } = req.query;
        const tableName = MODULE_TABLE_MAP[module];
        if (!tableName) return res.status(400).json({ success: false, message: `Modul "${module}" tidak dikenali.` });
        if (!project_id) return res.status(400).json({ success: false, message: 'project_id wajib diisi.' });

        const { data, error } = await supabase
            .from(tableName)
            .select('*')
            .eq('period_id', project_id)
            .order('created_at', { ascending: false });

        if (error) throw new Error(error.message);

        // Ambil foto untuk semua record
        const ids = (data || []).map(r => r.id);
        let photoMap = {};
        if (ids.length > 0) {
            const { data: photos } = await supabase
                .from('photo_assets')
                .select('record_id, photo_kind, file_path, file_name')
                .in('record_id', ids);
            (photos || []).forEach(p => {
                if (!photoMap[p.record_id]) photoMap[p.record_id] = [];
                photoMap[p.record_id].push(p);
            });
        }

        // Ambil master data untuk lookup nama ODC/ODP/Closure/Span
        const [{ data: odcMaster }, { data: odpMaster }, { data: closureMaster }, { data: spanMaster }] = await Promise.all([
            supabase.from('odc_master').select('id, odc_code'),
            supabase.from('odp_master').select('id, odp_code'),
            supabase.from('closure_master').select('id, closure_code'),
            supabase.from('span_master').select('id, span_code'),
        ]);

        const enriched = (data || []).map(row => ({
            ...row,
            odc_code:     odcMaster?.find(o => o.id === row.odc_id)?.odc_code     || null,
            odp_code:     odpMaster?.find(o => o.id === row.odp_id)?.odp_code     || null,
            closure_code: closureMaster?.find(o => o.id === row.closure_id)?.closure_code || null,
            span_code:    spanMaster?.find(o => o.id === row.span_id)?.span_code   || null,
            photos:       photoMap[row.id] || [],
        }));

        res.json({ success: true, data: enriched });
    } catch (err) {
        console.error('[GET LIST]', err.message);
        res.status(500).json({ success: false, message: err.message });
    }
});

// ─── [NEW] GET DETAIL SINGLE RECORD ──────────────────────────────────────────
// GET /api/data/:module/:id
app.get('/api/data/:module/:id', async (req, res) => {
    try {
        const { module, id } = req.params;
        const tableName = MODULE_TABLE_MAP[module];
        if (!tableName) return res.status(400).json({ success: false, message: `Modul "${module}" tidak dikenali.` });

        const { data, error } = await supabase.from(tableName).select('*').eq('id', id).single();
        if (error || !data) return res.status(404).json({ success: false, message: 'Record tidak ditemukan.' });

        const { data: photos } = await supabase
            .from('photo_assets')
            .select('id, photo_kind, file_path, file_name')
            .eq('record_id', id);

        // Buat signed URL untuk setiap foto agar bisa ditampilkan di preview
        const photosWithUrl = await Promise.all((photos || []).map(async p => {
            const { data: signed } = await supabase.storage
                .from(STORAGE_BUCKET)
                .createSignedUrl(p.file_path, 3600);
            return { ...p, signed_url: signed?.signedUrl || null };
        }));

        res.json({ success: true, data: { ...data, photos: photosWithUrl } });
    } catch (err) {
        console.error('[GET DETAIL]', err.message);
        res.status(500).json({ success: false, message: err.message });
    }
});

// ─── [NEW] PUT UPDATE RECORD ──────────────────────────────────────────────────
// PUT /api/data/:module/:id
// Body: { fields: {...}, deletedPhotoIds: [...], newPhotos: [{photo_kind, file_path, file_name, file_size, uploaded_by}] }
app.put('/api/data/:module/:id', async (req, res) => {
    try {
        const { module, id } = req.params;
        const tableName = MODULE_TABLE_MAP[module];
        if (!tableName) return res.status(400).json({ success: false, message: `Modul "${module}" tidak dikenali.` });

        const { fields = {}, deletedPhotoIds = [], newPhotos = [] } = req.body;

        // Update field utama
        const { error: updateErr } = await supabase.from(tableName).update(fields).eq('id', id);
        if (updateErr) throw new Error(updateErr.message);

        // Hapus foto lama yang ditandai dihapus
        if (deletedPhotoIds.length > 0) {
            const { data: toDelete } = await supabase
                .from('photo_assets')
                .select('file_path')
                .in('id', deletedPhotoIds);

            // Hapus dari storage
            const filePaths = (toDelete || []).map(p => p.file_path);
            if (filePaths.length > 0) {
                await supabase.storage.from(STORAGE_BUCKET).remove(filePaths);
            }
            // Hapus dari tabel photo_assets
            await supabase.from('photo_assets').delete().in('id', deletedPhotoIds);
        }

        // Insert foto baru (sudah diupload dari client ke storage)
        if (newPhotos.length > 0) {
            const insertData = newPhotos.map(p => ({
                module_name: module,
                record_id:   id,
                photo_kind:  p.photo_kind,
                file_path:   p.file_path,
                file_name:   p.file_name,
                mime_type:   'image/jpeg',
                file_size:   p.file_size || 0,
                uploaded_by: p.uploaded_by || null,
            }));
            await supabase.from('photo_assets').insert(insertData);
        }

        res.json({ success: true, message: 'Data berhasil diperbarui.' });
    } catch (err) {
        console.error('[PUT UPDATE]', err.message);
        res.status(500).json({ success: false, message: err.message });
    }
});

// ─── [NEW] DELETE RECORD ──────────────────────────────────────────────────────
// DELETE /api/data/:module/:id
app.delete('/api/data/:module/:id', async (req, res) => {
    try {
        const { module, id } = req.params;
        const tableName = MODULE_TABLE_MAP[module];
        if (!tableName) return res.status(400).json({ success: false, message: `Modul "${module}" tidak dikenali.` });

        // Ambil semua foto terkait
        const { data: photos } = await supabase
            .from('photo_assets')
            .select('file_path')
            .eq('record_id', id);

        // Hapus dari storage
        const filePaths = (photos || []).map(p => p.file_path);
        if (filePaths.length > 0) {
            await supabase.storage.from(STORAGE_BUCKET).remove(filePaths);
        }

        // Hapus dari photo_assets
        await supabase.from('photo_assets').delete().eq('record_id', id);

        // Hapus record utama
        const { error: delErr } = await supabase.from(tableName).delete().eq('id', id);
        if (delErr) throw new Error(delErr.message);

        res.json({ success: true, message: 'Data berhasil dihapus.' });
    } catch (err) {
        console.error('[DELETE]', err.message);
        res.status(500).json({ success: false, message: err.message });
    }
});

// ─── ENDPOINT PREVIEW FOTO ─────────────────────────────────────────────────────
app.get('/api/report/preview', async (req, res) => {
    try {
        const { project_id } = req.query;
        if (!project_id) return res.status(400).send('<h1>Error: project_id wajib diisi</h1>');

        const { data: project } = await supabase.from('projects').select('project_name,site_code,bulan,tahun').eq('id', project_id).single();
        if (!project) return res.status(404).send('<h1>Project tidak ditemukan</h1>');

        const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
        const periodeStr = `${namaBulan[project.bulan - 1]} ${project.tahun}`;

        const tables = [
            { name: 'pm_odc',              label: 'ODC' },
            { name: 'pm_odp',              label: 'ODP' },
            { name: 'pm_closure',          label: 'Closure' },
            { name: 'pm_span',             label: 'Span Kabel' },
            { name: 'corrective_customer', label: 'Gangguan' },
            { name: 'dismantling_records', label: 'Dismantling' },
            { name: 'psb_records',         label: 'Aktivasi PSB' },
        ];

        const allRecordIds = [];
        const recordModuleMap = {};
        for (const t of tables) {
            const { data: rows } = await supabase.from(t.name).select('id').eq('period_id', project_id);
            (rows || []).forEach(r => { allRecordIds.push(r.id); recordModuleMap[r.id] = t.label; });
        }

        if (allRecordIds.length === 0) {
            return res.send(`<html><body style="font-family:Arial;padding:20px"><h2>Preview Foto — ${project.project_name}</h2><p>Belum ada data untuk project ini.</p></body></html>`);
        }

        const { data: photos } = await supabase
            .from('photo_assets')
            .select('record_id, photo_kind, file_path, file_name')
            .in('record_id', allRecordIds);

        if (!photos || photos.length === 0) {
            return res.send(`<html><body style="font-family:Arial;padding:20px"><h2>Preview Foto — ${project.project_name}</h2><p>Belum ada foto tersimpan untuk project ini.</p></body></html>`);
        }

        const photoRows = [];
        for (const p of photos) {
            const { data: signedData, error: signErr } = await supabase.storage
                .from(STORAGE_BUCKET)
                .createSignedUrl(p.file_path, 3600);

            photoRows.push({
                module:     recordModuleMap[p.record_id] || p.module_name || '-',
                photo_kind: p.photo_kind,
                file_name:  p.file_name,
                url:        signErr ? null : signedData?.signedUrl
            });
        }

        const cards = photoRows.map(p => `
            <div style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:12px;break-inside:avoid">
                <p style="margin:0 0 4px;font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.05em">${p.module} &mdash; ${p.photo_kind}</p>
                <p style="margin:0 0 8px;font-size:9px;color:#94a3b8">${p.file_name}</p>
                ${ p.url
                    ? `<img src="${p.url}" style="width:100%;border-radius:8px;object-fit:cover;max-height:200px" loading="lazy" onerror="this.outerHTML='<div style=background:#fee2e2;padding:12px;border-radius:8px;font-size:11px;color:#ef4444>&#x274C; Gagal load gambar</div>'" />`
                    : `<div style="background:#fee2e2;padding:12px;border-radius:8px;font-size:11px;color:#ef4444">&#x274C; Gagal buat URL</div>`
                }
            </div>`).join('');

        const html = `<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Preview Foto — ${project.project_name}</title>
<style>
  body { margin:0; background:#f1f5f9; font-family:Arial,sans-serif; padding:16px; }
  h1 { font-size:18px; font-weight:900; color:#1e293b; margin:0 0 2px; }
  p.sub { font-size:12px; color:#64748b; margin:0 0 16px; }
  .grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(200px,1fr)); gap:12px; }
  .badge { display:inline-block; background:#7c3aed; color:#fff; font-size:10px; font-weight:700; padding:3px 8px; border-radius:6px; margin-bottom:12px; }
  .btn { display:block; margin:0 auto 16px; background:#7c3aed; color:#fff; border:none; padding:12px 24px; border-radius:10px; font-weight:700; font-size:13px; cursor:pointer; text-decoration:none; text-align:center; }
</style>
</head>
<body>
<span class="badge">📊 Admin Preview</span>
<h1>${project.project_name}</h1>
<p class="sub">${project.site_code} &bull; ${periodeStr} &bull; ${photoRows.length} foto ditemukan</p>
<a class="btn" href="/api/report/export?project_id=${project_id}" target="_blank">⬇️ Download Excel</a>
<div class="grid">${cards}</div>
</body>
</html>`;

        res.send(html);

    } catch (err) {
        console.error('Preview error:', err);
        res.status(500).send(`<h1>Error</h1><p>${err.message}</p>`);
    }
});

// ─── ENDPOINT GENERATE EXCEL ──────────────────────────────────────────────────
app.get('/api/report/export', async (req, res) => {
    try {
        const { project_id } = req.query;
        if (!project_id) return res.status(400).send('<h1>Error: Parameter project_id wajib dikirim!</h1>');

        const { data: project, error: projErr } = await supabase.from('projects').select('*').eq('id', project_id).single();
        if (projErr || !project) return res.status(404).send('<h1>Error: Project tidak ditemukan!</h1>');

        const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
        const stringBulanText = `${namaBulan[project.bulan - 1]} ${project.tahun}`;

        const [{ data: odcMaster }, { data: odpMaster }, { data: closureMaster }, { data: spanMaster }] = await Promise.all([
            supabase.from('odc_master').select('*'),
            supabase.from('odp_master').select('*'),
            supabase.from('closure_master').select('*'),
            supabase.from('span_master').select('*')
        ]);

        const [{ data: pmOdc }, { data: pmOdp }, { data: pmClosure }, { data: pmSpan }, { data: corrective }, { data: dismantling }, { data: psb }] = await Promise.all([
            supabase.from('pm_odc').select('*').eq('period_id', project_id),
            supabase.from('pm_odp').select('*').eq('period_id', project_id),
            supabase.from('pm_closure').select('*').eq('period_id', project_id),
            supabase.from('pm_span').select('*').eq('period_id', project_id),
            supabase.from('corrective_customer').select('*').eq('period_id', project_id),
            supabase.from('dismantling_records').select('*').eq('period_id', project_id),
            supabase.from('psb_records').select('*').eq('period_id', project_id)
        ]);

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
                for (let c = startCol; c <= endCol; c++) wsRow.getCell(c).border = borderStyle;
            }
        }

        async function buildReportSheet({ sheetName, isCorrective = false, columnsSetup, headers_r5, headers_r6, mergeSpecs, rowsData, photoFields, photoRowHeight = 189 }) {
            const ws = workbook.addWorksheet(sheetName);
            ws.views = [{ showGridLines: true }];

            ws.getCell('A1').value = isCorrective ? 'Corrective Manage Service' : 'Preventive Manage service';
            ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true };
            ws.getRow(1).height = 23.5;
            ws.getCell('A2').value = `Stasiun : ${project.site_code.split(' - ')[1] || project.site_code}`;
            ws.getCell('A2').font = { name: 'Arial', size: 11, bold: true };
            ws.getRow(2).height = 23.5;
            ws.getCell('A3').value = `Periode : ${stringBulanText}`;
            ws.getCell('A3').font = { name: 'Arial', size: 11, bold: true };
            ws.getRow(3).height = 23.5;
            ws.getRow(4).height = 16;

            columnsSetup.widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

            const r5 = ws.getRow(5);
            r5.height = 20;
            headers_r5.forEach((val, i) => {
                const cell = r5.getCell(i + 1);
                cell.value = val; cell.font = boldFont; cell.fill = tableHeaderFill;
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });
            const r6 = ws.getRow(6);
            r6.height = 16;
            headers_r6.forEach((val, i) => {
                const cell = r6.getCell(i + 1);
                cell.value = val; cell.font = boldFont; cell.fill = tableHeaderFill;
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });
            mergeSpecs.forEach(m => { ws.mergeCells(m); });

            const downloadPromises = [];
            rowsData.forEach(r => {
                const assets = photoMap[r.id] || [];
                photoFields.forEach(pf => {
                    const asset = assets.find(a => {
                        const stored = (a.photo_kind || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                        const target = (pf.field    || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                        return stored === target;
                    });
                    if (asset) {
                        downloadPromises.push(
                            downloadAsJpegBuffer(asset.file_path).then(jpegBuf => {
                                if (!jpegBuf) return null;
                                return { recordId: r.id, field: pf.field, col: pf.col, buffer: jpegBuf };
                            })
                        );
                    }
                });
            });

            const downloadedAssets = await Promise.all(downloadPromises);
            const activePhotoMap = {};
            downloadedAssets.forEach(asset => {
                if (asset) activePhotoMap[`${asset.recordId}_${asset.field}`] = asset;
            });

            let curRow = 7;
            for (const r of rowsData) {
                const wsRow = ws.getRow(curRow);
                wsRow.height = photoFields.length > 0 ? photoRowHeight : 22;

                r.data.forEach((val, colIdx) => {
                    const cell = wsRow.getCell(colIdx + 1);
                    cell.value = val; cell.font = globalFont;
                    cell.alignment = { vertical: 'middle', horizontal: columnsSetup.alignments[colIdx] || 'left', wrapText: true };
                });

                for (const pf of photoFields) {
                    const key   = `${r.id}_${pf.field}`;
                    const asset = activePhotoMap[key];
                    const cell  = wsRow.getCell(pf.col);

                    if (asset && asset.buffer) {
                        try {
                            const imgId = workbook.addImage({ buffer: asset.buffer, extension: 'jpeg' });
                            ws.addImage(imgId, {
                                tl: { col: pf.col - 1, row: curRow - 1 },
                                br: { col: pf.col,     row: curRow     },
                                editAs: 'oneCell'
                            });
                            cell.value = '';
                        } catch (e) {
                            console.error(`[FAMIKA] Gagal insert gambar ke sel:`, e.message);
                            cell.value = 'Foto Gagal';
                            cell.font  = { name: 'Arial', size: 8, color: { argb: 'EF4444' } };
                        }
                    } else {
                        const hasAssetInDb = (photoMap[r.id] || []).some(a => {
                            const stored = (a.photo_kind || '').toLowerCase().replace(/^(photo_|foto_)/, '');
                            const target = (pf.field    || '').toLowerCase().replace(/^(photo_|foto_)/, '');
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

        await buildReportSheet({
            sheetName: 'ODC',
            photoRowHeight: 189,
            columnsSetup: { widths: [11.5,10,18.83,15.5,19.5,25.58,25.58,25.58,25.58,25.58,9.5], alignments: ['center','center','center','center','center','center','center','center','center','center','center'] },
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

        await buildReportSheet({
            sheetName: 'ODP',
            photoRowHeight: 189,
            columnsSetup: { widths: [11.33,14,11.5,24.58,24.58,12.75,19.5,25.58,25.58,25.58,25.58,25.58,9.91], alignments: ['center','center','center','center','center','center','center','center','center','center','center','center','center'] },
            headers_r5: ['NO','TANGGAL','ODC','ODP','SISA PORT','KONDISI','KEGIATAN','BEFORE','','AFTER','','HASIL OPM',''],
            headers_r6: ['','','','','','','','TUTUP','BUKA','TUTUP','BUKA','FOTO','REDAMAN'],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:I5','J5:K5','L5:M5'],
            rowsData: (pmOdp || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, odcMaster?.find(o => o.id === item.odc_id)?.odc_code || 'N/A', odpMaster?.find(o => o.id === item.odp_id)?.odp_code || 'N/A', item.sisa_port, item.kondisi, item.kegiatan, '','','','','', item.hasil_opm]
            })),
            photoFields: [
                { field: 'before_tutup', col: 8 },
                { field: 'before_buka',  col: 9 },
                { field: 'after_tutup',  col: 10 },
                { field: 'after_buka',   col: 11 },
                { field: 'foto_opm',     col: 12 }
            ]
        });

        const closureCfg = { widths: [11.5,10,18.83,12,19.5,25.58,25.58,25.58], alignments: ['center','center','center','center','left','center','center','center'] };
        await buildReportSheet({
            sheetName: 'CLOSURE',
            photoRowHeight: 189,
            columnsSetup: closureCfg,
            headers_r5: ['NO','TANGGAL','KODE CLOSURE','KONDISI','KEGIATAN','FOTO CLOSURE','FOTO SPARE KABEL','FOTO KESELURUHAN'],
            headers_r6: ['','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6'],
            rowsData: (pmClosure || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, closureMaster?.find(o => o.id === item.closure_id)?.closure_code || 'N/A', item.kondisi, item.kegiatan, '','','']
            })),
            photoFields: [{ field: 'foto_closure', col: 6 }]
        });

        await buildReportSheet({
            sheetName: 'KABEL',
            photoRowHeight: 189,
            columnsSetup: closureCfg,
            headers_r5: ['NO','TANGGAL','KODE SPAN','KONDISI','KEGIATAN','FOTO SPAN','FOTO SPARE KABEL','FOTO KESELURUHAN'],
            headers_r6: ['','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6'],
            rowsData: (pmSpan || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, spanMaster?.find(o => o.id === item.span_id)?.span_code || 'N/A', item.kondisi, item.kegiatan, '','','']
            })),
            photoFields: [{ field: 'foto_span', col: 6 }]
        });

        await buildReportSheet({
            sheetName: 'CUSTOMER',
            isCorrective: true,
            photoRowHeight: 100,
            columnsSetup: { widths: [6,12,18,12,15,15,8,20,15,16,16,25.58,25.58,10,10], alignments: ['center','center','left','center','center','center','center','left','center','center','center','center','center','center','center'] },
            headers_r5: ['NO','DATE','NAMA','ID','ODC','ODP','PORT','ACTION','MATERIAL','SN ONT','','BEFORE','AFTER','JAM',''],
            headers_r6: ['','','','','','','','','','NEW','OLD','','','MULAI','SELESAI'],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:K5','L5:L6','M5:M6','N5:O5'],
            rowsData: (corrective || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, item.customer_name, item.service_id, odcMaster?.find(o => o.id === item.odc_id)?.odc_code || '-', odpMaster?.find(o => o.id === item.odp_id)?.odp_code || '-', item.port_no || '-', item.action, item.material || '-', item.sn_ont_new || '-', item.sn_ont_old || '-', '','', item.jam_mulai || '-', item.jam_selesai || '-']
            })),
            photoFields: [
                { field: 'foto_before', col: 12 },
                { field: 'foto_after',  col: 13 }
            ]
        });

        await buildReportSheet({
            sheetName: 'DISMANTLING',
            photoRowHeight: 189,
            columnsSetup: { widths: [11.5,10,18,12,12,8,16,15,20,25.58,25.58], alignments: ['center','center','left','center','center','center','center','center','left','center','center'] },
            headers_r5: ['NO','TANGGAL','NAMA PELANGGAN','ID PELANGGAN','ODP','PORT','SN ONT','NO HP','ALAMAT','FOTO SERIALNUMBER','FOTO RUMAH'],
            headers_r6: ['','','','','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:J6','K5:K6'],
            rowsData: (dismantling || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, item.customer_name, item.service_id, odpMaster?.find(o => o.id === item.odp_id)?.odp_code || 'N/A', item.port_no || '-', item.sn_ont || '-', item.no_hp || '-', item.alamat || '-', '','']
            })),
            photoFields: [
                { field: 'foto_sn_ont', col: 10 },
                { field: 'foto_rumah',  col: 11 }
            ]
        });

        await buildReportSheet({
            sheetName: 'AKTIVASI',
            photoRowHeight: 189,
            columnsSetup: { widths: [11.5,10,18,12,12,8,16,15,20,25.58,25.58,25.58,25.58,25.58,25.58,25.58], alignments: ['center','center','left','center','center','center','center','center','left','center','center','center','center','center','center','center'] },
            headers_r5: ['NO','TANGGAL','NAMA PELANGGAN','ID PELANGGAN','ODP','PORT','SN ONT','NO HP','ALAMAT','FOTO ODP','PORT ODP','REDAM ODP','REDAM ONT','FOTO SN','FOTO RUMAH','SPEEDTEST'],
            headers_r6: ['','','','','','','','','','','','','','','',''],
            mergeSpecs:  ['A5:A6','B5:B6','C5:C6','D5:D6','E5:E6','F5:F6','G5:G6','H5:H6','I5:I6','J5:J6','K5:K6','L5:L6','M5:M6','N5:N6','O5:O6','P5:P6'],
            rowsData: (psb || []).map((item, idx) => ({
                id: item.id,
                data: [idx+1, item.tanggal, item.customer_name, item.service_id, odpMaster?.find(o => o.id === item.odp_id)?.odp_code || 'N/A', item.port_no, item.sn_ont, item.no_hp || '-', item.alamat || '-', '','','','','','','']
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
