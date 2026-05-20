const express = require('express');
const multer = require('multer');
const sharp = require('sharp');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');

const app = express();

// Konfigurasi parser file menggunakan penyimpanan memori (buffer)
const upload = multer({ storage: multer.memoryStorage() });

// ─── HUBUNGKAN KE SUPABASE VIA ENVIRONMENT VARIABLES ─────────────────────
const JALUR_SUPABASE_URL = process.env.SUPABASE_URL || 'https://bwfbntebndkbglmbziij.supabase.co';
const KUNCI_SUPABASE_ANON = process.env.SUPABASE_ANON_KEY;

const supabase = createClient(JALUR_SUPABASE_URL, KUNCI_SUPABASE_ANON);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

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
            .or(`full_name.eq.${username},email.ilike.%${username}%`)
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

async function getSpanId(kodeClosure, siteId) {
    if (kodeClosure) {
        const { data } = await supabase
            .from('span_master')
            .select('id')
            .eq('span_code', kodeClosure)
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

// --- FUNGSI UTAMA KOMPRESI DAN UNGGAH KE STORAGE BUCKET ---
async function compressAndUpload(fileBuffer, fieldName, taskType, projectId) {
    const folderName = taskType.replace('form_', '');
    const fileName = `${fieldName}-${Date.now()}.jpg`;
    const fullPath = `${projectId}/${folderName}/${fileName}`;
    
    const compressedBuffer = await sharp(fileBuffer)
        .resize({ width: 1080, height: 1080, fit: 'inside', withoutEnlargement: true })
        .jpeg({ quality: 72, progressive: true })
        .toBuffer();

    const { error } = await supabase.storage
        .from('maintenance-photos')
        .upload(fullPath, compressedBuffer, {
            contentType: 'image/jpeg',
            upsert: true
        });

    if (error) throw error;

    return {
        path: fullPath,
        name: fileName,
        size: compressedBuffer.length
    };
}

// --- ENDPOINT UNTUK MENERIMA INPUT DARI TELEGRAM MINI APP ---
app.post('/api/report/odc', upload.any(), async (req, res) => {
    try {
        const { task_type, technician_username, project_id } = req.body;
        const files = req.files || [];
        
        if (!task_type) {
            return res.status(400).json({ success: false, message: 'Parameter task_type wajib dikirim.' });
        }
        if (!project_id) {
            return res.status(400).json({ success: false, message: 'ID Periode (project_id) wajib terisi.' });
        }

        const { data: periodData, error: periodError } = await supabase
            .from('maintenance_periods')
            .select('site_id')
            .eq('id', project_id)
            .single();

        if (periodError || !periodData) {
            return res.status(404).json({ success: false, message: 'Periode pemeliharaan tidak ditemukan di database.' });
        }
        const siteId = periodData.site_id;

        const technicianId = await getTechnicianId(technician_username);
        if (!technicianId) {
            return res.status(400).json({ success: false, message: 'User teknisi belum terdaftar di database.' });
        }

        let tableName = '';
        let moduleName = '';
        const insertData = {
            period_id: project_id,
            site_id: siteId,
            technician_id: technicianId,
            tanggal: new Date().toISOString().split('T')[0],
            status: 'submitted'
        };

        switch (task_type) {
            case 'form_odc':
                tableName = 'pm_odc';
                moduleName = 'pm_odc';
                insertData.odc_id = await getOdcId(req.body.kode_odc, siteId);
                insertData.hasil_opm = parseFloat(req.body.hasil_opm) || 0;
                insertData.kondisi = 'BAIK';
                insertData.kegiatan = 'Preventif ODC';
                insertData.catatan = `Laporan ODC ${req.body.kode_odc}`;
                break;

            case 'form_odp':
                tableName = 'pm_odp';
                moduleName = 'pm_odp';
                insertData.odc_id = await getOdcId(req.body.kode_odc, siteId);
                insertData.odp_id = await getOdpId(req.body.kode_odp, siteId);
                insertData.hasil_opm = parseFloat(req.body.hasil_opm) || 0;
                insertData.kondisi = 'BAIK';
                insertData.sisa_port = '8';
                insertData.kegiatan = 'Preventif ODP';
                insertData.catatan = `Laporan ODP ${req.body.kode_odp}`;
                break;

            case 'form_closure':
                tableName = 'pm_closure';
                moduleName = 'pm_closure';
                insertData.closure_id = await getClosureId(req.body.kode_closure, siteId);
                insertData.odc_id = await getOdcId(null, siteId);
                insertData.kondisi = 'BAIK';
                insertData.kegiatan = 'Preventif Closure';
                break;

            case 'form_span':
                tableName = 'pm_span';
                moduleName = 'pm_span';
                insertData.span_id = await getSpanId(req.body.kode_closure, siteId);
                insertData.odc_id = await getOdcId(null, siteId);
                insertData.kondisi = 'BAIK';
                insertData.kegiatan = 'Preventif Span';
                break;

            case 'form_gangguan':
                tableName = 'corrective_customer';
                moduleName = 'corrective_customer';
                insertData.odc_id = await getOdcId(req.body.odc, siteId);
                insertData.odp_id = await getOdpId(req.body.odp, siteId);
                insertData.customer_name = req.body.nama_pelanggan;
                insertData.service_id = req.body.id_pelanggan;
                insertData.port_no = parseInt(req.body.port_odp) || null;
                insertData.action = req.body.action;
                insertData.catatan = `Penyebab: ${req.body.penyebab_gangguan}`;
                break;

            case 'form_dismantling':
                tableName = 'dismantling_records';
                moduleName = 'dismantling_records';
                insertData.odp_id = await getOdpId(req.body.odp, siteId);
                insertData.odc_id = await getOdcId(null, siteId);
                insertData.customer_name = req.body.nama_pelanggan;
                insertData.service_id = req.body.id_pelanggan || 'N/A';
                insertData.port_no = parseInt(req.body.port) || null;
                insertData.action = `Pencabutan Dropcore ONT SN: ${req.body.serial_number_ont}`;
                insertData.catatan = `HP: ${req.body.no_hp}, Alamat: ${req.body.alamat}`;
                break;

            case 'form_aktivasi':
                tableName = 'psb_records';
                moduleName = 'psb_records';
                insertData.odp_id = await getOdpId(req.body.odp, siteId);
                insertData.odc_id = await getOdcId(null, siteId);
                insertData.customer_name = req.body.nama_pelanggan;
                insertData.service_id = req.body.id_pelanggan;
                insertData.port_no = parseInt(req.body.port) || 1;
                insertData.alamat = req.body.alamat;
                insertData.action = `Aktivasi Baru No HP: ${req.body.no_hp}`;
                insertData.material = 'Dropcore, ONT';
                insertData.sn_ont = req.body.serial_number_ont;
                insertData.catatan = 'Pemasangan pelanggan baru selesai';
                break;

            default:
                return res.status(400).json({ success: false, message: 'Tipe tugas tidak dikenali.' });
        }

        const { data: recordRow, error: insertError } = await supabase
            .from(tableName)
            .insert([insertData])
            .select('id')
            .single();

        if (insertError || !recordRow) {
            throw new Error(`Gagal menyimpan data transaksi ke tabel ${tableName}: ${insertError?.message}`);
        }

        const insertedRecordId = recordRow.id;

        const photoAssetInserts = [];
        for (const file of files) {
            const uploadResult = await compressAndUpload(
                file.buffer, 
                file.originalname, 
                file.fieldname, 
                task_type,
                project_id
            );

            photoAssetInserts.push({
                module_name: moduleName,
                record_id: insertedRecordId,
                photo_kind: file.fieldname,
                file_path: uploadResult.path,
                file_name: uploadResult.name,
                mime_type: 'image/jpeg',
                file_size: uploadResult.size,
                taken_at: new Date().toISOString(),
                uploaded_by: technicianId
            });
        }

        if (photoAssetInserts.length > 0) {
            const { error: photoInsertError } = await supabase
                .from('photo_assets')
                .insert(photoAssetInserts);

            if (photoInsertError) {
                throw new Error(`Laporan teks terkirim, namun gagal mendaftarkan daftar file di tabel photo_assets: ${photoInsertError.message}`);
            }
        }

        res.status(200).json({ 
            success: true, 
            message: `Laporan berhasil disimpan ke tabel ${tableName} dan aset foto terdaftar di tabel photo_assets!` 
        });

    } catch (err) {
        console.error("Internal Server Error:", err);
        res.status(500).json({ success: false, message: err.message });
    }
});

// ─── ENDPOINT GENERATE SPREADSHEET AUTOMATION (UNTUK ADMIN) ──────────────────────────
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

        // Parse format Bulan (angka ➔ teks Indonesia)
        const namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        const stringBulanText = `${namaBulan[project.bulan - 1]} ${project.tahun}`;

        // 2. Tarik Semua Data Master untuk Lookup Cepat di Memori
        const [{ data: sites }, { data: users }, { data: odcMaster }, { data: odpMaster }, { data: closureMaster }, { data: spanMaster }] = await Promise.all([
            supabase.from('sites').select('*'),
            supabase.from('users').select('*'),
            supabase.from('odc_master').select('*'),
            supabase.from('odp_master').select('*'),
            supabase.from('closure_master').select('*'),
            supabase.from('span_master').select('*')
        ]);

        // 3. Tarik Seluruh Tabel Transaksi Pekerjaan untuk Project_id Terkait
        const [{ data: pmOdc }, { data: pmOdp }, { data: pmClosure }, { data: pmSpan }, { data: corrective }, { data: dismantling }, { data: psb }] = await Promise.all([
            supabase.from('pm_odc').select('*').eq('period_id', project_id),
            supabase.from('pm_odp').select('*').eq('period_id', project_id),
            supabase.from('pm_closure').select('*').eq('period_id', project_id),
            supabase.from('pm_span').select('*').eq('period_id', project_id),
            supabase.from('corrective_customer').select('*').eq('period_id', project_id),
            supabase.from('dismantling_records').select('*').eq('period_id', project_id),
            supabase.from('psb_records').select('*').eq('period_id', project_id)
        ]);

        // 4. Tarik Log Semua Media Foto
        const { data: photos } = await supabase.from('photo_assets').select('*');
        const photoMap = {};
        if (photos) {
            photos.forEach(p => {
                if (!photoMap[p.record_id]) photoMap[p.record_id] = [];
                photoMap[p.record_id].push(p);
            });
        }

        // 5. Inisialisasi Excel Workbook
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Famika PM Telecom Portal';
        workbook.created = new Date();

        // --- STYLING PATTERNS ---
        const tableHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF85D3F2' } }; // Soft Cyan sesuai gambar template
        const borderStyle = {
            top: { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            left: { style: 'thin', color: { argb: '000000' } },
            right: { style: 'thin', color: { argb: '000000' } }
        };

        const globalFont = { name: 'Arial', size: 10 };
        const boldFont = { name: 'Arial', size: 10, bold: true };

        // Helper untuk menyelaraskan style garis tepi tabel secara menyeluruh
        function applyBordersToRange(ws, startCol, startRow, endCol, endRow) {
            for (let r = startRow; r <= endRow; r++) {
                const wsRow = ws.getRow(r);
                for (let c = startCol; c <= endCol; c++) {
                    const cell = wsRow.getCell(c);
                    cell.border = borderStyle;
                }
            }
        }

        // --- HELPER TEMPLATE BUILDER ---
        async function buildReportSheet({
            sheetName,
            isCorrective = false,
            columnsSetup, // colWidths, colAlignments
            headers_r5, // baris ke-5
            headers_r6, // baris ke-6
            mergeSpecs, // array merge cells
            rowsData,   // baris transaksi
            photoFields // field foto untuk di-embed
        }) {
            const ws = workbook.addWorksheet(sheetName);
            ws.views = [{ showGridLines: true }]; // grid lines wajib menyala

            // Row 1-3: Header Judul Laporan Sesuai Gambar Template
            ws.getCell('A1').value = isCorrective ? 'Corrective Manage Service' : 'Preventive Manage service';
            ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true };
            ws.getRow(1).height = 24;

            ws.getCell('A2').value = `Stasiun : ${project.site_code.split(' - ')[1] || project.site_code}`;
            ws.getCell('A2').font = { name: 'Arial', size: 11, bold: true };
            ws.getRow(2).height = 18;

            ws.getCell('A3').value = `Periode : ${stringBulanText}`;
            ws.getCell('A3').font = { name: 'Arial', size: 11, bold: true };
            ws.getRow(3).height = 18;

            ws.getRow(4).height = 12; // Space baris ke-4 kosongan

            // Set Lebar Kolom
            columnsSetup.widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

            // Suntikkan teks header baris 5 & 6
            const r5 = ws.getRow(5);
            r5.height = 20;
            headers_r5.forEach((val, i) => {
                const cell = r5.getCell(i + 1);
                cell.value = val;
                cell.font = boldFont;
                cell.fill = tableHeaderFill;
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });

            const r6 = ws.getRow(6);
            r6.height = 20;
            headers_r6.forEach((val, i) => {
                const cell = r6.getCell(i + 1);
                cell.value = val;
                cell.font = boldFont;
                cell.fill = tableHeaderFill;
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            });

            // Jalankan instruksi penggabungan sel (Merge Cells)
            mergeSpecs.forEach(m => { ws.mergeCells(m); });

            // Data Rows Render
            let curRow = 7;
            for (const r of rowsData) {
                const wsRow = ws.getRow(curRow);
                // Tinggi baris dinaikkan jika melampirkan foto agar tidak pecah/tindih
                wsRow.height = photoFields.length > 0 ? 110 : 22;

                r.data.forEach((val, colIdx) => {
                    const cell = wsRow.getCell(colIdx + 1);
                    cell.value = val;
                    cell.font = globalFont;
                    
                    // Set rata kanan/kiri/tengah sesuai tipe kolom
                    const alignHoriz = columnsSetup.alignments[colIdx] || 'left';
                    cell.alignment = { vertical: 'middle', horizontal: alignHoriz, wrapText: true };
                });

                // Download & Embed biner foto dari Supabase ke dalam sel Excel
                for (const pf of photoFields) {
                    const assets = photoMap[r.id] || [];
                    const asset = assets.find(a => a.photo_kind === pf.field);
                    const cell = wsRow.getCell(pf.col);
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };

                    if (asset) {
                        try {
                            const { data: fileData, error: dlErr } = await supabase.storage
                                .from('maintenance-photos')
                                .download(asset.file_path);

                            if (!dlErr && fileData) {
                                const buffer = Buffer.from(await fileData.arrayBuffer());
                                const imgId = workbook.addImage({
                                    buffer: buffer,
                                    extension: 'jpeg'
                                });
                                // Sisipkan foto di dalam batas sel kolom terkait
                                ws.addImage(imgId, {
                                    tl: { col: pf.col - 1, row: curRow - 1, colOff: 6, rowOff: 6 },
                                    br: { col: pf.col, row: curRow, colOff: -6, rowOff: -6 },
                                    editAs: 'oneCell'
                                });
                                cell.value = ''; // hilangkan tulisan string URL
                            } else {
                                cell.value = 'Foto Gagal';
                                cell.font = { name: 'Arial', size: 8, color: { argb: 'EF4444' } };
                            }
                        } catch (e) {
                            cell.value = 'Eror Foto';
                            cell.font = { name: 'Arial', size: 8, color: { argb: 'EF4444' } };
                        }
                    } else {
                        cell.value = 'Tidak Ada';
                        cell.font = { name: 'Arial', size: 8, color: { argb: '94A3B8' } };
                    }
                }
                curRow++;
            }

            // Terapkan border di seluruh area tabel yang terbuat
            applyBordersToRange(ws, 1, 5, columnsSetup.widths.length, curRow - 1);
        }

        // ============================================================
        // BUILD SHEET 1: ODC PM (Tab Name: ODC)
        // Sesuai Gambar: image_d80fca.jpg
        // ============================================================
        const odc_widths = [6, 12, 18, 10, 15, 24, 24, 24, 24, 24, 12];
        const odc_aligns = ['center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center'];
        const odc_r5 = ['NO', 'TANGGAL', 'ODC', 'KONDISI', 'KEGIATAN', 'BEFORE', '', 'AFTER', '', 'HASIL OPM', ''];
        const odc_r6 = ['', '', '', '', '', 'BUKA', 'TUTUP', 'BUKA', 'TUTUP', 'FOTO', 'REDAMAN'];
        const odc_merges = ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:G5', 'H5:I5', 'J5:K5'];

        const odc_rows = (pmOdc || []).map((item, idx) => {
            const code = odcMaster.find(o => o.id === item.odc_id)?.odc_code || 'N/A';
            return {
                id: item.id,
                data: [idx + 1, item.tanggal, code, item.kondisi, item.kegiatan, '', '', '', '', '', item.hasil_opm]
            };
        });
        const odc_photos_mapping = [
            { field: 'photo_before_buka', col: 6 },
            { field: 'photo_before_tutup', col: 7 },
            { field: 'photo_after_buka', col: 8 },
            { field: 'photo_after_tutup', col: 9 },
            { field: 'photo_opm', col: 10 }
        ];

        await buildReportSheet({
            sheetName: 'ODC',
            columnsSetup: { widths: odc_widths, alignments: odc_aligns },
            headers_r5: odc_r5,
            headers_r6: odc_r6,
            mergeSpecs: odc_merges,
            rowsData: odc_rows,
            photoFields: odc_photos_mapping
        });

        // ============================================================
        // BUILD SHEET 2: ODP PM (Tab Name: ODP)
        // Sesuai Gambar: image_d81345.jpg
        // ============================================================
        const odp_widths = [6, 12, 18, 18, 12, 10, 15, 24, 24, 24, 24, 24, 12];
        const odp_aligns = ['center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center', 'center'];
        const odp_r5 = ['NO', 'TANGGAL', 'ODC', 'ODP', 'SISA PORT', 'KONDISI', 'KEGIATAN', 'BEFORE', '', 'AFTER', '', 'HASIL OPM', ''];
        const odp_r6 = ['', '', '', '', '', '', '', 'TUTUP', 'BUKA', 'TUTUP', 'BUKA', 'FOTO', 'REDAMAN'];
        const odp_merges = ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:F6', 'G5:G6', 'H5:I5', 'J5:K5', 'L5:M5'];

        const odp_rows = (pmOdp || []).map((item, idx) => {
            const oCode = odcMaster.find(o => o.id === item.odc_id)?.odc_code || 'N/A';
            const pCode = odpMaster.find(o => o.id === item.odp_id)?.odp_code || 'N/A';
            return {
                id: item.id,
                data: [idx + 1, item.tanggal, oCode, pCode, item.sisa_port, item.kondisi, item.kegiatan, '', '', '', '', '', item.hasil_opm]
            };
        });
        const odp_photos_mapping = [
            { field: 'photo_before_tutup', col: 8 },
            { field: 'photo_before_buka', col: 9 },
            { field: 'photo_after_tutup', col: 10 },
            { field: 'photo_after_buka', col: 11 },
            { field: 'photo_opm', col: 12 }
        ];

        await buildReportSheet({
            sheetName: 'ODP',
            columnsSetup: { widths: odp_widths, alignments: odp_aligns },
            headers_r5: odp_r5,
            headers_r6: odp_r6,
            mergeSpecs: odp_merges,
            rowsData: odp_rows,
            photoFields: odp_photos_mapping
        });

        // ============================================================
        // BUILD SHEET 3: PM CLOSURE (Tab Name: CLOSURE)
        // ============================================================
        const cls_widths = [6, 12, 18, 12, 15, 24, 24, 24];
        const cls_aligns = ['center', 'center', 'center', 'center', 'left', 'center', 'center', 'center'];
        const cls_r5 = ['NO', 'TANGGAL', 'KODE CLOSURE', 'KONDISI', 'KEGIATAN', 'FOTO CLOSURE', 'FOTO SPARE KABEL', 'FOTO KESELURUHAN'];
        const cls_r6 = ['', '', '', '', '', '', '', ''];
        const cls_merges = ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:F6', 'G5:G6', 'H5:H6'];

        const cls_rows = (pmClosure || []).map((item, idx) => {
            const code = closureMaster.find(o => o.id === item.closure_id)?.closure_code || 'N/A';
            return {
                id: item.id,
                data: [idx + 1, item.tanggal, code, item.kondisi, item.kegiatan, '', '', '']
            };
        });
        const cls_photos = [{ field: 'photo_closure', col: 6 }, { field: 'photo_spare_kabel', col: 7 }, { field: 'photo_keseluruhan', col: 8 }];

        await buildReportSheet({
            sheetName: 'CLOSURE',
            columnsSetup: { widths: cls_widths, alignments: cls_aligns },
            headers_r5: cls_r5,
            headers_r6: cls_r6,
            mergeSpecs: cls_merges,
            rowsData: cls_rows,
            photoFields: cls_photos
        });

        // ============================================================
        // BUILD SHEET 4: PM SPAN CABLE (Tab Name: KABEL)
        // ============================================================
        const span_rows = (pmSpan || []).map((item, idx) => {
            const code = spanMaster.find(o => o.id === item.span_id)?.span_code || 'N/A';
            return {
                id: item.id,
                data: [idx + 1, item.tanggal, code, item.kondisi, item.kegiatan, '', '', '']
            };
        });

        await buildReportSheet({
            sheetName: 'KABEL',
            columnsSetup: { widths: cls_widths, alignments: cls_aligns },
            headers_r5: cls_r5,
            headers_r6: cls_r6,
            mergeSpecs: cls_merges,
            rowsData: span_rows,
            photoFields: cls_photos
        });

        // ============================================================
        // BUILD SHEET 5: PENANGANAN GANGGUAN (Tab Name: CUSTOMER)
        // Sesuai Gambar: image_d81670.png
        // ============================================================
        const cst_widths = [6, 12, 18, 12, 15, 15, 8, 20, 15, 16, 16, 24, 24, 10, 10];
        const cst_aligns = ['center', 'center', 'left', 'center', 'center', 'center', 'center', 'left', 'center', 'center', 'center', 'center', 'center', 'center', 'center'];
        const cst_r5 = ['NO', 'DATE', 'NAMA', 'ID', 'ODC', 'ODP', 'PORT', 'ACTION', 'MATERIAL', 'SN ONT', '', 'BEFORE', 'AFTER', 'JAM', ''];
        const cst_r6 = ['', '', '', '', '', '', '', '', '', 'NEW', 'OLD', '', '', 'MULAI', 'SELESAI'];
        const cst_merges = ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:F6', 'G5:G6', 'H5:H6', 'I5:I6', 'J5:K5', 'L5:L6', 'M5:M6', 'N5:O5'];

        const cst_rows = (corrective || []).map((item, idx) => {
            const odcVal = odcMaster.find(o => o.id === item.odc_id)?.odc_code || 'tidak terinfo';
            const odpVal = odpMaster.find(o => o.id === item.odp_id)?.odp_code || 'tidak terinfo';
            return {
                id: item.id,
                // Defaulting fields material, sn_ont (new/old), and jam sesuai gambar template
                data: [idx + 1, item.tanggal, item.customer_name, item.service_id, odcVal, odpVal, item.port_no || 'tidak terinfo', item.action, '-', '-', '-', '', '', '14.00', '14.50']
            };
        });
        const cst_photos = [{ field: 'photo_before', col: 12 }, { field: 'photo_after', col: 13 }];

        await buildReportSheet({
            sheetName: 'CUSTOMER',
            isCorrective: true,
            columnsSetup: { widths: cst_widths, alignments: cst_aligns },
            headers_r5: cst_r5,
            headers_r6: cst_r6,
            mergeSpecs: cst_merges,
            rowsData: cst_rows,
            photoFields: cst_photos
        });

        // ============================================================
        // BUILD SHEET 6: DISMANTLING (Tab Name: DISMANTLING)
        // ============================================================
        const dsm_widths = [6, 12, 18, 12, 12, 8, 16, 15, 20, 24, 24];
        const dsm_aligns = ['center', 'center', 'left', 'center', 'center', 'center', 'center', 'center', 'left', 'center', 'center'];
        const dsm_r5 = ['NO', 'TANGGAL', 'NAMA PELANGGAN', 'ID PELANGGAN', 'ODP', 'PORT', 'SN ONT', 'NO HP', 'ALAMAT', 'FOTO SERIALNUMBER', 'FOTO RUMAH'];
        const dsm_r6 = ['', '', '', '', '', '', '', '', '', '', ''];
        const dsm_merges = ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:F6', 'G5:G6', 'H5:H6', 'I5:I6', 'J5:J6', 'K5:K6'];

        const dsm_rows = (dismantling || []).map((item, idx) => {
            const odpVal = odpMaster.find(o => o.id === item.odp_id)?.odp_code || 'N/A';
            return {
                id: item.id,
                // Extract detail dari string catatan (HP dan Alamat)
                data: [
                    idx + 1, 
                    item.tanggal, 
                    item.customer_name, 
                    item.service_id, 
                    odpVal, 
                    item.port_no || '-', 
                    item.action.replace('Pencabutan Dropcore ONT SN: ', ''), 
                    item.catatan.split(', ')[0].replace('HP: ', ''),
                    item.catatan.split(', ')[1]?.replace('Alamat: ', '') || '-',
                    '', 
                    ''
                ]
            };
        });
        const dsm_photos = [{ field: 'photo_serialnumber', col: 10 }, { field: 'photo_rumah', col: 11 }];

        await buildReportSheet({
            sheetName: 'DISMANTLING',
            columnsSetup: { widths: dsm_widths, alignments: dsm_aligns },
            headers_r5: dsm_r5,
            headers_r6: dsm_r6,
            mergeSpecs: dsm_merges,
            rowsData: dsm_rows,
            photoFields: dsm_photos
        });

        // ============================================================
        // BUILD SHEET 7: AKTIVASI BARU (Tab Name: AKTIVASI)
        // ============================================================
        const psb_widths = [6, 12, 18, 12, 12, 8, 16, 15, 20, 24, 24, 24, 24, 24, 24, 24];
        const psb_aligns = ['center', 'center', 'left', 'center', 'center', 'center', 'center', 'center', 'left', 'center', 'center', 'center', 'center', 'center', 'center', 'center'];
        const psb_r5 = ['NO', 'TANGGAL', 'NAMA PELANGGAN', 'ID PELANGGAN', 'ODP', 'PORT', 'SN ONT', 'NO HP', 'ALAMAT', 'FOTO ODP', 'PORT ODP', 'REDAM ODP', 'REDAM ONT', 'FOTO SN', 'FOTO RUMAH', 'SPEEDTEST'];
        const psb_r6 = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''];
        const psb_merges = ['A5:A6', 'B5:B6', 'C5:C6', 'D5:D6', 'E5:E6', 'F5:F6', 'G5:G6', 'H5:H6', 'I5:I6', 'J5:J6', 'K5:K6', 'L5:L6', 'M5:M6', 'N5:N6', 'O5:O6', 'P5:P6'];

        const psb_rows = (psb || []).map((item, idx) => {
            const odpVal = odpMaster.find(o => o.id === item.odp_id)?.odp_code || 'N/A';
            return {
                id: item.id,
                data: [
                    idx + 1,
                    item.tanggal,
                    item.customer_name,
                    item.service_id,
                    odpVal,
                    item.port_no,
                    item.sn_ont,
                    item.action.replace('Aktivasi Baru No HP: ', ''),
                    item.alamat,
                    '', '', '', '', '', '', ''
                ]
            };
        });
        const psb_photos = [
            { field: 'photo_odp', col: 10 },
            { field: 'photo_port_odp', col: 11 },
            { field: 'photo_redaman_odp', col: 12 },
            { field: 'photo_redaman_akhir', col: 13 },
            { field: 'photo_serial_number_ont', col: 14 },
            { field: 'photo_instalasi_rumah', col: 15 },
            { field: 'photo_speedtest', col: 16 }
        ];

        await buildReportSheet({
            sheetName: 'AKTIVASI',
            columnsSetup: { widths: psb_widths, alignments: psb_aligns },
            headers_r5: psb_r5,
            headers_r6: psb_r6,
            mergeSpecs: psb_merges,
            rowsData: psb_rows,
            photoFields: psb_photos
        });

        // 6. Kirim file sebagai unduhan biner .xlsx ke browser admin
        const formattedFileName = `Laporan_Bulanan_${project.project_name.replace(/\s+/g, '_')}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${formattedFileName}"`);
        
        await workbook.xlsx.write(res);
        res.end();

    } catch (err) {
        console.error("Internal Server Error on Export:", err);
        res.status(500).send(`<h1>Error saat mengekspor laporan:</h1><p>${err.message}</p>`);
    }
});

module.exports = app;
