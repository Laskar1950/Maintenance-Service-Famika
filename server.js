const express = require('express');
const multer = require('multer');
const sharp = require('sharp');
const { createClient } = require('@supabase/supabase-js');

const app = express();

// Konfigurasi parser file menggunakan penyimpanan memori (buffer)
// Menggunakan upload.any() untuk menerima field foto secara fleksibel dari berbagai jenis form pekerjaan
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
    // Fallback pertama ke user aktif terdaftar
    const { data: fallback } = await supabase
        .from('users')
        .select('id')
        .eq('is_active', true)
        .limit(1);
    if (fallback && fallback.length > 0) return fallback[0].id;

    // Fallback akhir ke entitas user pertama yang ada di DB
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
    
    // Kompresi Sharp progresif dengan kualitas 72%
    const compressedBuffer = await sharp(fileBuffer)
        .resize({ width: 1080, height: 1080, fit: 'inside', withoutEnlargement: true })
        .jpeg({ quality: 72, progressive: true })
        .toBuffer();

    // Diunggah ke bucket pribadi 'maintenance-photos' sesuai konfigurasi RLS & Storage SQL
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

// --- ENDPOINT TUNGGAL PENERIMA MULTIPART FORM DATA ---
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

        // 1. Identifikasi Site ID yang terikat pada Maintenance Period
        const { data: periodData, error: periodError } = await supabase
            .from('maintenance_periods')
            .select('site_id')
            .eq('id', project_id)
            .single();

        if (periodError || !periodData) {
            return res.status(404).json({ success: false, message: 'Periode pemeliharaan tidak ditemukan di database.' });
        }
        const siteId = periodData.site_id;

        // 2. Lookup entitas ID Teknisi dari tabel public.users
        const technicianId = await getTechnicianId(technician_username);
        if (!technicianId) {
            return res.status(400).json({ success: false, message: 'User teknisi belum terdaftar di database.' });
        }

        // 3. Klasifikasi Mapping Kolom & Nama Tabel Sesuai SQL
        let tableName = '';
        let moduleName = '';
        const insertData = {
            period_id: project_id,
            site_id: siteId,
            technician_id: technicianId,
            tanggal: new Date().toISOString().split('T')[0], // format tanggal YYYY-MM-DD
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
                // Closure memerlukan ODC id rujukan
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

        // 4. Masukkan Data Transaksi Utama ke Tabel Rujukan (Dapatkan Kembali ID Baris)
        const { data: recordRow, error: insertError } = await supabase
            .from(tableName)
            .insert([insertData])
            .select('id')
            .single();

        if (insertError || !recordRow) {
            throw new Error(`Gagal menyimpan data transaksi ke tabel ${tableName}: ${insertError?.message}`);
        }

        const insertedRecordId = recordRow.id;

        // 5. Unggah Gambar secara Iteratif dan Catat ke Tabel Terpusat 'photo_assets'
        const photoAssetInserts = [];
        for (const file of files) {
            const uploadResult = await compressAndUpload(
                file.buffer, 
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

        // Masukkan semua log data media ke tabel central 'photo_assets' jika ada foto
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

module.exports = app;
