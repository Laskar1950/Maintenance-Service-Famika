require('dotenv').config();
const express = require('express');
const multer = require('multer');
const sharp = require('sharp');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

async function compressAndUpload(fileBuffer, originalName, fieldName) {
    const fileName = `${fieldName}-${Date.now()}-${originalName.replace(/\s+/g, '_')}.jpg`;
    
    const compressedBuffer = await sharp(fileBuffer)
        .resize({ width: 1080, height: 1080, fit: 'inside', withoutEnlargement: true })
        .jpeg({ quality: 72, progressive: true })
        .toBuffer();

    const { data, error } = await supabase.storage
        .from('pm-photos')
        .upload(`odc/${fileName}`, compressedBuffer, {
            contentType: 'image/jpeg',
            upsert: true
        });

    if (error) throw error;

    const { data: { publicUrl } } = supabase.storage.from('pm-photos').getPublicUrl(`odc/${fileName}`);
    return publicUrl;
}

app.post('/api/report/odc', upload.fields([
    { name: 'before_tutup', maxCount: 1 },
    { name: 'before_buka', maxCount: 1 },
    { name: 'after_tutup', maxCount: 1 },
    { name: 'after_buka', maxCount: 1 },
    { name: 'photo_opm', maxCount: 1 }
]), async (req, res) => {
    try {
        const { site_id, odc_id, hasil_opm, technician_username } = req.body;
        const files = req.files;
        
        const photoUrls = {};
        const fieldNames = ['before_tutup', 'before_buka', 'after_tutup', 'after_buka', 'photo_opm'];

        for (const field of fieldNames) {
            if (files[field] && files[field][0]) {
                const file = files[field][0];
                photoUrls[field] = await compressAndUpload(file.buffer, file.originalname, field);
            }
        }

        const { data, error } = await supabase
            .from('tx_pm_odc')
            .insert([{
                technician_username: technician_username || 'Teknisi Lapangan',
                site_id: parseInt(site_id),
                odc_id: parseInt(odc_id),
                hasil_opm: parseFloat(hasil_opm),
                photo_before_tutup: photoUrls['before_tutup'],
                photo_before_buka: photoUrls['before_buka'],
                photo_after_tutup: photoUrls['after_tutup'],
                photo_after_buka: photoUrls['after_buka'],
                photo_opm: photoUrls['photo_opm']
            }]);

        if (error) throw error;

        res.status(200).json({ success: true, message: 'Laporan PM ODC Berhasil Disimpan!' });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, message: err.message });
    }
});

module.exports = app;
