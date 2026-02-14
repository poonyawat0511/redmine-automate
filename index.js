const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const axios = require('axios');

// CONFIG
const SPREADSHEET_ID = '1jP0yyjM89bizhwfgu_1WBs51ySloqrVnh5qiy9GjoqQ';
const SHEET_NAME = 'กพ ปุณยวัจน์';
const REDMINE_URL = 'https://redmine.ochi.link';
const TARGET_ISSUE_ID = 6741;
const ACTIVITY_ID = 20;

async function run() {
    const auth = new JWT({
        email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(SPREADSHEET_ID, auth);
    await doc.loadInfo();
    const sheet = doc.sheetsByTitle[SHEET_NAME];
    const rows = await sheet.getRows();

    console.log(`✅ เริ่มประมวลผล ${rows.length} แถว...`);

    for (const row of rows) {
        // อ้างอิงชื่อคอลัมน์ตามหัว Excel (วันที่, manday, โครงการ, รายละเอียด, Status)
        if (row.get('Status') !== 'Synced' && row.get('วันที่')) {
            const manday = parseFloat(row.get('manday'));
            if (manday > 0) {
                const payload = {
                    time_entry: {
                        issue_id: TARGET_ISSUE_ID,
                        spent_on: row.get('วันที่'),
                        hours: manday * 8,
                        activity_id: ACTIVITY_ID,
                        comments: `[${row.get('โครงการ')}] ${row.get('รายละเอียด')}`.substring(0, 255)
                    }
                };

                try {
                    console.log(`⏳ กำลังบันทึกวันที่ ${row.get('วันที่')}...`);
                    const res = await axios.post(`${REDMINE_URL}/time_entries.json`, payload, {
                        headers: { 'X-Redmine-API-Key': process.env.REDMINE_API_KEY.trim() }
                    });
                    
                    if (res.status === 201) {
                        row.set('Status', 'Synced');
                        await row.save();
                        console.log(`✅ สำเร็จ!`);
                    }
                } catch (err) {
                    console.error(`❌ พลาดที่วันที่ ${row.get('วันที่')}: ${err.message}`);
                }
            }
        }
    }
}

run();