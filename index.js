const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    for (let i = 0; i < 60; i++) {
        const files = fs.readdirSync(downloadPath);
        downloadedFile = files.find(f => (f.endsWith('.xls') || f.endsWith('.xlsx')) && !f.endsWith('.crdownload') && !f.startsWith('Report_'));
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    const oldPath = path.join(downloadPath, downloadedFile);
    const newPath = path.join(downloadPath, newFileName);
    
    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    
    fs.renameSync(oldPath, newPath);
    console.log(`   ✅ Saved as: ${newFileName}`);
    return newPath;
}

function getTodayFormatted() {
    const date = new Date();
    const options = { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'Asia/Bangkok' };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

// --- Main Script ---

(async () => {
    const { DTC_USERNAME, DTC_PASSWORD, EMAIL_USER, EMAIL_PASS, EMAIL_TO } = process.env;
    if (!DTC_USERNAME || !DTC_PASSWORD) {
        console.error('❌ Error: Missing DTC_USERNAME or DTC_PASSWORD secrets.');
        process.exit(1);
    }

    const downloadPath = path.resolve('./downloads');
    if (fs.existsSync(downloadPath)) fs.rmSync(downloadPath, { recursive: true, force: true });
    fs.mkdirSync(downloadPath);

    console.log('🚀 Starting DTC Automation (Fixed Login Logic)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    // Timeout 5 นาที ตามไฟล์ตัวอย่าง
    page.setDefaultNavigationTimeout(300000);
    page.setDefaultTimeout(300000);
    
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    
    await page.setViewport({ width: 1920, height: 1080 });
    await page.emulateTimezone('Asia/Bangkok');

    try {
        // =================================================================
        // STEP 1: LOGIN (แก้ไขตามไฟล์แนบ)
        // =================================================================
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        
        console.log('   Clicking Login...');
        // ใช้ logic เดียวกับไฟล์ตัวอย่าง: คลิกผ่าน DOM และรอจนกว่ากล่อง user จะหายไป
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        // =================================================================
        // STEP 2: REPORT 1 - Over Speed (Updated from attached file)
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        
        // ใช้ URL จากไฟล์แนบ indexJS.txt (Report_03.php)
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        
        // รอ Selector สำคัญ
        await page.waitForSelector('#speed_max', { visible: true, timeout: 60000 });
        await page.waitForSelector('#ddl_truck', { visible: true });

        // คำนวณเวลา 06:00 - 18:00 ของวันนี้
        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`   Setting parameters (Speed: 55, Time: ${startDateTime} - ${endDateTime})...`);

        await page.evaluate((start, end) => {
            // Speed (Command 8 from attached file)
            document.getElementById('speed_max').value = '55';
            
            // Date (Fixed to 06:00 - 18:00)
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            
            // Trigger Events
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));

            // Minute (Command 13)
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            
            // Select Truck "ทั้งหมด" (Command 14 from attached file)
            var selectElement = document.getElementById('ddl_truck'); 
            var options = selectElement.options; 
            for (var i = 0; i < options.length; i++) { 
                if (options[i].text.includes('ทั้งหมด')) { 
                    selectElement.value = options[i].value; 
                    break; 
                } 
            } 
            var event = new Event('change', { bubbles: true }); 
            selectElement.dispatchEvent(event);
        }, startDateTime, endDateTime);

        // Search Logic from attached file
        console.log('   Searching Report 1...');
        await page.evaluate(() => {
            if(typeof sertch_data === 'function') sertch_data();
            else {
                const btn = document.querySelector("span[onclick='sertch_data();']");
                if(btn) btn.click();
            }
        });

        // Wait for Export button (Logic from attached file)
        console.log('   Waiting for data to load...');
        try {
            await page.waitForSelector('#btnexport', { visible: true, timeout: 300000 }); // 5 mins max
        } catch(e) {
            console.warn('   ⚠️ Warning: Export button wait timed out, attempting to click anyway...');
        }

        console.log('   Exporting Report 1...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');

        // =================================================================
        // STEP 3-6: Other Reports (Placeholder for Puppeteer Replay)
        // =================================================================
        // ... พื้นที่สำหรับวาง Code Report 2-5 ...


        // =================================================================
        // STEP 7: Generate PDF (Placeholder)
        // =================================================================
        console.log('📑 Generating PDF Summary (Pending)...');


        // =================================================================
        // STEP 8: Send Email
        // =================================================================
        console.log('📧 Step 8: Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        const attachments = allFiles.map(file => ({
            filename: file,
            path: path.join(downloadPath, file)
        }));

        if (attachments.length > 0) {
            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงาน DTC Report ประจำวันที่ ${todayStr} (06:00 - 18:00)`,
                text: 'เรียน ผู้เกี่ยวข้อง,\n\nระบบได้ทำการดึงรายงานและแนบไฟล์มาพร้อมกับอีเมลฉบับนี้\n\nขอบคุณครับ\nDTC Automation Bot',
                attachments: attachments
            });
            console.log('   ✅ Email Sent Successfully!');
        } else {
            console.warn('⚠️ No files to send!');
        }

        // =================================================================
        // STEP 9: Cleanup Files
        // =================================================================
        console.log('🧹 Step 9: Cleaning up files...');
        const filesToDelete = fs.readdirSync(downloadPath);
        for (const file of filesToDelete) {
            try {
                fs.unlinkSync(path.join(downloadPath, file));
            } catch (err) { }
        }
        console.log('   ✅ Cleanup Complete.');

    } catch (err) {
        console.error('❌ Fatal Error:', err);
        await page.screenshot({ path: path.join(downloadPath, 'error_screenshot.png') });
        process.exit(1);
    } finally {
        await browser.close();
        console.log('🏁 Browser Closed.');
    }
})();
