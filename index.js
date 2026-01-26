const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

// --- Helper Functions ---

async function waitForDownloadAndRename(downloadPath, newFileName) {
    console.log(`   Waiting for download: ${newFileName}...`);
    let downloadedFile = null;

    // รอไฟล์สูงสุด 120 วินาที
    for (let i = 0; i < 120; i++) {
        const files = fs.readdirSync(downloadPath);
        
        // LOGIC ใหม่: หาไฟล์ .xls/.xlsx ที่ไม่ใช่ไฟล์ชั่วคราว (.crdownload) 
        // และ **ต้องไม่มี** Prefix "DTC_Completed_" (เพื่อไม่ให้ซ้ำกับไฟล์ที่เราเปลี่ยนชื่อไปแล้ว)
        downloadedFile = files.find(f => 
            (f.endsWith('.xls') || f.endsWith('.xlsx')) && 
            !f.endsWith('.crdownload') && 
            !f.startsWith('DTC_Completed_') // จุดสำคัญ: กรองไฟล์ที่ทำเสร็จแล้วออกไป
        );
        
        if (downloadedFile) break;
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    if (!downloadedFile) {
        throw new Error(`Download failed or timed out for ${newFileName}`);
    }

    // รออีกนิดเพื่อให้เขียนไฟล์เสร็จสมบูรณ์
    await new Promise(resolve => setTimeout(resolve, 3000));

    const oldPath = path.join(downloadPath, downloadedFile);
    // เติม Prefix "DTC_Completed_" นำหน้าชื่อไฟล์ใหม่เสมอ เพื่อให้ Logic การค้นหาข้างบนทำงานถูกต้อง
    const finalFileName = `DTC_Completed_${newFileName}`;
    const newPath = path.join(downloadPath, finalFileName);
    
    // ตรวจสอบขนาดไฟล์
    const stats = fs.statSync(oldPath);
    console.log(`   Found File: ${downloadedFile} (Size: ${stats.size} bytes)`);
    
    if (stats.size === 0) {
        throw new Error(`Downloaded file ${downloadedFile} is empty!`);
    }

    // ลบไฟล์ปลายทางถ้ามีซ้ำ
    if (fs.existsSync(newPath)) fs.unlinkSync(newPath);
    
    fs.renameSync(oldPath, newPath);
    console.log(`   ✅ Renamed to: ${finalFileName}`);
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

    console.log('🚀 Starting DTC Automation (5 Reports Structure)...');
    
    const browser = await puppeteer.launch({
        headless: true,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--start-maximized']
    });

    const page = await browser.newPage();
    // Timeout รวม 20 นาที (เผื่อ 5 รายงาน)
    page.setDefaultNavigationTimeout(1200000);
    page.setDefaultTimeout(1200000);
    
    const client = await page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: downloadPath });
    
    await page.setViewport({ width: 1920, height: 1080 });
    await page.emulateTimezone('Asia/Bangkok');

    try {
        // =================================================================
        // STEP 1: LOGIN
        // =================================================================
        console.log('1️⃣ Step 1: Login...');
        await page.goto('https://gps.dtc.co.th/ultimate/index.php', { waitUntil: 'domcontentloaded' });
        await page.waitForSelector('#txtname', { visible: true, timeout: 60000 });
        await page.type('#txtname', DTC_USERNAME);
        await page.type('#txtpass', DTC_PASSWORD);
        
        await Promise.all([
            page.evaluate(() => document.getElementById('btnLogin').click()),
            page.waitForFunction(() => !document.querySelector('#txtname'), { timeout: 60000 })
        ]);
        console.log('✅ Login Success');

        const todayStr = getTodayFormatted();
        const startDateTime = `${todayStr} 06:00`;
        const endDateTime = `${todayStr} 18:00`;
        console.log(`🕒 Time Settings: ${startDateTime} to ${endDateTime}`);

        // =================================================================
        // STEP 2: REPORT 1 - Over Speed
        // =================================================================
        console.log('📊 Processing Report 1: Over Speed...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_03.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#speed_max', { visible: true });
        await page.waitForSelector('#ddl_truck', { visible: true });
        await new Promise(r => setTimeout(r, 2000));

        await page.evaluate((start, end) => {
            document.getElementById('speed_max').value = '55';
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
            if(document.getElementById('ddlMinute')) document.getElementById('ddlMinute').value = '1';
            
            var selectElement = document.getElementById('ddl_truck'); 
            var options = selectElement.options; 
            for (var i = 0; i < options.length; i++) { 
                if (options[i].text.includes('ทั้งหมด')) { selectElement.value = options[i].value; break; } 
            } 
            selectElement.dispatchEvent(new Event('change', { bubbles: true }));
        }, startDateTime, endDateTime);

        console.log('   Searching Report 1...');
        await page.evaluate(() => {
            if(typeof sertch_data === 'function') sertch_data();
            else document.querySelector("span[onclick='sertch_data();']").click();
        });

        console.log('   ⏳ Waiting 5 mins...');
        await new Promise(resolve => setTimeout(resolve, 300000));
        
        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        console.log('   Exporting Report 1...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        // บันทึกไฟล์ที่ 1
        await waitForDownloadAndRename(downloadPath, 'Report1_OverSpeed.xls');


        // =================================================================
        // STEP 3: REPORT 2 - Idling
        // =================================================================
        console.log('📊 Processing Report 2: Idling...');
        await page.goto('https://gps.dtc.co.th/ultimate/Report/Report_02.php', { waitUntil: 'domcontentloaded' });
        
        await page.waitForSelector('#date9', { visible: true });
        await page.waitForSelector('#date10', { visible: true });
        
        await page.evaluate((start, end) => {
            document.getElementById('date9').value = start;
            document.getElementById('date10').value = end;
            document.getElementById('date9').dispatchEvent(new Event('change'));
            document.getElementById('date10').dispatchEvent(new Event('change'));
        }, startDateTime, endDateTime);

        console.log('   Searching Report 2...');
        await page.click('td:nth-of-type(6) > span');

        console.log('   ⏳ Waiting 5 mins...');
        await new Promise(resolve => setTimeout(resolve, 300000));

        try { await page.waitForSelector('#btnexport', { visible: true, timeout: 60000 }); } catch(e) {}
        console.log('   Exporting Report 2...');
        await page.evaluate(() => document.getElementById('btnexport').click());
        
        // บันทึกไฟล์ที่ 2
        await waitForDownloadAndRename(downloadPath, 'Report2_Idling.xls');


        // =================================================================
        // STEP 4: REPORT 3 (รอ Code)
        // =================================================================
        console.log('📊 Processing Report 3...');
        // TODO: วาง Code Puppeteer สำหรับ Report 3 ตรงนี้
        // ...
        // ...
        // เมื่อวาง Code เสร็จ ให้ Uncomment บรรทัดล่างนี้เพื่อให้ระบบรอดาวน์โหลด
        // await waitForDownloadAndRename(downloadPath, 'Report3_Name.xls');


        // =================================================================
        // STEP 5: REPORT 4 (รอ Code)
        // =================================================================
        console.log('📊 Processing Report 4...');
        // TODO: วาง Code Puppeteer สำหรับ Report 4 ตรงนี้
        // ...
        // ...
        // await waitForDownloadAndRename(downloadPath, 'Report4_Name.xls');


        // =================================================================
        // STEP 6: REPORT 5 (รอ Code)
        // =================================================================
        console.log('📊 Processing Report 5...');
        // TODO: วาง Code Puppeteer สำหรับ Report 5 ตรงนี้
        // ...
        // ...
        // await waitForDownloadAndRename(downloadPath, 'Report5_Name.xls');


        // =================================================================
        // STEP 7: Generate PDF (Pending)
        // =================================================================
        console.log('📑 Generating PDF Summary (Pending)...');


        // =================================================================
        // STEP 8: Send Email
        // =================================================================
        console.log('📧 Step 8: Sending Email...');
        
        const allFiles = fs.readdirSync(downloadPath);
        
        // กรองเอาเฉพาะไฟล์ที่ผ่านการ Rename แล้ว (DTC_Completed_...) หรือ PDF
        const filesToSend = allFiles.filter(file => 
            file.startsWith('DTC_Completed_') || file.endsWith('.pdf')
        );
        
        const attachments = filesToSend.map(file => {
            const filePath = path.join(downloadPath, file);
            return { filename: file, path: filePath };
        });

        if (attachments.length > 0) {
            const transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: { user: EMAIL_USER, pass: EMAIL_PASS }
            });

            await transporter.sendMail({
                from: `"DTC Reporter" <${EMAIL_USER}>`,
                to: EMAIL_TO,
                subject: `รายงาน DTC Report (5 ฉบับ) ประจำวันที่ ${todayStr}`,
                text: `เรียน ผู้เกี่ยวข้อง,\n\nระบบส่งไฟล์รายงานจำนวน ${attachments.length} ฉบับ ดังแนบ\n(ข้อมูลช่วงเวลา 06:00 - 18:00)\n\nขอบคุณครับ\nDTC Automation Bot`,
                attachments: attachments
            });
            console.log(`   ✅ Email Sent Successfully! (${attachments.length} files)`);
        } else {
            console.warn('⚠️ No "DTC_Completed_" files found to send!');
        }

        // =================================================================
        // STEP 9: Cleanup
        // =================================================================
        console.log('🧹 Cleanup...');
        // (Optional) ลบไฟล์หลังส่ง ถ้าต้องการเก็บไว้ดู Debug ให้ Comment บรรทัดล่างทิ้ง
        // fs.rmSync(downloadPath, { recursive: true, force: true });
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
