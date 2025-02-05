const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");

const app = express();
const PORT = 3000;
const FILE_PATH = path.join("C:\\Campus\\Uniair\\trainingLaravel", "users.xlsx");

const sleep = (ms) => new Promise((res) => setTimeout(res, ms));

app.use(cors());
app.use(bodyParser.json());

// Fungsi membaca data dari Excel
const readExcel = () => {
    if (!fs.existsSync(FILE_PATH)) {
        console.log("File Excel tidak ditemukan!");
        return [];
    }
    const workbook = XLSX.readFile(FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const users = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    console.log("Data dari Excel:", users);
    return users;
};

// Async function untuk menjalankan robot
async function runRobot() {
    const users = readExcel();
    if (users.length === 0) {
        console.log("Tidak ada data user di Excel.");
        return;
    }

    console.log("Memulai robot...");
    const browser = await puppeteer.launch({ headless: false, args: ["--no-sandbox", "--disable-setuid-sandbox"] });
    const page = await browser.newPage();

    for (let user of users) {
        console.log(`Memproses user: ${user.Name} - ${user.Email}`);

        await page.goto("http://localhost:8000/add_user", { waitUntil: "networkidle2" });

        // Pastikan elemen sudah ada sebelum mengetik
        await page.waitForSelector("#name");
        await page.waitForSelector("#email");
        await page.waitForSelector("button[type=submit]");

        await page.type("#name", user.Name);
        await page.type("#email", user.Email);
        await page.click("button[type=submit]");

        console.log(`User ${user.Name} dikirim`);

        // Tunggu sebelum lanjut ke user berikutnya
        await sleep(3000);
    }

    await browser.close();
    console.log("Semua user berhasil dikirim!");
}

// Menjalankan server
app.listen(PORT, async () => {
    console.log(`Server berjalan di http://localhost:${PORT}`);
    
    // Langsung jalankan robot saat server dinyalakan
    await runRobot();
});
