import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import * as xlsx from "xlsx";
import path from "path";
import { parseNakladna } from "../logic/parseNakladna";
import { fillTemplatesFromParsed } from "../logic/fillTemplates";
import { NakladnaWithMeta } from "../interfaces/NakladnaWithMeta";
import { getDataRangeFromSheet, OperationType, stringifyTableAsTSV } from "../units";

const app = express();
const port = 3001;

const upload = multer({ dest: "temp/" });

app.use(cors());
app.use(express.json());
app.use(express.static("public"));

app.post("/api/process-nakladni", upload.array("nakladni"), async (req, res) => {
    try {
        const files = req.files as Express.Multer.File[];
        const meta = JSON.parse(req.body.meta) as { operation: OperationType; cell: string }[];

        if (!files || !meta || files.length !== meta.length) {
            res.status(400).json({ error: "Файли та мета не збігаються" });
            return;
        }

        const nakladniWithMeta: NakladnaWithMeta[] = [];

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const buffer = fs.readFileSync(file.path);
            const workbook = xlsx.read(buffer);
            const parsed = parseNakladna(workbook);

            nakladniWithMeta.push(new NakladnaWithMeta(
                parsed,
                meta[i].operation,
                meta[i].cell
            ));

            fs.unlinkSync(file.path);
        }

        const templates = fillTemplatesFromParsed(nakladniWithMeta);

        const response: Record<string, { НАДІЙШЛО: string; ВИБУЛО: string }> = {};

        for (const [key, wb] of Object.entries(templates)) {
            response[key] = {
                "НАДІЙШЛО": stringifyTableAsTSV(getDataRangeFromSheet(wb, "НАДІЙШЛО")),
                "ВИБУЛО": stringifyTableAsTSV(getDataRangeFromSheet(wb, "ВИБУЛО")),
            };
        }

        res.json(response);
    } catch (e) {
        console.error("❌ Помилка при обробці:", e);
        res.status(500).json({ error: "Помилка сервера" });
    }
});


app.listen(port, () => {
    console.log(`✅ Сервер запущено на http://localhost:${port}`);
});
