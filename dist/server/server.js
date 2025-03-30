"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const multer_1 = __importDefault(require("multer"));
const cors_1 = __importDefault(require("cors"));
const fs_1 = __importDefault(require("fs"));
const xlsx = __importStar(require("xlsx"));
const parseNakladna_1 = require("../logic/parseNakladna");
const fillTemplates_1 = require("../logic/fillTemplates");
const NakladnaWithMeta_1 = require("../interfaces/NakladnaWithMeta");
const units_1 = require("../units");
const app = (0, express_1.default)();
const port = 3001;
const upload = (0, multer_1.default)({ dest: "temp/" });
app.use((0, cors_1.default)());
app.use(express_1.default.json());
app.use(express_1.default.static("public"));
app.post("/api/process-nakladni", upload.array("nakladni"), async (req, res) => {
    try {
        const files = req.files;
        const meta = JSON.parse(req.body.meta);
        if (!files || !meta || files.length !== meta.length) {
            res.status(400).json({ error: "Файли та мета не збігаються" });
            return;
        }
        const nakladniWithMeta = [];
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const buffer = fs_1.default.readFileSync(file.path);
            const workbook = xlsx.read(buffer);
            const parsed = (0, parseNakladna_1.parseNakladna)(workbook);
            nakladniWithMeta.push(new NakladnaWithMeta_1.NakladnaWithMeta(parsed, meta[i].operation, meta[i].cell));
            fs_1.default.unlinkSync(file.path);
        }
        const templates = (0, fillTemplates_1.fillTemplatesFromParsed)(nakladniWithMeta);
        const response = {};
        for (const [key, wb] of Object.entries(templates)) {
            response[key] = {
                "НАДІЙШЛО": (0, units_1.stringifyTableAsTSV)((0, units_1.getDataRangeFromSheet)(wb, "НАДІЙШЛО")),
                "ВИБУЛО": (0, units_1.stringifyTableAsTSV)((0, units_1.getDataRangeFromSheet)(wb, "ВИБУЛО")),
            };
        }
        res.json(response);
    }
    catch (e) {
        console.error("❌ Помилка при обробці:", e);
        res.status(500).json({ error: "Помилка сервера" });
    }
});
app.listen(port, () => {
    console.log(`✅ Сервер запущено на http://localhost:${port}`);
});
