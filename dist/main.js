"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const multer_1 = __importDefault(require("multer"));
const cors_1 = __importDefault(require("cors"));
const app = (0, express_1.default)();
const port = 3000;
app.use((0, cors_1.default)());
app.use(express_1.default.static("public"));
app.use(express_1.default.json());
const upload = (0, multer_1.default)({ dest: "temp/" });
app.post("/api/convert-excel", upload.single("excel"), (req, res) => {
    (async () => {
    })();
});
app.post("/api/update-excel", upload.single("excel"), (req, res) => {
    (async () => {
    })();
});
app.post("/api/ordered-quantities", (req, res) => {
});
app.listen(port, () => {
    console.log(`✅ Сервер запущено на http://localhost:${port}`);
});
