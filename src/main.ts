import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import path from "path";

const app = express();
const port = 3000;

app.use(cors());
app.use(express.static("public"));
app.use(express.json());

const upload = multer({ dest: "temp/" });

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
