import { Router } from "express";
import multer from "multer";
import { readMatrizExcel } from "../../controllers/excelController";

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });

router.post("/readMatriz", upload.single("file"), readMatrizExcel);

export default router;
