import { Router } from "express";
import {
	getProjectExcelData,
	updateProjectExcel,
} from "../../controllers/calculateExcelController";

const router = Router();

router.post("/update-project-excel", updateProjectExcel);

// Obtener datos actuales del Excel
router.get("/get-project-excel", getProjectExcelData);

export default router;
