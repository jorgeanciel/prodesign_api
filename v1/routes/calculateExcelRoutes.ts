import { Router } from "express";
import {
	getExcelConfiguration,
	getProjectExcelData,
	updateProjectExcelAndCalculate,
} from "../../controllers/calculateExcelController";

const router = Router();

// Obtener configuración del Excel (áreas unitarias, nombres de ambientes)
router.get("/configuracion", getExcelConfiguration);

// Obtener valores actuales del Excel
router.get("/datos-actuales", getProjectExcelData);

// Actualizar Excel y calcular costos
router.post("/costos/calcular", updateProjectExcelAndCalculate);

export default router;
