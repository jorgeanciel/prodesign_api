// controllers/costosController.ts
import { Request, Response } from "express";
import xlsx from "xlsx";
import path from "path";

/**
 * Obtiene el costo de infraestructura del Excel IDEAS_PRODESIGN
 * Lee la celda J42 de la hoja "COSTO INFRA"
 */
export const getCostosInfraestructura = (req: Request, res: Response) => {
	try {
		// 📌 Cargar el Excel desde el backend
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");

		const workbook = xlsx.readFile(excelPath);

		// Verificar que la hoja existe
		if (!workbook.Sheets["COSTO INFRA"]) {
			return res.status(404).json({
				error: "No se encontró la hoja 'COSTO INFRA' en el archivo Excel",
				hojas_disponibles: workbook.SheetNames,
			});
		}

		const costoInfraSheet = workbook.Sheets["COSTO INFRA"];

		// 📌 Leer la celda J42 (PRESUPUESTO TOTAL)
		const getCellValue = (cellRef: string): number => {
			const cell = costoInfraSheet[cellRef];
			return cell?.v ?? 0;
		};

		// Extraer todos los valores de la columna J relacionados al presupuesto
		const costosInfraestructura = {
			costo_directo: getCellValue("J37"),
			gastos_generales: getCellValue("J38"),
			utilidad: getCellValue("J39"),
			subtotal: getCellValue("J40"),
			igv: getCellValue("J41"),
			presupuesto_total: getCellValue("J42"), // ⭐ Este es el valor principal
		};

		// 📌 Responder con los datos
		return res.json({
			success: true,
			data: {
				infraestructura: costosInfraestructura,
			},
		});
	} catch (error) {
		console.error("Error leyendo costos del Excel:", error);
		return res.status(500).json({
			error: "Error al leer el archivo Excel",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};

/**
 * Obtiene el costo de equipamiento del Excel IDEAS_PRODESIGN
 * Lee el total de la hoja "COSTO EQUIPAMIENTO"
 */
export const getCostosEquipamiento = (req: Request, res: Response) => {
	try {
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);

		if (!workbook.Sheets["COSTO EQUIPAMIENTO"]) {
			return res.status(404).json({
				error: "No se encontró la hoja 'COSTO EQUIPAMIENTO' en el archivo Excel",
			});
		}

		const costoEquipSheet = workbook.Sheets["COSTO EQUIPAMIENTO"];

		// Ajusta la celda según donde esté el total en tu archivo
		const getCellValue = (cellRef: string): number => {
			const cell = costoEquipSheet[cellRef];
			return cell?.v ?? 0;
		};

		const costosEquipamiento = {
			presupuesto_total: getCellValue("J42"), // Ajusta la celda si es necesario
		};

		return res.json({
			success: true,
			data: {
				equipamiento: costosEquipamiento,
			},
		});
	} catch (error) {
		console.error("Error leyendo costos de equipamiento:", error);
		return res.status(500).json({
			error: "Error al leer el archivo Excel",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};

/**
 * Obtiene TODOS los costos del proyecto (infraestructura + equipamiento)
 */
export const getCostosCompletos = (req: Request, res: Response) => {
	try {
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);

		// Helper para leer valores de cualquier hoja
		const getCellValue = (sheetName: string, cellRef: string): number => {
			const sheet = workbook.Sheets[sheetName];
			if (!sheet) return 0;
			const cell = sheet[cellRef];
			return cell?.v ?? 0;
		};

		// Leer costos de infraestructura
		const infraestructura = {
			costo_directo: getCellValue("COSTO INFRA", "J37"),
			gastos_generales: getCellValue("COSTO INFRA", "J38"),
			utilidad: getCellValue("COSTO INFRA", "J39"),
			subtotal: getCellValue("COSTO INFRA", "J40"),
			igv: getCellValue("COSTO INFRA", "J41"),
			presupuesto_total: getCellValue("COSTO INFRA", "J42"),
		};

		// Leer costos de equipamiento (ajusta la celda si es necesario)
		const equipamiento = {
			presupuesto_total: getCellValue("COSTO EQUIPAMIENTO", "J42"),
		};

		// Calcular total general
		const total_general =
			infraestructura.presupuesto_total + equipamiento.presupuesto_total;

		return res.json({
			success: true,
			data: {
				infraestructura,
				equipamiento,
				total_general,
			},
		});
	} catch (error) {
		console.error("Error leyendo costos completos:", error);
		return res.status(500).json({
			error: "Error al leer el archivo Excel",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};
