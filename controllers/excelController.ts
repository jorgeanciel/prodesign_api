import { Request, Response } from "express";
import xlsx from "xlsx";
import fs from "fs";
import path from "path";

export const readMatrizExcel = (req: Request, res: Response) => {
	try {
		if (!req.file) {
			return res
				.status(400)
				.json({ error: "No se ha subido ningún archivo" });
		}

		// 📌 Leer "Plantilla Excel.xlsx"
		const plantillaWorkbook = xlsx.read(req.file.buffer, { type: "buffer" });
		const plantillaSheet =
			plantillaWorkbook.Sheets[plantillaWorkbook.SheetNames[0]];
		const plantillaData = xlsx.utils.sheet_to_json(plantillaSheet, {
			header: 1,
		});

		// 📌 Extraer aforos
		let aforoPorNivel = {
			inicial: plantillaData[2][1] || 0,
			primaria: plantillaData[3][1] || 0,
			secundaria: plantillaData[4][1] || 0,
		};

		// 📌 Cargar "MATRIZ_PLATAFORMA_MASTER.xlsx"
		const matrizPath = path.join(
			__dirname,
			"../../uploads/MATRIZ_PLATAFORMA_MASTER.xlsx"
		);
		const matrizWorkbook = xlsx.readFile(matrizPath);
		const matrizSheetName = "CALCULO AFORO";
		const matrizSheet = matrizWorkbook.Sheets[matrizSheetName];

		// 📌 Insertar aforos en la matriz
		matrizSheet["C10"] = { t: "n", v: aforoPorNivel.inicial };
		matrizSheet["C19"] = { t: "n", v: aforoPorNivel.primaria };
		matrizSheet["C28"] = { t: "n", v: aforoPorNivel.secundaria };

		// 📌 Guardar cambios
		xlsx.writeFile(matrizWorkbook, matrizPath);

		// 📌 Leer de nuevo el archivo
		const updatedMatrizWorkbook = xlsx.readFile(matrizPath);
		const updatedMatrizSheet = updatedMatrizWorkbook.Sheets[matrizSheetName];

		// 📌 Extraer resultados de Excel
		let result_data = {
			aforo_maximo: updatedMatrizSheet["B47"]?.v || 0,
			area_parcial: updatedMatrizSheet["L45"]?.v || 0,
			area_total: updatedMatrizSheet["B46"]?.v || 0,
			aulas: updatedMatrizSheet["B5"]?.v || 0,
			circulacion: updatedMatrizSheet["B45"]?.v || 0,
		};

		return res.json({ levels: aforoPorNivel, result_data });
	} catch (error) {
		return res.status(500).json({ error: "Error procesando el archivo" });
	}
};
