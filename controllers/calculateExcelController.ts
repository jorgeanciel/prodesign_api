import { Request, Response } from "express";
import xlsx from "xlsx";
import path from "path";

// Interfaz para los datos que llegan desde el frontend
interface ProjectExcelData {
	// Aulas por nivel
	aulas_inicial_ciclo1?: number;
	aulas_inicial_ciclo2?: number;
	aulas_primaria?: number;
	aulas_secundaria?: number;

	// Ambientes compartidos
	aula_psicomotricidad?: number;
	sum_inicial?: number;
	biblioteca?: number;
	innovacion_primaria?: number;
	innovacion_secundaria?: number;
	taller_creativo_primaria?: number;
	taller_creativo_secundaria?: number;
	taller_ept?: number;
	laboratorio?: number;
	sum_prim_sec?: number;

	// Ambientes administrativos
	direccion_admin?: number;
	sala_reuniones?: number;
	sala_profesores?: number;
	sshh_admin?: number;
	cocina?: number;
	sshh_cocina?: number;
	depositos?: number;
	canchas_deportivas?: number;
	quiosco?: number;
	topico?: number;
	lactario?: number;
}

/**
 * Actualiza el archivo Excel IDEAS PRODESIGN con los datos del proyecto
 */
export const updateProjectExcel = (req: Request, res: Response) => {
	try {
		const projectData: ProjectExcelData = req.body;

		// 📌 Cargar el archivo Excel del backend
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);

		const sheetName = "CONSOLIDADO";
		const sheet = workbook.Sheets[sheetName];

		if (!sheet) {
			return res.status(400).json({
				error: `No se encontró la hoja '${sheetName}' en el archivo Excel`,
			});
		}

		// 📌 Mapeo de datos del proyecto a las celdas del Excel
		// Las celdas D64-D86 contienen las cantidades de aulas/ambientes
		const cellMapping: Record<string, keyof ProjectExcelData> = {
			D64: "aulas_inicial_ciclo1", // AULAS CICLO I
			D65: "aulas_inicial_ciclo2", // AULAS CICLO II
			D66: "aula_psicomotricidad", // AULA PSICOMOTRICIDAD
			D67: "aulas_primaria", // AULAS PRIMARIA
			D68: "aulas_secundaria", // AULAS SECUNDARIA
			D69: "sum_inicial", // SUM INICIAL
			D70: "biblioteca", // BIBLIOTECA
			D71: "innovacion_secundaria", // INNOVACION (suma de primaria + secundaria)
			D72: "taller_creativo_secundaria", // TALLER CREATIVO (suma)
			D73: "taller_ept", // TALLER EPT
			D74: "laboratorio", // LABORATORIO
			D75: "sum_prim_sec", // SUM PRIM + SEC
			D76: "direccion_admin", // DIRECCIÓN ADM.
			D77: "sala_reuniones", // SALA DE REUNIONES
			D78: "sala_profesores", // SALA DE PROFESORES
			D79: "sshh_admin", // SSHH ADM.
			D80: "cocina", // COCINA
			D81: "sshh_cocina", // SSHH COCINA
			D82: "depositos", // DEPOSITOS
			D83: "canchas_deportivas", // CANCHAS DEPORTIVAS
			D84: "quiosco", // QUIOSCO
			D85: "topico", // TOPICO
			D86: "lactario", // LACTARIO
		};

		// 📌 Actualizar las celdas con los datos del proyecto
		Object.entries(cellMapping).forEach(([cellRef, dataKey]) => {
			const value = projectData[dataKey];

			// Solo actualizar si el valor existe y es un número válido
			if (value !== undefined && value !== null && !isNaN(Number(value))) {
				sheet[cellRef] = {
					t: "n", // tipo número
					v: Number(value),
				};
			}
		});

		// 📌 También actualizar las celdas de origen si es necesario
		// Estas son las celdas que las D64-D86 referencian con fórmulas
		const originCellMapping: Record<string, keyof ProjectExcelData> = {
			D4: "aulas_inicial_ciclo1", // [192]INICIAL!G3
			D5: "aulas_inicial_ciclo2", // [192]INICIAL!G4
			D8: "aula_psicomotricidad", // [192]INICIAL!I34
			D26: "aulas_primaria", // [192]PRIMARIA!G3
			D43: "aulas_secundaria", // [192]SECUNDARIA!G3
			D9: "sum_inicial", // SUM
			D16: "topico", // TOPICO
			D17: "lactario", // LACTARIO
			D56: "cocina", // COCINA
			D46: "biblioteca", // BIBLIOTECA
		};

		Object.entries(originCellMapping).forEach(([cellRef, dataKey]) => {
			const value = projectData[dataKey];

			if (value !== undefined && value !== null && !isNaN(Number(value))) {
				sheet[cellRef] = {
					t: "n",
					v: Number(value),
				};
			}
		});

		// 📌 Guardar los cambios en el archivo
		xlsx.writeFile(workbook, excelPath);

		// 📌 Leer de nuevo el archivo para obtener los valores calculados
		const updatedWorkbook = xlsx.readFile(excelPath);
		const updatedSheet = updatedWorkbook.Sheets[sheetName];

		// 📌 Función auxiliar para leer valores
		const getCellValue = (cellRef: string): number => {
			return updatedSheet[cellRef]?.v ?? 0;
		};

		// 📌 Extraer los resultados calculados (metros cuadrados por ambiente)
		const results: Record<string, any> = {};

		// Recorrer las filas D64-D86 y obtener los m² calculados
		for (let i = 64; i <= 86; i++) {
			const labelCell = updatedSheet[`B${i}`];
			const cantidadCell = updatedSheet[`D${i}`];
			const m2Cell = updatedSheet[`E${i}`]; // Columna E suele tener los m² totales

			if (labelCell?.v) {
				const key = labelCell.v
					.toString()
					.toLowerCase()
					.replace(/\s+/g, "_")
					.replace(/[^a-z0-9_]/g, "");

				results[key] = {
					cantidad: cantidadCell?.v ?? 0,
					m2_total: m2Cell?.v ?? 0,
					label: labelCell.v,
				};
			}
		}

		// 📌 Respuesta con los datos actualizados
		return res.json({
			success: true,
			message: "Excel actualizado correctamente",
			data: {
				updated_cells: Object.keys(cellMapping).length,
				calculated_results: results,
			},
		});
	} catch (error) {
		console.error("Error actualizando Excel:", error);
		return res.status(500).json({
			error: "Error al actualizar el archivo Excel",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};

/**
 * Obtiene los valores actuales del Excel sin modificarlo
 */
export const getProjectExcelData = (req: Request, res: Response) => {
	try {
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);
		const sheet = workbook.Sheets["CONSOLIDADO"];

		const getCellValue = (cellRef: string) => sheet[cellRef]?.v ?? 0;

		// Extraer los valores actuales
		const currentData: ProjectExcelData = {
			aulas_inicial_ciclo1: getCellValue("D4"),
			aulas_inicial_ciclo2: getCellValue("D5"),
			aulas_primaria: getCellValue("D26"),
			aulas_secundaria: getCellValue("D43"),
			aula_psicomotricidad: getCellValue("D8"),
			sum_inicial: getCellValue("D9"),
			topico: getCellValue("D16"),
			lactario: getCellValue("D17"),
			cocina: getCellValue("D56"),
			biblioteca: getCellValue("D46"),
		};

		return res.json({
			success: true,
			data: currentData,
		});
	} catch (error) {
		console.error("Error leyendo Excel:", error);
		return res.status(500).json({
			error: "Error al leer el archivo Excel",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};
