import { Request, Response } from "express";
import xlsx from "xlsx";
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
		const aforoInicial = plantillaSheet["B4"]?.v || 0;
		const aforoPrimaria = plantillaSheet["B12"]?.v || 0;
		const aforoSecundaria = plantillaSheet["B20"]?.v || 0;

		console.log("Aforos extraídos de la plantilla:");
		console.log("- Inicial (B4):", aforoInicial);
		console.log("- Primaria (B12):", aforoPrimaria);
		console.log("- Secundaria (B20):", aforoSecundaria);

		// 📌 Cargar "MATRIZ_PLATAFORMA_MASTER.xlsx"
		const matrizPath = path.resolve("uploads", "CALCULADORA_V2.xlsx");
		const matrizWorkbook = xlsx.readFile(matrizPath);

		const sheetInicial = matrizWorkbook.Sheets["INICIAL"];
		const sheetPrimaria = matrizWorkbook.Sheets["PRIMARIA"];
		const sheetSecundaria = matrizWorkbook.Sheets["SECUNDARIA"];

		// Insertar aforos en celda A3 de cada hoja
		sheetInicial["A3"] = { t: "n", v: aforoInicial };
		sheetPrimaria["A3"] = { t: "n", v: aforoPrimaria };
		sheetSecundaria["A3"] = { t: "n", v: aforoSecundaria };
		console.log("Aforos insertados en CALCULADORA_V2:");
		console.log("- INICIAL.A3 =", aforoInicial);
		console.log("- PRIMARIA.A3 =", aforoPrimaria);
		console.log("- SECUNDARIA.A3 =", aforoSecundaria);

		// 📌 PASO 5: Guardar cambios en el archivo
		xlsx.writeFile(matrizWorkbook, matrizPath);

		// 📌 PASO 6: Leer de nuevo para obtener los valores calculados
		const updatedCalculadoraWorkbook = xlsx.readFile(matrizPath);
		const updatedSheetInicial = updatedCalculadoraWorkbook.Sheets["INICIAL"];
		const updatedSheetPrimaria =
			updatedCalculadoraWorkbook.Sheets["PRIMARIA"];
		const updatedSheetSecundaria =
			updatedCalculadoraWorkbook.Sheets["SECUNDARIA"];

		// Función auxiliar para obtener valor de celda
		const getCellValue = (sheet: any, cellRef: string): number => {
			const cell = sheet[cellRef];
			return cell?.v ?? 0;
		};

		// 📌 PASO 7: Extraer cantidad de aulas de G3 de cada hoja
		const aulasInicial = getCellValue(updatedSheetInicial, "G3");
		const aulasPrimaria = getCellValue(updatedSheetPrimaria, "G3");
		const aulasSecundaria = getCellValue(updatedSheetSecundaria, "G3");
		const totalAulas = aulasInicial + aulasPrimaria + aulasSecundaria;

		console.log("Cantidad de aulas calculadas:");
		console.log("- INICIAL.G3 =", aulasInicial);
		console.log("- PRIMARIA.G3 =", aulasPrimaria);
		console.log("- SECUNDARIA.G3 =", aulasSecundaria);
		console.log("- Total =", totalAulas);

		// 📌 PASO 8: Extraer otros valores si existen en las hojas
		// (Puedes agregar más celdas según necesites)

		// Verificar si existe la hoja CONSOLIDADO para obtener totales
		let consolidadoData: any = {};
		if (updatedCalculadoraWorkbook.SheetNames.includes("CONSOLIDADO")) {
			const consolidadoSheet =
				updatedCalculadoraWorkbook.Sheets["CONSOLIDADO"];

			// Aquí puedes extraer los datos que necesites de CONSOLIDADO
			// Por ejemplo, si hay algún total general u otra información
			consolidadoData = {
				// Agregar las celdas que necesites del CONSOLIDADO
				// ejemplo: total_general: getCellValue(consolidadoSheet, "B10")
			};
		}

		// 📌 PASO 9: Construir la respuesta
		const response = {
			levels: {
				inicial: {
					aforo: aforoInicial,
					aulas: aulasInicial,
				},
				primaria: {
					aforo: aforoPrimaria,
					aulas: aulasPrimaria,
				},
				secundaria: {
					aforo: aforoSecundaria,
					aulas: aulasSecundaria,
				},
			},
			result_data: {
				aulas: totalAulas,
				aforo_maximo: aforoInicial + aforoPrimaria + aforoSecundaria,
				// Puedes agregar más campos si los necesitas del CONSOLIDADO u otras hojas
			},
			// Mantener estructura de respuesta para compatibilidad
			classroom_measurements: {
				columna: 0.25,
				muro_vertical: 5,
				muro_horizontal: 8,
			},
			construction_info: {
				pisos: 2,
				area_general: "-",
			},
			toilets_per_student: {
				inicial: {
					ninos: 25,
					ninas: 25,
				},
				primaria: {
					ninos: 60,
					ninas: 60,
				},
				secundaria: {
					ninos: 60,
					ninas: 60,
				},
			},
			stairs: {
				paso: 28,
				contrapaso: 17,
				ancho: 1.2,
				cantidad_de_contrapasos: 16,
				modulo: {
					largo: 4.2,
					ancho: 2.4,
				},
			},
		};

		return res.json(response);
	} catch (error) {
		console.error("Error procesando matriz:", error);
		return res.status(500).json({
			error: "Error al procesar el archivo",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};
