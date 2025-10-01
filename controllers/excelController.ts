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
		const plantillaData = xlsx.utils.sheet_to_json(plantillaSheet, {
			header: 1,
		});

		// 📌 Extraer aforos desde la plantilla
		// Basado en la estructura real del archivo:
		// - Fila 3 (Excel 4): EDUCACION INICIAL
		// - Fila 11 (Excel 12): EDUCACION PRIMARIA
		// - Fila 19 (Excel 20): EDUCACION SECUNDARIA
		const aforoInicial = plantillaData[3]?.[1] || 0;
		const aforoPrimaria = plantillaData[11]?.[1] || 0;
		const aforoSecundaria = plantillaData[19]?.[1] || 0;

		// 📌 Cargar "MATRIZ_PLATAFORMA_MASTER.xlsx"
		const matrizPath = path.resolve(
			"uploads",
			"MATRIZ_PLATAFORMA_MASTER.xlsx"
		);
		const matrizWorkbook = xlsx.readFile(matrizPath);
		const matrizSheetName = "CALCULO AFORO";
		const matrizSheet = matrizWorkbook.Sheets[matrizSheetName];

		// 📌 Insertar aforos en las celdas individuales
		// INICIAL (3 tipos de aulas: C11-C13)
		matrizSheet["C11"] = { t: "n", v: aforoInicial };
		matrizSheet["C12"] = { t: "n", v: aforoInicial };
		matrizSheet["C13"] = { t: "n", v: aforoInicial };

		// PRIMARIA (6 tipos de aulas: C20-C25)
		for (let i = 20; i <= 25; i++) {
			matrizSheet[`C${i}`] = { t: "n", v: aforoPrimaria };
		}

		// SECUNDARIA (5 tipos de aulas: C29-C33)
		for (let i = 29; i <= 33; i++) {
			matrizSheet[`C${i}`] = { t: "n", v: aforoSecundaria };
		}

		// 📌 Función auxiliar para obtener valor numérico de una celda
		const getCellValue = (cellRef: string): number => {
			const cell = matrizSheet[cellRef];
			return cell?.v ?? 0;
		};

		// 📌 Calcular cantidad de aulas manualmente
		const calcularAulas = (startRow: number, endRow: number): number => {
			let totalAulas = 0;
			for (let i = startRow; i <= endRow; i++) {
				const B = getCellValue(`B${i}`); // Número de ambientes
				const C = getCellValue(`C${i}`); // Aforo por aula
				const E = getCellValue(`E${i}`); // m² por persona

				// Fórmula: G = ROUNDUP(D/F, 0) donde D=B*C y F=E*B
				const D = B * C;
				const F = E * B;

				if (F > 0) {
					const aulasCalculadas = Math.ceil(D / F);
					totalAulas += aulasCalculadas;
				}
			}
			return totalAulas;
		};

		// Calcular aulas por nivel
		const aulasInicial = calcularAulas(11, 13);
		const aulasPrimaria = calcularAulas(20, 25);
		const aulasSecundaria = calcularAulas(29, 33);
		const totalAulas = aulasInicial + aulasPrimaria + aulasSecundaria;

		// 📌 Guardar cambios en el archivo
		xlsx.writeFile(matrizWorkbook, matrizPath);

		// 📌 Leer de nuevo el archivo para obtener otros valores calculados
		const updatedMatrizWorkbook = xlsx.readFile(matrizPath);
		const updatedMatrizSheet = updatedMatrizWorkbook.Sheets[matrizSheetName];

		// 📌 Función para obtener valores del archivo actualizado
		const getUpdatedValue = (cellRef: string): number => {
			const cell = updatedMatrizSheet[cellRef];
			return cell?.v ?? 0;
		};

		// 📌 AJUSTA ESTAS REFERENCIAS DE CELDAS SEGÚN TU ARCHIVO MAESTRO
		// Estas son estimaciones basadas en estructuras típicas de Excel
		// Debes verificar y ajustar cada referencia

		// Medidas de aula (classroom_measurements)
		// Buscar en las celdas donde estén definidas estas medidas
		const columna = getUpdatedValue("E5") || 0.25; // Ajusta la celda
		const muroVertical = getUpdatedValue("E6") || 5; // Ajusta la celda
		const muroHorizontal = getUpdatedValue("E7") || 8; // Ajusta la celda

		// Información de construcción (construction_info)
		const pisos = getUpdatedValue("E3") || 2; // Ajusta la celda
		const areaGeneral = updatedMatrizSheet["E4"]?.v || "-"; // Ajusta la celda

		// Baños por estudiante (toilets_per_student)
		// Estas celdas también deben ser ajustadas según tu archivo
		const toiletsInicialNinos = getUpdatedValue("K11") || 25;
		const toiletsInicialNinas = getUpdatedValue("L11") || 25;
		const toiletsPrimariaNinos = getUpdatedValue("K20") || 60;
		const toiletsPrimariaNinas = getUpdatedValue("L20") || 60;
		const toiletsSecundariaNinos = getUpdatedValue("K29") || 60;
		const toiletsSecundariaNinas = getUpdatedValue("L29") || 60;

		// Escaleras (stairs)
		// Ajusta estas celdas según donde estén en tu archivo
		const paso = getUpdatedValue("O5") || 28;
		const contrapaso = getUpdatedValue("O6") || 17;
		const anchoEscalera = getUpdatedValue("O7") || 1.2;
		const cantidadContrapasos = getUpdatedValue("O8") || 16;
		const moduloLargo = getUpdatedValue("O9") || 4.2;
		const moduloAncho = getUpdatedValue("O10") || 2.4;

		// 📌 Construir respuesta en el formato solicitado
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
				area_parcial: getUpdatedValue("L45"),
				circulacion: getUpdatedValue("B45"),
				area_total: getUpdatedValue("B46"),
				aforo_maximo: getUpdatedValue("B47"),
				aulas: totalAulas,
			},
			classroom_measurements: {
				columna: columna,
				muro_vertical: muroVertical,
				muro_horizontal: muroHorizontal,
			},
			construction_info: {
				pisos: pisos,
				area_general: areaGeneral,
			},
			toilets_per_student: {
				inicial: {
					ninos: toiletsInicialNinos,
					ninas: toiletsInicialNinas,
				},
				primaria: {
					ninos: toiletsPrimariaNinos,
					ninas: toiletsPrimariaNinas,
				},
				secundaria: {
					ninos: toiletsSecundariaNinos,
					ninas: toiletsSecundariaNinas,
				},
			},
			stairs: {
				paso: paso,
				contrapaso: contrapaso,
				ancho: anchoEscalera,
				cantidad_de_contrapasos: cantidadContrapasos,
				modulo: {
					largo: moduloLargo,
					ancho: moduloAncho,
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
