import { Request, Response } from "express";
import xlsx from "xlsx";

export const readMatrizExcel = (req: Request, res: Response) => {
	try {
		if (!req.file) {
			return res
				.status(400)
				.json({ error: "No se ha subido ningún archivo" });
		}

		// 📌 Leer el archivo Excel subido (NUEVO_FORMATO.xlsx)
		const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
		const sheet = workbook.Sheets[workbook.SheetNames[0]]; // "PROGRAMA ARQUITECTÓNICO"

		// Función auxiliar para obtener valor de celda
		const getCellValue = (cellRef: string): number => {
			const cell = sheet[cellRef];
			return cell?.v ?? 0;
		};

		// 📌 EDUCACIÓN INICIAL - Máximo 25 estudiantes por aula
		const aforoCicloI = getCellValue("B5");
		const aforoCicloII = getCellValue("B6");

		// Calcular aulas por ciclo (redondeo hacia arriba)
		const aulasCicloI = Math.ceil(aforoCicloI / 25);
		const aulasCicloII = Math.ceil(aforoCicloII / 25);
		const aulasInicial = aulasCicloI + aulasCicloII;
		const aforoInicial = aforoCicloI + aforoCicloII;

		console.log("=== EDUCACIÓN INICIAL ===");
		console.log("Ciclo I - Aforo:", aforoCicloI, "→ Aulas:", aulasCicloI);
		console.log("Ciclo II - Aforo:", aforoCicloII, "→ Aulas:", aulasCicloII);
		console.log(
			"Total Inicial - Aforo:",
			aforoInicial,
			"→ Aulas:",
			aulasInicial
		);

		// 📌 EDUCACIÓN PRIMARIA - Máximo 30 estudiantes por aula
		const aforo1roPrim = getCellValue("B8");
		const aforo2doPrim = getCellValue("B9");
		const aforo3roPrim = getCellValue("B10");
		const aforo4toPrim = getCellValue("B11");
		const aforo5toPrim = getCellValue("B12");
		const aforo6toPrim = getCellValue("B13");

		// Calcular aulas por grado (redondeo hacia arriba)
		const aulas1roPrim = Math.ceil(aforo1roPrim / 30);
		const aulas2doPrim = Math.ceil(aforo2doPrim / 30);
		const aulas3roPrim = Math.ceil(aforo3roPrim / 30);
		const aulas4toPrim = Math.ceil(aforo4toPrim / 30);
		const aulas5toPrim = Math.ceil(aforo5toPrim / 30);
		const aulas6toPrim = Math.ceil(aforo6toPrim / 30);

		const aulasPrimaria =
			aulas1roPrim +
			aulas2doPrim +
			aulas3roPrim +
			aulas4toPrim +
			aulas5toPrim +
			aulas6toPrim;
		const aforoPrimaria =
			aforo1roPrim +
			aforo2doPrim +
			aforo3roPrim +
			aforo4toPrim +
			aforo5toPrim +
			aforo6toPrim;

		console.log("\n=== EDUCACIÓN PRIMARIA ===");
		console.log("1° - Aforo:", aforo1roPrim, "→ Aulas:", aulas1roPrim);
		console.log("2° - Aforo:", aforo2doPrim, "→ Aulas:", aulas2doPrim);
		console.log("3° - Aforo:", aforo3roPrim, "→ Aulas:", aulas3roPrim);
		console.log("4° - Aforo:", aforo4toPrim, "→ Aulas:", aulas4toPrim);
		console.log("5° - Aforo:", aforo5toPrim, "→ Aulas:", aulas5toPrim);
		console.log("6° - Aforo:", aforo6toPrim, "→ Aulas:", aulas6toPrim);
		console.log(
			"Total Primaria - Aforo:",
			aforoPrimaria,
			"→ Aulas:",
			aulasPrimaria
		);

		// 📌 EDUCACIÓN SECUNDARIA - Máximo 30 estudiantes por aula
		const aforo1roSec = getCellValue("B15");
		const aforo2doSec = getCellValue("B16");
		const aforo3roSec = getCellValue("B17");
		const aforo4toSec = getCellValue("B18");
		const aforo5toSec = getCellValue("B19");

		// Calcular aulas por grado (redondeo hacia arriba)
		const aulas1roSec = Math.ceil(aforo1roSec / 30);
		const aulas2doSec = Math.ceil(aforo2doSec / 30);
		const aulas3roSec = Math.ceil(aforo3roSec / 30);
		const aulas4toSec = Math.ceil(aforo4toSec / 30);
		const aulas5toSec = Math.ceil(aforo5toSec / 30);

		const aulasSecundaria =
			aulas1roSec + aulas2doSec + aulas3roSec + aulas4toSec + aulas5toSec;
		const aforoSecundaria =
			aforo1roSec + aforo2doSec + aforo3roSec + aforo4toSec + aforo5toSec;

		console.log("\n=== EDUCACIÓN SECUNDARIA ===");
		console.log("1° - Aforo:", aforo1roSec, "→ Aulas:", aulas1roSec);
		console.log("2° - Aforo:", aforo2doSec, "→ Aulas:", aulas2doSec);
		console.log("3° - Aforo:", aforo3roSec, "→ Aulas:", aulas3roSec);
		console.log("4° - Aforo:", aforo4toSec, "→ Aulas:", aulas4toSec);
		console.log("5° - Aforo:", aforo5toSec, "→ Aulas:", aulas5toSec);
		console.log(
			"Total Secundaria - Aforo:",
			aforoSecundaria,
			"→ Aulas:",
			aulasSecundaria
		);

		// 📌 TOTALES
		const totalAulas = aulasInicial + aulasPrimaria + aulasSecundaria;
		const aforoMaximo = aforoInicial + aforoPrimaria + aforoSecundaria;

		console.log("\n=== TOTALES ===");
		console.log("Total de Aulas:", totalAulas);
		console.log("Aforo Máximo:", aforoMaximo);

		// 📌 Construir la respuesta (manteniendo la misma estructura)
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
				aforo_maximo: aforoMaximo,
			},
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
		console.error("Error procesando el archivo Excel:", error);
		return res.status(500).json({
			error: "Error al procesar el archivo",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};
