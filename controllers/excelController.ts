import { Request, Response } from "express";
import xlsx from "xlsx";

const normalizeNumber = (value: any): number => {
	if (value === null || value === undefined) return 0;

	const num = Number(String(value).replace(",", ".").trim());

	return isNaN(num) ? 0 : num;
};

const getCellNumber = (sheet: xlsx.WorkSheet, ref: string): number => {
	const cell = sheet[ref];
	return normalizeNumber(cell?.v);
};

const calcAulas = (aforo: number, maxPorAula: number): number => {
	if (aforo <= 0) return 0;
	return Math.ceil(aforo / maxPorAula);
};

export const readMatrizExcel = (req: Request, res: Response) => {
	try {
		if (!req.file) {
			return res.status(400).json({
				error: "No se ha subido ningún archivo",
			});
		}

		const workbook = xlsx.read(req.file.buffer, { type: "buffer" });
		const sheet = workbook.Sheets[workbook.SheetNames[0]];

		/* ===============================
		   EDUCACIÓN INICIAL
		================================ */
		const aforoCicloI = getCellNumber(sheet, "B5");
		const aforoCicloII = getCellNumber(sheet, "B6");

		const aforoInicial = aforoCicloI + aforoCicloII;
		const aulasInicial =
			aforoInicial > 0
				? calcAulas(aforoCicloI, 25) + calcAulas(aforoCicloII, 25)
				: 0;

		/* ===============================
		   EDUCACIÓN PRIMARIA
		================================ */
		const primariaCells = ["B8", "B9", "B10", "B11", "B12", "B13"];
		const aforosPrimaria = primariaCells.map((c) => getCellNumber(sheet, c));

		const aforoPrimaria = aforosPrimaria.reduce((sum, val) => sum + val, 0);

		const aulasPrimaria =
			aforoPrimaria > 0
				? aforosPrimaria.reduce((sum, val) => sum + calcAulas(val, 30), 0)
				: 0;

		/* ===============================
		   EDUCACIÓN SECUNDARIA
		================================ */
		const secundariaCells = ["B15", "B16", "B17", "B18", "B19"];
		const aforosSecundaria = secundariaCells.map((c) =>
			getCellNumber(sheet, c)
		);

		const aforoSecundaria = aforosSecundaria.reduce(
			(sum, val) => sum + val,
			0
		);

		const aulasSecundaria =
			aforoSecundaria > 0
				? aforosSecundaria.reduce((sum, val) => sum + calcAulas(val, 30), 0)
				: 0;

		/* ===============================
		   TOTALES
		================================ */
		const totalAulas = aulasInicial + aulasPrimaria + aulasSecundaria;

		const aforoMaximo = aforoInicial + aforoPrimaria + aforoSecundaria;

		return res.json({
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
				inicial: { ninos: 25, ninas: 25 },
				primaria: { ninos: 60, ninas: 60 },
				secundaria: { ninos: 60, ninas: 60 },
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
		});
	} catch (error) {
		console.error("Error procesando Excel:", error);
		return res.status(500).json({
			error: "Error al procesar el archivo",
		});
	}
};
