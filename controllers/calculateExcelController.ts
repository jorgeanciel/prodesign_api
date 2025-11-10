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
	innovacion?: number;
	taller_creativo?: number;
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

// Interfaz para ambiente calculado
interface AmbienteCalculado {
	id: string;
	nombre: string;
	cantidad: number;
	area_unitaria: number;
	area_total: number;
	costo_unitario: number;
	costo_total: number;
}

// Interfaz para la respuesta completa
interface CalculoCostosResponse {
	success: boolean;
	message: string;
	data: {
		ambientes: AmbienteCalculado[];
		resumen: {
			total_ambientes: number;
			area_total_m2: number;
			costo_ambientes: number;
			costo_directo: number;
			gastos_generales: number;
			utilidad: number;
			subtotal: number;
			igv: number;
			presupuesto_total: number;
		};
		aulas: {
			inicial_ciclo1: number;
			inicial_ciclo2: number;
			primaria: number;
			secundaria: number;
			total: number;
		};
	};
}

/**
 * Actualiza el Excel y calcula costos en el backend
 */
export const updateProjectExcelAndCalculate = (
	req: Request,
	res: Response<CalculoCostosResponse>
) => {
	try {
		const projectData: ProjectExcelData = req.body;

		console.log("📥 Datos recibidos del frontend:", projectData);

		// 📌 PASO 1: Cargar IDEAS_PRODESIGN.xlsx
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);

		const consolidadoSheet = workbook.Sheets["CONSOLIDADO"];
		if (!consolidadoSheet) {
			return res.status(400).json({
				success: false,
				message: "No se encontró la hoja CONSOLIDADO",
			} as any);
		}

		// Función auxiliar para leer valores
		const getCellValue = (cellRef: string): number => {
			return consolidadoSheet[cellRef]?.v ?? 0;
		};

		// 📌 PASO 2: Leer áreas unitarias (E64-E86) ANTES de actualizar
		// Estas son las áreas por unidad/aula que están en el Excel
		const areasUnitarias: Record<number, number> = {};

		for (let i = 64; i <= 86; i++) {
			areasUnitarias[i] = getCellValue(`E${i}`);
		}

		console.log("📐 Áreas unitarias leídas del Excel:", areasUnitarias);

		// 📌 PASO 3: Actualizar cantidades (D64-D86)
		const cellMapping: Record<string, keyof ProjectExcelData> = {
			D64: "aulas_inicial_ciclo1",
			D65: "aulas_inicial_ciclo2",
			D66: "aula_psicomotricidad",
			D67: "aulas_primaria",
			D68: "aulas_secundaria",
			D69: "sum_inicial",
			D70: "biblioteca",
			D71: "innovacion",
			D72: "taller_creativo",
			D73: "taller_ept",
			D74: "laboratorio",
			D75: "sum_prim_sec",
			D76: "direccion_admin",
			D77: "sala_reuniones",
			D78: "sala_profesores",
			D79: "sshh_admin",
			D80: "cocina",
			D81: "sshh_cocina",
			D82: "depositos",
			D83: "canchas_deportivas",
			D84: "quiosco",
			D85: "topico",
			D86: "lactario",
		};

		// Actualizar las cantidades (D) y calcular áreas totales (F)
		const ambientesCalculados: AmbienteCalculado[] = [];
		const COSTO_M2 = 1848.29096045198; // Costo por m² (leer de F6 si quieres)

		Object.entries(cellMapping).forEach(([cellRef, dataKey]) => {
			const cantidad = projectData[dataKey];

			if (
				cantidad !== undefined &&
				cantidad !== null &&
				!isNaN(Number(cantidad))
			) {
				const cantidadNum = Number(cantidad);
				const fila = parseInt(cellRef.substring(1));

				// Actualizar cantidad (D)
				consolidadoSheet[cellRef] = {
					t: "n",
					v: cantidadNum,
				};

				// Obtener área unitaria de E
				const areaUnitaria = areasUnitarias[fila] || 0;

				// 🔥 CALCULAR F = D × E manualmente
				const areaTotal = cantidadNum * areaUnitaria;
				const cellF = `F${fila}`;

				consolidadoSheet[cellF] = {
					t: "n",
					v: areaTotal,
				};

				// Calcular costo
				const costoTotal = areaTotal * COSTO_M2;

				// Obtener nombre del ambiente
				const nombreCelda = consolidadoSheet[`B${fila}`];
				const nombre = nombreCelda?.v?.toString() || dataKey;

				// Agregar a resultados (solo si tiene cantidad > 0)
				if (cantidadNum > 0) {
					ambientesCalculados.push({
						id: dataKey,
						nombre: nombre,
						cantidad: cantidadNum,
						area_unitaria: areaUnitaria,
						area_total: areaTotal,
						costo_unitario: COSTO_M2,
						costo_total: costoTotal,
					});
				}

				console.log(
					`✅ ${nombre}: ${cantidadNum} × ${areaUnitaria}m² = ${areaTotal}m² (S/ ${costoTotal.toFixed(
						2
					)})`
				);
			}
		});

		// 📌 PASO 4: También actualizar celdas origen (D4, D5, D26, D43)
		const originCellMapping: Record<string, keyof ProjectExcelData> = {
			D4: "aulas_inicial_ciclo1",
			D5: "aulas_inicial_ciclo2",
			D26: "aulas_primaria",
			D43: "aulas_secundaria",
		};

		Object.entries(originCellMapping).forEach(([cellRef, dataKey]) => {
			const value = projectData[dataKey];
			if (value !== undefined && value !== null && !isNaN(Number(value))) {
				consolidadoSheet[cellRef] = {
					t: "n",
					v: Number(value),
				};
			}
		});

		// 📌 PASO 5: Guardar cambios
		xlsx.writeFile(workbook, excelPath);
		console.log("💾 Excel guardado con valores actualizados");

		// 📌 PASO 6: Calcular totales
		const totalArea = ambientesCalculados.reduce(
			(sum, amb) => sum + amb.area_total,
			0
		);

		const costoAmbientes = ambientesCalculados.reduce(
			(sum, amb) => sum + amb.costo_total,
			0
		);

		// Fórmulas de COSTO INFRA (J37-J42)
		const costoDirecto = costoAmbientes / 2; // Simplificado según Excel
		const gastosGenerales = costoDirecto * 0.1;
		const utilidad = costoDirecto * 0.1;
		const subtotal = costoDirecto + gastosGenerales + utilidad;
		const igv = subtotal * 0.18;
		const presupuestoTotal = subtotal + igv;

		// 📌 PASO 7: Preparar respuesta
		const response: CalculoCostosResponse = {
			success: true,
			message: "Costos calculados correctamente",
			data: {
				ambientes: ambientesCalculados,
				resumen: {
					total_ambientes: ambientesCalculados.length,
					area_total_m2: Math.round(totalArea * 100) / 100,
					costo_ambientes: Math.round(costoAmbientes * 100) / 100,
					costo_directo: Math.round(costoDirecto * 100) / 100,
					gastos_generales: Math.round(gastosGenerales * 100) / 100,
					utilidad: Math.round(utilidad * 100) / 100,
					subtotal: Math.round(subtotal * 100) / 100,
					igv: Math.round(igv * 100) / 100,
					presupuesto_total: Math.round(presupuestoTotal * 100) / 100,
				},
				aulas: {
					inicial_ciclo1: projectData.aulas_inicial_ciclo1 || 0,
					inicial_ciclo2: projectData.aulas_inicial_ciclo2 || 0,
					primaria: projectData.aulas_primaria || 0,
					secundaria: projectData.aulas_secundaria || 0,
					total:
						(projectData.aulas_inicial_ciclo1 || 0) +
						(projectData.aulas_inicial_ciclo2 || 0) +
						(projectData.aulas_primaria || 0) +
						(projectData.aulas_secundaria || 0),
				},
			},
		};

		console.log("✅ Cálculo completado:");
		console.log(`   - Total ambientes: ${response.data.ambientes.length}`);
		console.log(`   - Área total: ${response.data.resumen.area_total_m2} m²`);
		console.log(
			`   - Presupuesto: S/ ${response.data.resumen.presupuesto_total.toLocaleString()}`
		);

		return res.json(response);
	} catch (error) {
		console.error("❌ Error en updateProjectExcelAndCalculate:", error);
		return res.status(500).json({
			success: false,
			message: "Error al procesar el archivo Excel",
			data: {
				ambientes: [],
				resumen: {
					total_ambientes: 0,
					area_total_m2: 0,
					costo_ambientes: 0,
					costo_directo: 0,
					gastos_generales: 0,
					utilidad: 0,
					subtotal: 0,
					igv: 0,
					presupuesto_total: 0,
				},
				aulas: {
					inicial_ciclo1: 0,
					inicial_ciclo2: 0,
					primaria: 0,
					secundaria: 0,
					total: 0,
				},
			},
		} as any);
	}
};

/**
 * Obtiene la configuración base del Excel (áreas unitarias y etiquetas)
 * Útil para que el frontend sepa qué campos mostrar
 */
export const getExcelConfiguration = (req: Request, res: Response) => {
	try {
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);
		const consolidadoSheet = workbook.Sheets["CONSOLIDADO"];

		if (!consolidadoSheet) {
			return res.status(400).json({
				error: "No se encontró la hoja CONSOLIDADO",
			});
		}

		const getCellValue = (cellRef: string): any => {
			return consolidadoSheet[cellRef]?.v ?? null;
		};

		// Extraer configuración de filas 64-86
		const configuracion = [];

		for (let i = 64; i <= 86; i++) {
			const label = getCellValue(`B${i}`);
			const areaUnitaria = getCellValue(`E${i}`);

			if (label) {
				configuracion.push({
					fila: i,
					nombre: label,
					area_unitaria: areaUnitaria || 0,
					campo_id: label
						.toString()
						.toLowerCase()
						.replace(/\s+/g, "_")
						.replace(/[^a-z0-9_]/g, ""),
				});
			}
		}

		return res.json({
			success: true,
			data: {
				costo_m2: 1848.29096045198,
				ambientes: configuracion,
			},
		});
	} catch (error) {
		console.error("Error leyendo configuración:", error);
		return res.status(500).json({
			error: "Error al leer el archivo Excel",
			details: error instanceof Error ? error.message : String(error),
		});
	}
};

/**
 * Obtiene solo los valores actuales sin calcular nada
 */
export const getProjectExcelData = (req: Request, res: Response) => {
	try {
		const excelPath = path.resolve("uploads", "IDEAS_PRODESIGN.xlsx");
		const workbook = xlsx.readFile(excelPath);
		const sheet = workbook.Sheets["CONSOLIDADO"];

		const getCellValue = (cellRef: string) => sheet[cellRef]?.v ?? 0;

		const currentData: ProjectExcelData = {
			aulas_inicial_ciclo1: getCellValue("D64"),
			aulas_inicial_ciclo2: getCellValue("D65"),
			aula_psicomotricidad: getCellValue("D66"),
			aulas_primaria: getCellValue("D67"),
			aulas_secundaria: getCellValue("D68"),
			sum_inicial: getCellValue("D69"),
			biblioteca: getCellValue("D70"),
			innovacion: getCellValue("D71"),
			taller_creativo: getCellValue("D72"),
			taller_ept: getCellValue("D73"),
			laboratorio: getCellValue("D74"),
			sum_prim_sec: getCellValue("D75"),
			direccion_admin: getCellValue("D76"),
			sala_reuniones: getCellValue("D77"),
			sala_profesores: getCellValue("D78"),
			sshh_admin: getCellValue("D79"),
			cocina: getCellValue("D80"),
			sshh_cocina: getCellValue("D81"),
			depositos: getCellValue("D82"),
			canchas_deportivas: getCellValue("D83"),
			quiosco: getCellValue("D84"),
			topico: getCellValue("D85"),
			lactario: getCellValue("D86"),
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
