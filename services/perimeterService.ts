import Project from "../models/mariadb/projects";
import { Op } from "sequelize";
import type { Request, Response } from "express";

// Interfaz para los datos de perímetros
export interface PerimeterData {
	distribution: {
		layoutMode: string;
		totalFloors: number;
		capacityInfo: {
			inicial: { total: number; max: number };
			primaria: { total: number; max: number; hasBiblioteca: boolean };
			secundaria: { total: number; max: number; hasLaboratorio: boolean };
		};
	};
	pabellones: {
		inicial: any;
		primaria: any;
		secundaria: any;
		superior: any;
	};
	ambientesCancha: any;
	resumenGeneral: {
		perimetroTotalColegio: string;
		totalAulas: number;
		totalAmbientes: number;
		totalBanos: number;
		totalEscaleras: number;
		areaConstruida: number;
		areaCancha: number;
	};
}

// Guardar o actualizar perímetros de un proyecto
export const saveProjectPerimetersService = async (
	projectId: number,
	perimetersData: PerimeterData
) => {
	try {
		const project = await Project.findByPk(projectId);

		if (!project) {
			throw new Error(`Proyecto con ID ${projectId} no encontrado`);
		}

		// Actualizar con los nuevos datos
		await project.update({
			perimeters_data: JSON.stringify(perimetersData),
			perimeters_calculated_at: new Date(),
		});

		return {
			success: true,
			message: "Perímetros guardados exitosamente",
			projectId: project.id,
			projectName: project.name,
			calculatedAt: project.perimeters_calculated_at,
		};
	} catch (error) {
		console.error("Error en saveProjectPerimetersService:", error);
		throw error;
	}
};

// Obtener perímetros de un proyecto
export const getProjectPerimetersService = async (projectId: number) => {
	try {
		const project = await Project.findByPk(projectId, {
			attributes: [
				"id",
				"name",
				"code",
				"perimeters_data",
				"perimeters_calculated_at",
			],
		});

		if (!project) {
			throw new Error(`Proyecto con ID ${projectId} no encontrado`);
		}

		if (!project.perimeters_data) {
			return {
				success: false,
				message: "Este proyecto aún no tiene perímetros calculados",
				projectId: project.id,
				projectName: project.name,
				data: null,
			};
		}

		// Parsear el JSON
		let perimetersData;
		try {
			perimetersData =
				typeof project.perimeters_data === "string"
					? JSON.parse(project.perimeters_data)
					: project.perimeters_data;
		} catch (parseError) {
			console.error("Error al parsear perimeters_data:", parseError);
			throw new Error("Datos de perímetros corruptos");
		}

		return {
			success: true,
			projectId: project.id,
			projectName: project.name,
			projectCode: project.code,
			calculatedAt: project.perimeters_calculated_at,
			data: perimetersData,
		};
	} catch (error) {
		console.error("Error en getProjectPerimetersService:", error);
		throw error;
	}
};

// Obtener resumen simplificado para costos
export const getProjectPerimetersCostSummaryService = async (
	projectId: number
) => {
	try {
		const result = await getProjectPerimetersService(projectId);

		if (!result.success || !result.data) {
			return {
				success: false,
				message: "No hay datos de perímetros disponibles",
				data: null,
			};
		}

		const data = result.data;

		// Estructura simplificada para costos
		const costSummary = {
			projectId: result.projectId,
			projectName: result.projectName,
			projectCode: result.projectCode,
			calculatedAt: result.calculatedAt,

			perimetros: {
				pabellonInicial: {
					total: parseFloat(
						data.pabellones?.inicial?.perimetroTotal || "0"
					),
					exterior: parseFloat(
						data.pabellones?.inicial?.perimetroExterior || "0"
					),
					divisiones: parseFloat(
						data.pabellones?.inicial?.divisionesInternas || "0"
					),
					elementos: data.pabellones?.inicial?.elementos || 0,
					desglose: data.pabellones?.inicial?.desglose || [],
				},
				pabellonPrimaria: {
					total: parseFloat(
						data.pabellones?.primaria?.perimetroTotal || "0"
					),
					exterior: parseFloat(
						data.pabellones?.primaria?.perimetroExterior || "0"
					),
					divisiones: parseFloat(
						data.pabellones?.primaria?.divisionesInternas || "0"
					),
					elementos: data.pabellones?.primaria?.elementos || 0,
					desglose: data.pabellones?.primaria?.desglose || [],
				},
				pabellonSecundaria: {
					total: parseFloat(
						data.pabellones?.secundaria?.perimetroTotal || "0"
					),
					exterior: parseFloat(
						data.pabellones?.secundaria?.perimetroExterior || "0"
					),
					divisiones: parseFloat(
						data.pabellones?.secundaria?.divisionesInternas || "0"
					),
					elementos: data.pabellones?.secundaria?.elementos || 0,
					desglose: data.pabellones?.secundaria?.desglose || [],
				},
				pabellonSuperior: {
					total: parseFloat(
						data.pabellones?.superior?.perimetroTotal || "0"
					),
					exterior: parseFloat(
						data.pabellones?.superior?.perimetroExterior || "0"
					),
					divisiones: parseFloat(
						data.pabellones?.superior?.divisionesInternas || "0"
					),
					elementos: data.pabellones?.superior?.elementos || 0,
					desglose: data.pabellones?.superior?.desglose || [],
				},
				ambientesCancha: {
					bottom: data.ambientesCancha?.bottom || null,
					top: data.ambientesCancha?.top || null,
					left: data.ambientesCancha?.left || null,
					right: data.ambientesCancha?.right || null,
					total: parseFloat(
						data.ambientesCancha?.totales?.perimetroTotal || "0"
					),
					elementos: data.ambientesCancha?.totales?.elementos || 0,
				},
			},

			resumen: data.resumenGeneral || {},

			// Total consolidado
			perimetroTotalGeneral: (
				parseFloat(data.pabellones?.inicial?.perimetroTotal || "0") +
				parseFloat(data.pabellones?.primaria?.perimetroTotal || "0") +
				parseFloat(data.pabellones?.secundaria?.perimetroTotal || "0") +
				parseFloat(data.pabellones?.superior?.perimetroTotal || "0") +
				parseFloat(data.ambientesCancha?.totales?.perimetroTotal || "0")
			).toFixed(2),
		};

		return {
			success: true,
			data: costSummary,
		};
	} catch (error) {
		console.error("Error en getProjectPerimetersCostSummaryService:", error);
		throw error;
	}
};

// Eliminar perímetros de un proyecto
export const deleteProjectPerimetersService = async (projectId: number) => {
	try {
		const project = await Project.findByPk(projectId);

		if (!project) {
			throw new Error(`Proyecto con ID ${projectId} no encontrado`);
		}

		await project.update({
			perimeters_data: null,
			perimeters_calculated_at: null,
		});

		return {
			success: true,
			message: "Perímetros eliminados exitosamente",
			projectId: project.id,
		};
	} catch (error) {
		console.error("Error en deleteProjectPerimetersService:", error);
		throw error;
	}
};

// Obtener todos los proyectos con perímetros calculados
export const getAllProjectsWithPerimetersService = async (userId?: number) => {
	try {
		const whereClause: any = {
			perimeters_data: {
				[Op.ne]: null,
			},
		};

		if (userId) {
			whereClause.user_id = userId;
		}

		const projects = await Project.findAll({
			where: whereClause,
			attributes: [
				"id",
				"name",
				"code",
				"thumbnail",
				"perimeters_calculated_at",
				"user_id",
			],
			order: [["perimeters_calculated_at", "DESC"]],
		});

		return {
			success: true,
			count: projects.length,
			projects: projects.map((p) => ({
				id: p.id,
				name: p.name,
				code: p.code,
				thumbnail: p.thumbnail,
				calculatedAt: p.perimeters_calculated_at,
				userId: p.user_id,
			})),
		};
	} catch (error) {
		console.error("Error en getAllProjectsWithPerimetersService:", error);
		throw error;
	}
};
