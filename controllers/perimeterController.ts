import type { Request, Response } from "express";
import {
	saveProjectPerimetersService,
	getProjectPerimetersService,
	getProjectPerimetersCostSummaryService,
	deleteProjectPerimetersService,
	getAllProjectsWithPerimetersService,
	type PerimeterData,
} from "../services/perimeterService";

// Guardar perímetros de un proyecto
export const saveProjectPerimeters = async (req: Request, res: Response) => {
	try {
		const projectId = Number(req.params.id);
		const perimetersData: PerimeterData = req.body;

		// ✅ AGREGAR ESTOS LOGS
		console.log("📥 Recibiendo request para guardar perímetros:");
		console.log("  - Project ID:", projectId);
		console.log(
			"  - Body recibido:",
			JSON.stringify(perimetersData, null, 2)
		);
		console.log("  - Tiene pabellones?:", !!perimetersData?.pabellones);

		// Validación básica
		if (!projectId || isNaN(projectId)) {
			console.log("❌ ID de proyecto inválido");
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "ID de proyecto inválido",
			});
		}

		if (!perimetersData || !perimetersData.pabellones) {
			console.log("❌ Datos de perímetros incompletos");
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "Datos de perímetros incompletos",
			});
		}

		console.log("✅ Validaciones pasadas, llamando al service...");

		const result = await saveProjectPerimetersService(
			projectId,
			perimetersData
		);

		console.log("✅ Resultado del service:", result);

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("❌ Error en saveProjectPerimeters:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al guardar perímetros",
			error: error.toString(),
		});
	}
};

// Obtener perímetros de un proyecto
export const getProjectPerimeters = async (req: Request, res: Response) => {
	try {
		const projectId = Number(req.params.id);

		if (!projectId || isNaN(projectId)) {
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "ID de proyecto inválido",
			});
		}

		const result = await getProjectPerimetersService(projectId);

		if (!result.success) {
			return res.status(404).json({
				statusCode: 404,
				...result,
			});
		}

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("Error en getProjectPerimeters:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al obtener perímetros",
			error: error.toString(),
		});
	}
};

// Obtener resumen de costos
export const getProjectPerimetersCostSummary = async (
	req: Request,
	res: Response
) => {
	try {
		const projectId = Number(req.params.id);

		if (!projectId || isNaN(projectId)) {
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "ID de proyecto inválido",
			});
		}

		const result = await getProjectPerimetersCostSummaryService(projectId);

		if (!result.success) {
			return res.status(404).json({
				statusCode: 404,
				...result,
			});
		}

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("Error en getProjectPerimetersCostSummary:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al obtener resumen de costos",
			error: error.toString(),
		});
	}
};

// Eliminar perímetros
export const deleteProjectPerimeters = async (req: Request, res: Response) => {
	try {
		const projectId = Number(req.params.id);

		if (!projectId || isNaN(projectId)) {
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "ID de proyecto inválido",
			});
		}

		const result = await deleteProjectPerimetersService(projectId);

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("Error en deleteProjectPerimeters:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al eliminar perímetros",
			error: error.toString(),
		});
	}
};

// Obtener todos los proyectos con perímetros
export const getAllProjectsWithPerimeters = async (
	req: Request,
	res: Response
) => {
	try {
		const userId = req.query.user_id ? Number(req.query.user_id) : undefined;

		const result = await getAllProjectsWithPerimetersService(userId);

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("Error en getAllProjectsWithPerimeters:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al obtener proyectos",
			error: error.toString(),
		});
	}
};
