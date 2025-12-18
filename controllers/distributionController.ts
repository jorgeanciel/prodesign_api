import type { Request, Response } from "express";
import {
	saveProjectDistributionService,
	getProjectDistributionService,
	deleteProjectDistributionService,
	type DistributionData,
} from "../services/distributionService";

// Guardar distribución
export const saveProjectDistribution = async (req: Request, res: Response) => {
	try {
		const projectId = Number(req.params.id);
		const distributionData: DistributionData = req.body;

		console.log("📥 Recibiendo distribución para proyecto:", projectId);

		if (!projectId || isNaN(projectId)) {
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "ID de proyecto inválido",
			});
		}

		if (!distributionData || !distributionData.floors) {
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "Datos de distribución incompletos",
			});
		}

		const result = await saveProjectDistributionService(
			projectId,
			distributionData
		);

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("❌ Error en saveProjectDistribution:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al guardar distribución",
			error: error.toString(),
		});
	}
};

// Obtener distribución
export const getProjectDistribution = async (req: Request, res: Response) => {
	try {
		const projectId = Number(req.params.id);

		if (!projectId || isNaN(projectId)) {
			return res.status(400).json({
				statusCode: 400,
				success: false,
				message: "ID de proyecto inválido",
			});
		}

		const result = await getProjectDistributionService(projectId);

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
		console.error("❌ Error en getProjectDistribution:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al obtener distribución",
			error: error.toString(),
		});
	}
};

// Eliminar distribución
export const deleteProjectDistribution = async (
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

		const result = await deleteProjectDistributionService(projectId);

		res.status(200).json({
			statusCode: 200,
			...result,
		});
	} catch (error: any) {
		console.error("❌ Error en deleteProjectDistribution:", error);
		res.status(500).json({
			statusCode: 500,
			success: false,
			message: error.message || "Error al eliminar distribución",
			error: error.toString(),
		});
	}
};
