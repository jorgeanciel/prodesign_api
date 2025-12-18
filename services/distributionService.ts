import Project from "../models/mariadb/projects";

// Interfaz para los datos de distribución
export interface DistributionData {
	layoutMode: string;
	totalFloors: number;
	currentFloor: number;

	// Información de la distribución por piso
	floors: {
		[key: number]: {
			inicial: number;
			primaria: number;
			secundaria: number;
			primariaEnPabellonSecundaria?: number;
			secundariaEnPabellonPrimaria?: number;
			ambientesSuperiores?: any[];
			ambientesInicialLibre?: any[];
			ambientesPrimariaLibre?: any[];
			ambientesSecundariaLibre?: any[];
			ambientesReubicadosPrimaria?: any[];
			ambientesReubicadosSecundaria?: any[];
			distribucionCuadrante?: any;
		};
	};

	// Configuración de pabellones
	pabellonInferiorEs: string;
	pabellonIzquierdoEs?: string;
	pabellonDerechoEs?: string;

	// Ambientes
	ambientesEnPabellones: any[];
	ambientesLateralesCancha: any[];

	// Capacidad
	capacityInfo: {
		inicial: { total: number; max: number };
		primaria: { total: number; max: number; hasBiblioteca: boolean };
		secundaria: { total: number; max: number; hasLaboratorio: boolean };
		ambientesSuperiores?: any;
	};

	// Coordenadas y rectángulo
	coordinates?: any[];
	maxRectangle?: any;
}

// Guardar distribución
export const saveProjectDistributionService = async (
	projectId: number,
	distributionData: DistributionData
) => {
	try {
		console.log("🔍 Guardando distribución para proyecto:", projectId);

		const project = await Project.findByPk(projectId);

		if (!project) {
			throw new Error(`Proyecto con ID ${projectId} no encontrado`);
		}

		console.log("✅ Proyecto encontrado:", project.id, project.name);

		// Guardar distribución
		project.distribution_data = JSON.stringify(distributionData);
		project.distribution_saved_at = new Date();

		await project.save();

		console.log("💾 Distribución guardada exitosamente");

		return {
			success: true,
			message: "Distribución guardada exitosamente",
			projectId: project.id,
			projectName: project.name,
			savedAt: project.distribution_saved_at,
		};
	} catch (error) {
		console.error("❌ Error en saveProjectDistributionService:", error);
		throw error;
	}
};

// Obtener distribución guardada
export const getProjectDistributionService = async (projectId: number) => {
	try {
		const project = await Project.findByPk(projectId, {
			attributes: [
				"id",
				"name",
				"code",
				"distribution_data",
				"distribution_saved_at",
			],
		});

		if (!project) {
			throw new Error(`Proyecto con ID ${projectId} no encontrado`);
		}

		if (!project.distribution_data) {
			return {
				success: false,
				message: "Este proyecto aún no tiene distribución guardada",
				projectId: project.id,
				projectName: project.name,
				data: null,
			};
		}

		// Parsear JSON
		let distributionData;
		try {
			distributionData =
				typeof project.distribution_data === "string"
					? JSON.parse(project.distribution_data)
					: project.distribution_data;
		} catch (parseError) {
			console.error("Error al parsear distribution_data:", parseError);
			throw new Error("Datos de distribución corruptos");
		}

		return {
			success: true,
			projectId: project.id,
			projectName: project.name,
			savedAt: project.distribution_saved_at,
			data: distributionData,
		};
	} catch (error) {
		console.error("❌ Error en getProjectDistributionService:", error);
		throw error;
	}
};

// Eliminar distribución
export const deleteProjectDistributionService = async (projectId: number) => {
	try {
		const project = await Project.findByPk(projectId);

		if (!project) {
			throw new Error(`Proyecto con ID ${projectId} no encontrado`);
		}

		await project.update({
			distribution_data: null,
			distribution_saved_at: null,
		});

		return {
			success: true,
			message: "Distribución eliminada exitosamente",
			projectId: project.id,
		};
	} catch (error) {
		console.error("❌ Error en deleteProjectDistributionService:", error);
		throw error;
	}
};
