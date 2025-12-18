import { Router } from "express";
import {
	getAllProjects,
	createProject,
	deleteProject,
	putProject,
	getProjectsByUserID,
	getProjectByID,
	test,
	createThumbnail,
	getProjectsCosts,
	updateProjectCosts,
} from "../../controllers/projectController";
import multer from "multer";
import fs from "fs";
import path from "path";
import projectErrorHandler from "../../middlewares/project-errorHandler";
import type { Request, Response, NextFunction } from "express";
import {
	deleteProjectPerimeters,
	getAllProjectsWithPerimeters,
	getProjectPerimeters,
	getProjectPerimetersCostSummary,
	saveProjectPerimeters,
} from "../../controllers/perimeterController";
import {
	deleteProjectDistribution,
	getProjectDistribution,
	saveProjectDistribution,
} from "../../controllers/distributionController";

const errPipe =
	(fn: Function) => (req: Request, res: Response, next: NextFunction) => {
		Promise.resolve(fn(req, res, next)).catch((err) => next(err)); // catch(next);
	};

const storage = multer.diskStorage({
	filename(req, file, callback) {
		callback(null, "thumbnail-" + req.params.id + ".png");
	},
	destination: "uploads/",
});
const upload = multer({ storage: storage });

const router = Router();

// RUTAS DE DISTRIBUCIÓN (agregar después de perímetros)
router.post("/:id/distribution", errPipe(saveProjectDistribution));
router.get("/:id/distribution", errPipe(getProjectDistribution));
router.delete("/:id/distribution", errPipe(deleteProjectDistribution));

//RUTAS DE PERÍMETROS (más específicas primero)
router.post("/:id/perimeters", errPipe(saveProjectPerimeters));
router.get(
	"/:id/perimeters/cost-summary",
	errPipe(getProjectPerimetersCostSummary)
);
router.get("/:id/perimeters", errPipe(getProjectPerimeters));
router.delete("/:id/perimeters", errPipe(deleteProjectPerimeters));
router.get("/perimeters/all", errPipe(getAllProjectsWithPerimeters));

// RUTAS DE COSTOS
router.get("/costs/:id", errPipe(getProjectsCosts));
router.put("/costs/:id", errPipe(updateProjectCosts));

// RUTAS DE THUMBNAIL
router.post(
	"/thumbnail/:id",
	upload.single("thumbnail"),
	errPipe(createThumbnail)
);
router.get("/thumbnail/:id", (req, res) => {
	const filePath = path.resolve(
		__dirname,
		"..",
		"..",
		"uploads",
		`thumbnail-${req.params.id}.png`
	);
	console.log(filePath);
	res.writeHead(200, {
		"Content-Type": "image/png",
	});
	const x = fs.createReadStream(filePath);
	x.on("error", (err) => {
		res.end("error");
	});
	x.pipe(res);
});

// RUTAS DE PROYECTOS ESPECÍFICOS
router.get("/test/dataRet", errPipe(test));
router.get("/id/:id", errPipe(getProjectByID));
router.get("/:id", errPipe(getProjectsByUserID));
router.put("/:id", errPipe(putProject));
router.delete("/:id", errPipe(deleteProject));

// RUTAS GENÉRICAS (al final)
router.get("/", errPipe(getAllProjects));
router.post("/", errPipe(createProject));

/* error handler */
router.use(projectErrorHandler);

export default router;
