import { request } from "./request";
import jwt from "jsonwebtoken";
export async function checkAuthMaster(action: AuthAction, body: any) {
	return new Promise((resolve, reject) => {
		const cr = request(
			`${process.env.APP_URL}/api/v1/${action}`,
			{
				method: "POST",
				headers: {
					"Content-Type": "application/x-www-form-urlencoded",
					// Authorization:
					// 	"Bearer " +
					// 	jwt.sign({ key: "AH3iH16PXS2" }, process.env.TOKEN_KEY),
					Authorization: `Bearer ${process.env.TOKEN_KEY}`,
				},
			},
			(res) => {
				var completeData = "";

				res.setEncoding("utf8");
				res.on("data", (chuck) => {
					completeData += chuck;
				});

				res.on("end", () => {
					if (!completeData.trim()) {
						reject(new Error("Respuesta vacía del servidor"));
						return;
					}

					try {
						resolve(JSON.parse(completeData));
					} catch (error) {
						console.error("Error al parsear JSON:", error);
						reject(new Error("La respuesta no es un JSON válido"));
					}
				});
			}
		);

		cr.on("error", (err) => {
			reject(err);
		});
		let params = [];
		for (let key of Object.keys(body)) {
			params.push(key + "=" + body[key]);
		}
		cr.end(params.join("&"));
	});
}
type AuthAction = "register" | "login";
