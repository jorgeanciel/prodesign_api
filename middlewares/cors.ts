import cors from "cors";

export const Cors = cors({
	origin: [
		"https://prodesign.pro-invest.pe",
		"https://apiprodesign.pro-invest.pe",
		"http://localhost:5199",
		"http://localhost:8005",
	],
	credentials: true,
	methods: ["GET", "POST", "PUT", "DELETE", "PATCH", "OPTIONS"],
	allowedHeaders: [
		"Origin",
		"X-Requested-With",
		"Content-Type",
		"Accept",
		"Authorization",
		"x-token",
	],
	optionsSuccessStatus: 200,
});
