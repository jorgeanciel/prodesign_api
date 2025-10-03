import { Sequelize, Options } from "sequelize";

const dbname: string = process.env.DB_NAME || "";
const username: string = process.env.DB_USER || "";
const password: string = process.env.DB_PASS || "";
const options: Options = {
	host: process.env.DB_HOST || "",
	dialect: "mysql",
	port: parseInt(process.env.DB_PORT || "3306", 10),
	logging: false,
	dialectOptions: {
		ssl: {
			require: true,
			rejectUnauthorized: false,
		},
	},
	pool: {
		max: 5,
		min: 0,
		acquire: 30000,
		idle: 10000,
	},
};

const mariaDB = new Sequelize(dbname, username, password, options);

export default mariaDB;
