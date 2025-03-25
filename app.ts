import dotenv from "dotenv";
dotenv.config();
import Server from "./config/server";

console.log("URL de AUTH_SSO:", process.env.AUTH_SSO);

const server = new Server();
server.listen();
