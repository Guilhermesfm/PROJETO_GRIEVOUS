const dotenv = require("dotenv");
dotenv.config();

const email_Login = process.env.EMAIL || "";
const password_Login = process.env.PASSWORD || "";

module.exports = { email_Login, password_Login };
