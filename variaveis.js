import dotenv from 'dotenv';
dotenv.config();

const email_Login = process.env.EMAIL || '';
const password_Login = process.env.PASSWORD || '';

export { email_Login, password_Login };
