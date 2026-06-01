import fs from "fs";
import path from "path";
import nodemailer from "nodemailer";
import dotenv from "dotenv";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.join(__dirname, ".env") });

function text(value) {
  return String(value ?? "").trim();
}

function toBool(value, fallback = false) {
  const normalized = text(value).toLowerCase();
  if (!normalized) return fallback;
  return ["1", "true", "yes", "y", "on"].includes(normalized);
}

function requiredEnv(name) {
  const value = text(process.env[name]);
  if (!value) {
    throw new Error(`Missing required environment variable ${name}.`);
  }
  return value;
}

function firstEnv(...names) {
  for (const name of names) {
    const value = text(process.env[name]);
    if (value) {
      return value;
    }
  }
  return "";
}

function requiredAnyEnv(...names) {
  const value = firstEnv(...names);
  if (!value) {
    throw new Error(`Missing required environment variable ${names[0]}.`);
  }
  return value;
}

export function getOvhMailerConfig() {
  const fromEmail = requiredAnyEnv("OVH_MAIL_FROM", "FROM_EMAIL");
  const fromName = firstEnv("OVH_MAIL_FROM_NAME", "FROM_NAME");

  return {
    host: requiredAnyEnv("OVH_SMTP_HOST", "SMTP_HOST"),
    port: Number(firstEnv("OVH_SMTP_PORT", "SMTP_PORT") || 587),
    secure: toBool(firstEnv("OVH_SMTP_SECURE", "SMTP_SECURE"), false),
    auth: {
      user: requiredAnyEnv("OVH_SMTP_USER", "SMTP_USER"),
      pass: requiredAnyEnv("OVH_SMTP_PASS", "SMTP_PASS"),
    },
    from: fromName ? `"${fromName}" <${fromEmail}>` : fromEmail,
  };
}

export function createOvhTransporter() {
  const config = getOvhMailerConfig();
  return nodemailer.createTransport({
    host: config.host,
    port: config.port,
    secure: config.secure,
    auth: config.auth,
    requireTLS: !config.secure,
  });
}

export async function verifyOvhTransporter() {
  const transporter = createOvhTransporter();
  await transporter.verify();
  return transporter;
}

export function loadRecipientsFromJson(filePath = "customers.json") {
  const absolutePath = path.isAbsolute(filePath) ? filePath : path.join(__dirname, filePath);
  const raw = fs.readFileSync(absolutePath, "utf8");
  let parsed;

  try {
    parsed = JSON.parse(raw);
  } catch {
    parsed = raw.split(/\r?\n/);
  }

  if (!Array.isArray(parsed)) {
    throw new Error(`Recipients file must contain a JSON array or a plain text list.`);
  }

  const uniqueRecipients = [];
  const seen = new Set();

  for (const item of parsed) {
    const email = text(item).toLowerCase();
    if (!email || seen.has(email)) continue;
    seen.add(email);
    uniqueRecipients.push(email);
  }

  return uniqueRecipients;
}

export function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function sendOvhEmail({
  to,
  subject,
  text: plainText,
  html,
  attachments = [],
}) {
  const transporter = createOvhTransporter();
  const { from } = getOvhMailerConfig();

  return transporter.sendMail({
    from,
    to,
    subject,
    text: plainText,
    html,
    attachments,
  });
}
