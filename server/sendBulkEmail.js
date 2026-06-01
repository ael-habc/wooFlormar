import fs from "fs";
import path from "path";
import dotenv from "dotenv";
import { fileURLToPath } from "url";
import {
  getOvhMailerConfig,
  loadRecipientsFromJson,
  sendOvhEmail,
  sleep,
  verifyOvhTransporter,
} from "./ovhMailer.js";

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

function readBodyFromEnv() {
  const textBody = text(process.env.MAIL_TEXT);
  const htmlBody = text(process.env.MAIL_HTML);
  const textFile = text(process.env.MAIL_TEXT_FILE);
  const htmlFile = text(process.env.MAIL_HTML_FILE);

  const resolvedText = textFile
    ? fs.readFileSync(path.isAbsolute(textFile) ? textFile : path.join(__dirname, textFile), "utf8")
    : textBody;
  const resolvedHtml = htmlFile
    ? fs.readFileSync(path.isAbsolute(htmlFile) ? htmlFile : path.join(__dirname, htmlFile), "utf8")
    : htmlBody;

  if (!text(resolvedText) && !text(resolvedHtml)) {
    throw new Error("Provide MAIL_TEXT, MAIL_HTML, MAIL_TEXT_FILE, or MAIL_HTML_FILE.");
  }

  return {
    text: resolvedText || undefined,
    html: resolvedHtml || undefined,
  };
}

async function main() {
  const { host, port, from } = getOvhMailerConfig();
  const subject = text(process.env.MAIL_SUBJECT);
  if (!subject) {
    throw new Error("MAIL_SUBJECT is required.");
  }

  const body = readBodyFromEnv();
  const recipientsFile = text(process.env.MAIL_RECIPIENTS_FILE || "customers.json");
  const recipients = loadRecipientsFromJson(recipientsFile);
  const batchSize = Math.max(1, Number(process.env.MAIL_BATCH_SIZE || 25));
  const delayMs = Math.max(0, Number(process.env.MAIL_DELAY_MS || 1500));
  const dryRun = toBool(process.env.MAIL_DRY_RUN, true);

  if (!recipients.length) {
    throw new Error("No recipients found.");
  }

  console.log(`OVH SMTP host: ${host}:${port}`);
  console.log(`From: ${from}`);
  console.log(`Recipients loaded: ${recipients.length}`);
  console.log(`Batch size: ${batchSize}`);
  console.log(`Delay: ${delayMs}ms`);
  console.log(`Dry run: ${dryRun ? "yes" : "no"}`);

  if (dryRun) {
    console.log("Dry run enabled. No emails were sent.");
    console.log("First recipients:", recipients.slice(0, 10));
    return;
  }

  await verifyOvhTransporter();

  let sent = 0;
  let failed = 0;

  for (let index = 0; index < recipients.length; index += batchSize) {
    const batch = recipients.slice(index, index + batchSize);

    for (const email of batch) {
      try {
        await sendOvhEmail({
          to: email,
          subject,
          text: body.text,
          html: body.html,
        });
        sent += 1;
        console.log(`Sent ${sent}/${recipients.length}: ${email}`);
      } catch (error) {
        failed += 1;
        console.error(`Failed for ${email}: ${error.message || String(error)}`);
      }
    }

    if (index + batchSize < recipients.length && delayMs > 0) {
      await sleep(delayMs);
    }
  }

  console.log(`Completed. Sent: ${sent}. Failed: ${failed}.`);
}

main().catch((error) => {
  console.error(error.message || String(error));
  process.exit(1);
});
