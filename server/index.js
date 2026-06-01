import express from "express";
import multer from "multer";
import axios from "axios";
import dotenv from "dotenv";
import crypto from "crypto";
import fs from "fs";
import path from "path";
import XLSX from "xlsx";
import { fileURLToPath } from "url";
import {
  createOvhTransporter,
  getOvhMailerConfig,
  loadRecipientsFromJson,
  sleep,
} from "./ovhMailer.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.join(__dirname, ".env") });

const PORT = Number(process.env.PORT || 5177);
const WC_URL = (process.env.WC_URL || process.env.WOO_SITE || "").replace(/\/$/, "");
const WC_KEY = process.env.WC_KEY || process.env.WOO_CK || "";
const WC_SECRET = process.env.WC_SECRET || process.env.WOO_CS || "";
const DEFAULT_TAG_NAME = "RAMADAN BUNDLES";
const ALLOWED_LOGINS = new Set([
  "a.elhabchi@flormar.ma",
  "y.bajou@flormar.ma",
  "a.chafa@flormar.ma",
]);
const SHARED_LOGIN_PASSWORD = "Ecom1243";
const authSessions = new Map();

if (!WC_URL || !WC_KEY || !WC_SECRET) {
  console.error("Missing WC_URL/WC_KEY/WC_SECRET in server/.env");
  process.exit(1);
}

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json({ limit: "10mb" }));

app.use("/api", (req, res, next) => {
  if (req.path === "/auth/login" || req.path === "/health") {
    return next();
  }

  const token = text(req.get("x-auth-token"));
  const session = authSessions.get(token);
  if (!session) {
    return res.status(401).json({ ok: false, error: "Authentication required." });
  }

  req.authUser = session;
  next();
});

const api = axios.create({
  baseURL: `${WC_URL}/wp-json/wc/v3`,
  auth: { username: WC_KEY, password: WC_SECRET },
  timeout: 60000,
});

function isTransientApiError(error) {
  const status = Number(error?.response?.status || 0);
  const code = text(error?.code).toUpperCase();
  const message = text(error?.message).toLowerCase();

  return (
    status >= 500 ||
    code === "ECONNRESET" ||
    code === "ETIMEDOUT" ||
    code === "ECONNABORTED" ||
    message.includes("socket hang up") ||
    message.includes("timeout")
  );
}

async function apiGetWithRetry(endpoint, config = {}, options = {}) {
  const retries = Math.max(0, Number(options.retries ?? 2) || 0);
  const retryDelayMs = Math.max(0, Number(options.retryDelayMs ?? 500) || 0);
  let attempt = 0;

  while (true) {
    try {
      return await api.get(endpoint, config);
    } catch (error) {
      if (attempt >= retries || !isTransientApiError(error)) {
        throw error;
      }
      attempt += 1;
      if (retryDelayMs > 0) {
        await sleep(retryDelayMs * attempt);
      }
    }
  }
}

const workflowCatalog = [
  // {
  //   id: "barcode-to-skus",
  //   title: "Barcodes To SKUs",
  //   category: "Lookups",
  //   endpoint: "/api/barcodes/resolve",
  //   method: "POST",
  //   description: "Resolve barcodes through productSkuCodeBar.json and download the matched SKUs.",
  //   fields: [
  //     { name: "barcodesText", label: "Barcodes", type: "textarea", placeholder: "One barcode per line" },
  //     { name: "barcodesFile", label: "Barcode File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
  //   ],
  // },
  {
    id: "get-user",
    title: "Get User",
    category: "Lookups",
    endpoint: "/api/users/get",
    method: "POST",
    description: "Fetch all WooCommerce customers with one click.",
    fields: [],
  },
  {
    id: "export-users",
    title: "Export Users",
    category: "Exports",
    endpoint: "/api/users/export",
    method: "POST",
    description: "Export all WooCommerce customers with one click.",
    fields: [],
  },
  {
    id: "send-bulk-email",
    title: "Send Bulk Email",
    category: "Email",
    endpoint: "/api/email/send-bulk",
    method: "POST",
    description: "Send an OVH SMTP campaign with subject, content, banner, attachment, and a pasted/uploaded user list. Stops on the first send error.",
    fields: [
      {
        name: "subject",
        label: "Subject",
        type: "text",
        placeholder: "Subject of the email",
        defaultValue: "Les 12 Heures d’Éclat arrivent sur flormar.ma - Offre exceptionnelle -50% dès 1000 dh d’achat",
      },
      {
        name: "content",
        label: "Content",
        type: "textarea",
        placeholder: "Email content",
        defaultValue:
          "Bonjour,\nPréparez-vous : Les 12 Heures d’Éclat arrivent sur flormar.ma.\nDu mercredi 05 mai à 18h, profitez d’une offre exceptionnelle et bénéficiez de -50% dès 1000 dh d’achat sur tous les produits.\nC’est l’occasion idéale de refaire votre trousse beauté, shopper vos indispensables ou découvrir de nouveaux produits.\nOffre valable exclusivement sur flormar.ma, dans la limite des stocks disponibles.\nRendez-vous mercredi à 18h.\nÀ très vite,\nFlormar",
      },
      {
        name: "bannerFile",
        label: "Banner",
        type: "file",
        accept: ".png,.jpg,.jpeg,.webp,.gif",
        help: "If empty, the app uses server/banner.jpeg automatically when available.",
      },
      { name: "attachmentFile", label: "Fichier Joint", type: "file", accept: ".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt,.png,.jpg,.jpeg,.webp,.zip" },
      { name: "usersText", label: "List Of Users", type: "textarea", placeholder: "One email per line, or separated by commas/semicolons" },
      { name: "usersFile", label: "Users File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      { name: "recipientsFile", label: "Fallback JSON File", type: "text", defaultValue: "customers.json", help: "Used only if no users are pasted or uploaded." },
      { name: "batchSize", label: "Batch Size", type: "number", defaultValue: "25" },
      { name: "delayMs", label: "Delay (ms)", type: "number", defaultValue: "1500" },
      {
        name: "stopOnError",
        label: "Stop On Error",
        type: "checkbox",
        defaultValue: true,
        help: "Stop immediately on the first error or SMTP sending limit error.",
      },
      {
        name: "dryRun",
        label: "Dry Run",
        type: "checkbox",
        defaultValue: true,
        help: "Preview the campaign without sending emails.",
      },
    ],
  },
  {
    id: "sku-descriptions",
    title: "SKU Descriptions",
    category: "Lookups",
    endpoint: "/api/sku/descriptions",
    method: "POST",
    description: "Fetch descriptions for products or variations by SKU.",
    fields: [
      { name: "skusText", label: "SKUs", type: "textarea", placeholder: "One SKU per line" },
      { name: "skusFile", label: "SKU File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      {
        name: "plainText",
        label: "Plain Text",
        type: "checkbox",
        defaultValue: true,
        help: "Strip HTML tags from descriptions.",
      },
    ],
  },
  {
    id: "export-skus",
    title: "Export All SKUs",
    category: "Exports",
    endpoint: "/api/sku/export",
    method: "POST",
    description: "Export all product and variation SKUs.",
    fields: [],
  },
  {
    id: "discount-from-barcodes",
    title: "Discount From Barcodes",
    category: "Pricing",
    endpoint: "/api/pricing/discount-from-barcodes",
    method: "POST",
    description: "Resolve barcodes and apply a percentage discount to matching variations.",
    fields: [
      { name: "barcodesText", label: "Barcodes", type: "textarea", placeholder: "One barcode per line" },
      { name: "barcodesFile", label: "Barcode File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      { name: "percentOff", label: "Percent Off", type: "number", defaultValue: "15" },
      {
        name: "clear",
        label: "Clear Sale Price",
        type: "checkbox",
        defaultValue: false,
        help: "Remove the sale price instead of applying a percentage.",
      },
      {
        name: "dryRun",
        label: "Dry Run",
        type: "checkbox",
        defaultValue: true,
        help: "Preview the changes without updating WooCommerce.",
      },
    ],
  },
  {
    id: "flat-price-from-barcodes",
    title: "Flat Price From Barcodes",
    category: "Pricing",
    endpoint: "/api/pricing/flat-price-from-barcodes",
    method: "POST",
    description: "Resolve barcodes and set one sale price for all matching variations.",
    fields: [
      { name: "barcodesText", label: "Barcodes", type: "textarea", placeholder: "One barcode per line" },
      { name: "barcodesFile", label: "Barcode File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      { name: "flatPrice", label: "Flat Price", type: "text", placeholder: "29.00" },
      {
        name: "dryRun",
        label: "Dry Run",
        type: "checkbox",
        defaultValue: true,
        help: "Preview the changes without updating WooCommerce.",
      },
    ],
  },
  {
    id: "tag-products",
    title: "Tag Products",
    category: "Tagging",
    endpoint: "/api/tag-products",
    method: "POST",
    description: "Apply a tag to products by SKU list or resolved barcodes.",
    fields: [
      { name: "tagName", label: "Tag Name", type: "text", defaultValue: DEFAULT_TAG_NAME },
      { name: "skusText", label: "SKUs", type: "textarea", placeholder: "One SKU per line" },
      { name: "skusFile", label: "SKU File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      { name: "barcodesText", label: "Barcodes", type: "textarea", placeholder: "One barcode per line" },
      { name: "barcodesFile", label: "Barcode File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      {
        name: "dryRun",
        label: "Dry Run",
        type: "checkbox",
        defaultValue: true,
        help: "Preview the changes without updating WooCommerce.",
      },
    ],
  },
  {
    id: "remove-tag",
    title: "Remove Tag",
    category: "Tagging",
    endpoint: "/api/remove-tag",
    method: "POST",
    description: "Remove a tag from selected products by SKU list or resolved barcodes.",
    fields: [
      { name: "tagName", label: "Tag Name", type: "text", defaultValue: DEFAULT_TAG_NAME },
      { name: "skusText", label: "SKUs", type: "textarea", placeholder: "One SKU per line" },
      { name: "skusFile", label: "SKU File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      { name: "barcodesText", label: "Barcodes", type: "textarea", placeholder: "One barcode per line" },
      { name: "barcodesFile", label: "Barcode File", type: "file", accept: ".txt,.csv,.json,.xlsx,.xls" },
      {
        name: "dryRun",
        label: "Dry Run",
        type: "checkbox",
        defaultValue: true,
        help: "Preview the changes without updating WooCommerce.",
      },
    ],
  },
  {
    id: "export-orders",
    title: "Export Orders",
    category: "Exports",
    endpoint: "/api/orders/export",
    method: "POST",
    description: "Export orders, items, SKUs, and optional notes/refunds.",
    fields: [
      {
        name: "datePreset",
        label: "Range",
        type: "select",
        defaultValue: "last-3-months",
        options: [
          { label: "Last 3 Months", value: "last-3-months" },
          { label: "Last 6 Months", value: "last-6-months" },
          { label: "Custom", value: "custom" },
        ],
      },
      { name: "dateFrom", label: "From", type: "date" },
      { name: "dateTo", label: "To", type: "date" },
      { name: "includeNotes", label: "Include Notes", type: "checkbox", defaultValue: false, help: "Generate order_notes.csv." },
      { name: "includeRefunds", label: "Include Refunds", type: "checkbox", defaultValue: false, help: "Generate order_refunds.csv." },
    ],
  },
  {
    id: "export-products",
    title: "Export Products",
    category: "Exports",
    endpoint: "/api/products/export",
    method: "POST",
    description: "Export products and variations into a CSV file.",
    fields: [],
  },
];

function text(value) {
  return String(value ?? "").replace(/^\uFEFF/, "").trim();
}

function truthy(value) {
  return ["1", "true", "yes", "y", "on"].includes(text(value).toLowerCase());
}

function unique(values) {
  return [...new Set(values.map(text).filter(Boolean))];
}

function normalizeEmail(value) {
  return text(value).toLowerCase();
}

function getField(req, name, fallback = "") {
  return req.body?.[name] ?? fallback;
}

function getFile(req, name) {
  const files = Array.isArray(req.files) ? req.files : [];
  return files.find((file) => file.fieldname === name) || null;
}

function parseTextList(value) {
  return unique(String(value || "").split(/[\r\n,;]+/));
}

function parseUploadedEntries(file) {
  if (!file) return [];
  const ext = path.extname(file.originalname || "").toLowerCase();

  if (ext === ".json") {
    const parsed = JSON.parse(file.buffer.toString("utf8"));
    if (Array.isArray(parsed)) {
      return unique(parsed.flatMap((item) => (typeof item === "object" && item !== null ? Object.values(item) : [item])));
    }
    if (parsed && typeof parsed === "object") {
      return unique(Object.values(parsed));
    }
    return unique([parsed]);
  }

  if (ext === ".xlsx" || ext === ".xls") {
    const workbook = XLSX.read(file.buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    return unique(rows.flatMap((row) => row));
  }

  return parseTextList(file.buffer.toString("utf8"));
}

function collectEntries(req, textField, fileField) {
  return unique([
    ...parseTextList(getField(req, textField, "")),
    ...parseUploadedEntries(getFile(req, fileField)),
  ]);
}

function csvEscape(value) {
  if (value === null || value === undefined) return "";
  const str = String(value);
  return /[",\n]/.test(str) ? `"${str.replace(/"/g, '""')}"` : str;
}

function rowsToCsv(rows) {
  return rows.map((row) => row.map(csvEscape).join(",")).join("\n") + "\n";
}

function htmlToText(value) {
  return String(value || "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function slugifyTag(value) {
  return text(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function loadBarcodeMap() {
  const filePath = path.join(__dirname, "productSkuCodeBar.json");
  const raw = fs.readFileSync(filePath, "utf8");
  const parsed = JSON.parse(raw);
  const rows = Array.isArray(parsed) ? parsed : [parsed];
  const map = new Map();

  for (const row of rows) {
    const barcode = text(
      row?.["Code à barre"] ??
        row?.["Code barre"] ??
        row?.["Code a barre"] ??
        row?.barcode ??
        row?.EAN,
    );
    const sku = text(
      row?.["Code article"] ??
        row?.["Code Article"] ??
        row?.code_article ??
        row?.sku ??
        row?.SKU,
    );
    if (barcode && sku) map.set(barcode, sku);
  }

  return map;
}

function loadSkuToBarcodesMap() {
  const barcodeMap = loadBarcodeMap();
  const skuToBarcodes = new Map();

  for (const [barcode, sku] of barcodeMap.entries()) {
    const current = skuToBarcodes.get(sku) || [];
    current.push(barcode);
    skuToBarcodes.set(sku, current);
  }

  return skuToBarcodes;
}

function resolveBarcodesToSkus(barcodes) {
  const map = loadBarcodeMap();
  const skus = [];
  const missing = [];
  const matches = [];

  for (const barcode of unique(barcodes)) {
    const sku = map.get(barcode);
    if (!sku) {
      missing.push(barcode);
      continue;
    }
    skus.push(sku);
    matches.push({ barcode, sku });
  }

  return { skus: unique(skus), missing, matches };
}

async function fetchAllPages(endpoint, params = {}, perPage = 100) {
  let page = 1;
  const out = [];

  while (true) {
    const { data, headers } = await api.get(endpoint, {
      params: { ...params, per_page: perPage, page },
    });
    const list = Array.isArray(data) ? data : [];
    out.push(...list);
    const totalPages = Number(headers["x-wp-totalpages"] || "1") || 1;
    if (page >= totalPages) break;
    page += 1;
  }

  return out;
}

async function fetchAllVariations(parentId) {
  return fetchAllPages(`/products/${parentId}/variations`, {}, 100);
}

async function resolveProductBySku(sku) {
  const wanted = text(sku);
  if (!wanted) return null;

  const variationIds = await findVariationBySku(wanted);
  if (variationIds) {
    const [parentResponse, variationResponse] = await Promise.all([
      api.get(`/products/${variationIds.productId}`),
      api.get(
        `/products/${variationIds.productId}/variations/${variationIds.variationId}`,
      ),
    ]);

    return {
      kind: "variation",
      parent: parentResponse.data,
      variation: variationResponse.data,
    };
  }

  const direct = await api.get("/products", {
    params: { sku: wanted, status: "any", per_page: 5 },
  });
  if (Array.isArray(direct.data) && direct.data.length) {
    return { kind: "product", product: direct.data[0] };
  }

  const candidates = await fetchAllPages("/products", {
    search: wanted,
    status: "any",
  });

  for (const parent of candidates) {
    if (parent.type !== "variable") continue;
    try {
      const vars = await api.get(`/products/${parent.id}/variations`, {
        params: { sku: wanted, status: "any", per_page: 100 },
      });
      if (Array.isArray(vars.data) && vars.data.length) {
        return { kind: "variation", parent, variation: vars.data[0] };
      }
    } catch {}
  }

  for (const parent of candidates) {
    if (parent.type !== "variable") continue;
    const allVars = await fetchAllVariations(parent.id);
    const hit = allVars.find((variation) => text(variation?.sku) === wanted);
    if (hit) return { kind: "variation", parent, variation: hit };
  }

  return null;
}

async function findVariationBySku(sku) {
  const wanted = text(sku);
  const { data } = await api.get("/products", {
    params: { sku: wanted, status: "any", per_page: 5 },
  });
  const directVariation = (data || []).find(
    (item) =>
      text(item?.sku) === wanted &&
      (item.type === "variation" || (item.parent_id && Number(item.parent_id) > 0)),
  );
  if (directVariation) {
    return { productId: directVariation.parent_id, variationId: directVariation.id };
  }

  const parents = await fetchAllPages("/products", {
    type: "variable",
    status: "any",
  });
  for (const parent of parents) {
    const vars = await fetchAllVariations(parent.id);
    const hit = vars.find((variation) => text(variation?.sku) === wanted);
    if (hit) {
      return { productId: parent.id, variationId: hit.id };
    }
  }

  return null;
}

async function getOrCreateTagId(tagName) {
  const slug = slugifyTag(tagName);
  const { data } = await api.get("/products/tags", {
    params: { search: tagName, per_page: 100, _fields: "id,name,slug" },
  });
  const match = (data || []).find(
    (tag) => text(tag?.name).toLowerCase() === text(tagName).toLowerCase() || text(tag?.slug) === slug,
  );
  if (match) return match.id;

  const created = await api.post("/products/tags", { name: tagName, slug });
  return created.data?.id;
}

function productHasTag(product, tagId, tagName) {
  const slug = slugifyTag(tagName);
  return (product.tags || []).some(
    (tag) =>
      tag.id === tagId ||
      text(tag?.name).toLowerCase() === text(tagName).toLowerCase() ||
      text(tag?.slug) === slug,
  );
}

function buildDateRange(preset, dateFrom, dateTo) {
  if (preset === "custom" && dateFrom && dateTo) {
    return {
      after: new Date(`${dateFrom}T00:00:00Z`).toISOString(),
      before: new Date(`${dateTo}T23:59:59Z`).toISOString(),
      label: `${dateFrom} to ${dateTo}`,
    };
  }

  const now = new Date();
  const start = new Date(now);
  start.setUTCMonth(start.getUTCMonth() - (preset === "last-6-months" ? 6 : 3));
  return {
    after: start.toISOString(),
    before: now.toISOString(),
    label: preset === "last-6-months" ? "Last 6 months" : "Last 3 months",
  };
}

async function fetchOrders(range) {
  return fetchAllPages("/orders", {
    status: "any",
    orderby: "date",
    order: "desc",
    after: range.after,
    before: range.before,
  });
}

function customerMatches(customer, query) {
  const wanted = text(query).toLowerCase();
  if (!wanted) return false;

  const haystacks = [
    customer?.id,
    customer?.email,
    customer?.username,
    customer?.first_name,
    customer?.last_name,
    [customer?.first_name, customer?.last_name].filter(Boolean).join(" "),
    customer?.billing?.first_name,
    customer?.billing?.last_name,
    [customer?.billing?.first_name, customer?.billing?.last_name].filter(Boolean).join(" "),
    customer?.billing?.email,
    customer?.billing?.phone,
    customer?.shipping?.first_name,
    customer?.shipping?.last_name,
    [customer?.shipping?.first_name, customer?.shipping?.last_name].filter(Boolean).join(" "),
  ]
    .map((value) => text(value).toLowerCase())
    .filter(Boolean);

  return haystacks.some((value) => value.includes(wanted));
}

function customerSnapshot(customer) {
  return {
    id: customer?.id || "",
    date_created: customer?.date_created || "",
    date_modified: customer?.date_modified || "",
    email: customer?.email || "",
    username: customer?.username || "",
    first_name: customer?.first_name || "",
    last_name: customer?.last_name || "",
    role: customer?.role || "",
    is_paying_customer: Boolean(customer?.is_paying_customer),
    avatar_url: customer?.avatar_url || "",
    billing: {
      first_name: customer?.billing?.first_name || "",
      last_name: customer?.billing?.last_name || "",
      company: customer?.billing?.company || "",
      address_1: customer?.billing?.address_1 || "",
      address_2: customer?.billing?.address_2 || "",
      city: customer?.billing?.city || "",
      state: customer?.billing?.state || "",
      postcode: customer?.billing?.postcode || "",
      country: customer?.billing?.country || "",
      email: customer?.billing?.email || "",
      phone: customer?.billing?.phone || "",
    },
    shipping: {
      first_name: customer?.shipping?.first_name || "",
      last_name: customer?.shipping?.last_name || "",
      company: customer?.shipping?.company || "",
      address_1: customer?.shipping?.address_1 || "",
      address_2: customer?.shipping?.address_2 || "",
      city: customer?.shipping?.city || "",
      state: customer?.shipping?.state || "",
      postcode: customer?.shipping?.postcode || "",
      country: customer?.shipping?.country || "",
    },
  };
}

function buildCustomersCsv(customers) {
  const rows = [
    [
      "id",
      "date_created",
      "date_modified",
      "email",
      "username",
      "first_name",
      "last_name",
      "role",
      "is_paying_customer",
      "billing_first_name",
      "billing_last_name",
      "billing_company",
      "billing_address_1",
      "billing_address_2",
      "billing_city",
      "billing_state",
      "billing_postcode",
      "billing_country",
      "billing_email",
      "billing_phone",
      "shipping_first_name",
      "shipping_last_name",
      "shipping_company",
      "shipping_address_1",
      "shipping_address_2",
      "shipping_city",
      "shipping_state",
      "shipping_postcode",
      "shipping_country",
    ],
  ];

  for (const customer of customers) {
    rows.push([
      customer?.id || "",
      customer?.date_created || "",
      customer?.date_modified || "",
      customer?.email || "",
      customer?.username || "",
      customer?.first_name || "",
      customer?.last_name || "",
      customer?.role || "",
      customer?.is_paying_customer ? "true" : "false",
      customer?.billing?.first_name || "",
      customer?.billing?.last_name || "",
      customer?.billing?.company || "",
      customer?.billing?.address_1 || "",
      customer?.billing?.address_2 || "",
      customer?.billing?.city || "",
      customer?.billing?.state || "",
      customer?.billing?.postcode || "",
      customer?.billing?.country || "",
      customer?.billing?.email || "",
      customer?.billing?.phone || "",
      customer?.shipping?.first_name || "",
      customer?.shipping?.last_name || "",
      customer?.shipping?.company || "",
      customer?.shipping?.address_1 || "",
      customer?.shipping?.address_2 || "",
      customer?.shipping?.city || "",
      customer?.shipping?.state || "",
      customer?.shipping?.postcode || "",
      customer?.shipping?.country || "",
    ]);
  }

  return rowsToCsv(rows);
}

function buildOrdersCsv(orders) {
  const rows = [
    ["id", "number", "date_created", "status", "currency", "total", "billing_name", "billing_email", "billing_phone", "billing_city", "shipping_city"],
  ];

  for (const order of orders) {
    const billingName = [order?.billing?.first_name, order?.billing?.last_name].filter(Boolean).join(" ");
    rows.push([
      order.id,
      order.number,
      order.date_created,
      order.status,
      order.currency,
      order.total,
      billingName,
      order?.billing?.email || "",
      order?.billing?.phone || "",
      order?.billing?.city || "",
      order?.shipping?.city || "",
    ]);
  }

  return rowsToCsv(rows);
}

function buildOrderItemsCsv(orders) {
  const rows = [["order_id", "order_number", "item_id", "name", "product_id", "variation_id", "sku", "quantity", "price", "subtotal", "total"]];

  for (const order of orders) {
    for (const item of order.line_items || []) {
      const sku =
        item?.sku ||
        (item?.meta_data || []).find((entry) => text(entry.key).toLowerCase() === "_sku")?.value ||
        "";
      rows.push([
        order.id,
        order.number,
        item.id,
        item.name,
        item.product_id,
        item.variation_id,
        sku,
        item.quantity,
        item.price ?? "",
        item.subtotal,
        item.total,
      ]);
    }
  }

  return rowsToCsv(rows);
}

function buildOrderSkusCsv(orders) {
  const rows = [["order_id", "order_number", "date_created", "status", "skus", "item_count"]];

  for (const order of orders) {
    const skus = [];
    let itemCount = 0;
    for (const item of order.line_items || []) {
      itemCount += Number(item?.quantity || 0);
      const sku =
        item?.sku ||
        (item?.meta_data || []).find((entry) => text(entry.key).toLowerCase() === "_sku")?.value ||
        "";
      if (sku) skus.push(sku);
    }
    rows.push([
      order.id,
      order.number,
      order.date_created,
      order.status,
      unique(skus).join(", "),
      itemCount,
    ]);
  }

  return rowsToCsv(rows);
}

async function buildOrderNotesCsv(orders) {
  const rows = [["order_id", "note_id", "date_created", "note", "customer_note", "added_by_user"]];
  const failures = [];
  for (const order of orders) {
    try {
      const { data } = await apiGetWithRetry(`/orders/${order.id}/notes`);
      for (const note of data || []) {
        rows.push([order.id, note.id, note.date_created, note.note, note.customer_note, note.added_by_user]);
      }
    } catch (error) {
      failures.push({
        orderId: order.id,
        orderNumber: order.number,
        resource: "notes",
        error: error.message || String(error),
      });
    }
  }
  return { csv: rowsToCsv(rows), failures };
}

async function buildOrderRefundsCsv(orders) {
  const rows = [["order_id", "refund_id", "date_created", "amount", "reason"]];
  const failures = [];
  for (const order of orders) {
    try {
      const { data } = await apiGetWithRetry(`/orders/${order.id}/refunds`);
      for (const refund of data || []) {
        rows.push([order.id, refund.id, refund.date_created, refund.amount, refund.reason || ""]);
      }
    } catch (error) {
      failures.push({
        orderId: order.id,
        orderNumber: order.number,
        resource: "refunds",
        error: error.message || String(error),
      });
    }
  }
  return { csv: rowsToCsv(rows), failures };
}

function joinNames(items) {
  return (items || []).map((item) => item?.name || item?.slug || item?.id).filter(Boolean).join(" | ");
}

function joinAttributes(items) {
  return (items || [])
    .map((item) => {
      const name = item?.name || "";
      const value = Array.isArray(item?.options) ? item.options.join("/") : item?.option || item?.options || "";
      return `${name}:${value}`;
    })
    .filter(Boolean)
    .join(" ; ");
}

function extractColor(attributes = []) {
  const hit = (attributes || []).find((item) => {
    const name = text(item?.name).toLowerCase();
    return name.includes("color") || name.includes("colour") || name.includes("couleur");
  });
  if (!hit) return "";
  if (Array.isArray(hit.options) && hit.options.length) return hit.options.join(" | ");
  return text(hit.option || hit.value || "");
}

function extractAttributeMap(attributes = []) {
  const out = {};
  for (const attribute of attributes || []) {
    const key = text(attribute?.name || attribute?.slug);
    if (!key) continue;
    const value = Array.isArray(attribute?.options)
      ? attribute.options.join(" | ")
      : text(attribute?.option || attribute?.value || "");
    out[key] = value;
  }
  return out;
}

function normalizeImages(images = []) {
  return (images || []).map((image) => ({
    id: image?.id || "",
    src: image?.src || "",
    name: image?.name || "",
    alt: image?.alt || "",
  }));
}

function normalizeCategories(items = []) {
  return (items || []).map((item) => ({
    id: item?.id || "",
    name: item?.name || "",
    slug: item?.slug || "",
  }));
}

function normalizeTags(items = []) {
  return (items || []).map((item) => ({
    id: item?.id || "",
    name: item?.name || "",
    slug: item?.slug || "",
  }));
}

function normalizeMetaData(items = []) {
  return (items || []).map((item) => ({
    id: item?.id || "",
    key: item?.key || "",
    value: item?.value ?? "",
  }));
}

function normalizeDimensions(item) {
  return {
    length: text(item?.dimensions?.length),
    width: text(item?.dimensions?.width),
    height: text(item?.dimensions?.height),
    weight: text(item?.weight),
  };
}

function extractCodeBars(product, parent, variation, fallbackCodeBars = []) {
  const values = unique([
    text(variation?.global_unique_id),
    text(product?.global_unique_id),
    text(parent?.global_unique_id),
    ...(fallbackCodeBars || []),
  ]);
  return values;
}

function productSnapshot({
  sku,
  codeBars,
  plainText,
  product,
  parent = null,
  variation = null,
}) {
  const isVariation = Boolean(parent && variation);
  const base = isVariation ? parent : product;
  const item = isVariation ? variation : product;
  const mergedAttributes = [
    ...(base?.attributes || []),
    ...(variation?.attributes || []),
  ];
  const descriptionValue = isVariation
    ? variation?.description || parent?.description || ""
    : product?.description || "";
  const shortDescriptionValue = isVariation
    ? parent?.short_description || ""
    : product?.short_description || "";
  const itemImages = isVariation
    ? [variation?.image, ...(parent?.images || [])].filter(Boolean)
    : product?.images || [];
  const dimensions = normalizeDimensions(item);

  return {
    sku,
    code_bars: extractCodeBars(product, parent, variation, codeBars),
    type: isVariation ? "variation" : "product",
    productId: isVariation ? parent.id : product.id,
    variationId: isVariation ? variation.id : "",
    parentId: isVariation ? parent.id : "",
    parent_sku: isVariation ? text(parent?.sku) : "",
    global_unique_id: text(item?.global_unique_id || base?.global_unique_id),
    product_type: base?.type || (isVariation ? "variable" : ""),
    name: base?.name || item?.name || "",
    variation_name: isVariation ? variation?.name || "" : "",
    slug: base?.slug || item?.slug || "",
    permalink: base?.permalink || item?.permalink || "",
    status: item?.status || base?.status || "",
    catalog_visibility: base?.catalog_visibility || "",
    featured: Boolean(base?.featured),
    virtual: Boolean(item?.virtual),
    downloadable: Boolean(item?.downloadable),
    purchasable: Boolean(item?.purchasable),
    on_sale: Boolean(item?.on_sale),
    total_sales: item?.total_sales ?? base?.total_sales ?? "",
    price: text(item?.price),
    regular_price: text(item?.regular_price),
    sale_price: text(item?.sale_price),
    date_on_sale_from_gmt: item?.date_on_sale_from_gmt || "",
    date_on_sale_to_gmt: item?.date_on_sale_to_gmt || "",
    stock_status: item?.stock_status || "",
    stock_quantity: item?.stock_quantity ?? "",
    manage_stock: Boolean(item?.manage_stock),
    backorders: item?.backorders || "",
    backorders_allowed: Boolean(item?.backorders_allowed),
    backordered: Boolean(item?.backordered),
    sold_individually: Boolean(base?.sold_individually),
    tax_status: item?.tax_status || base?.tax_status || "",
    tax_class: item?.tax_class || base?.tax_class || "",
    shipping_class: item?.shipping_class || base?.shipping_class || "",
    shipping_class_id: item?.shipping_class_id ?? base?.shipping_class_id ?? "",
    average_rating: item?.average_rating ?? base?.average_rating ?? "",
    rating_count: item?.rating_count ?? base?.rating_count ?? "",
    reviews_allowed: Boolean(base?.reviews_allowed),
    menu_order: item?.menu_order ?? base?.menu_order ?? "",
    color: extractColor(mergedAttributes),
    attributes: joinAttributes(base?.attributes),
    variation_attributes: joinAttributes(variation?.attributes),
    attribute_map: extractAttributeMap(mergedAttributes),
    categories: normalizeCategories(base?.categories),
    tags: normalizeTags(base?.tags),
    category_names: joinNames(base?.categories),
    tag_names: joinNames(base?.tags),
    featured_image: itemImages?.[0]?.src || "",
    image_count: itemImages.length,
    images: normalizeImages(itemImages),
    dimensions,
    description: plainText ? htmlToText(descriptionValue) : descriptionValue,
    short_description: plainText ? htmlToText(shortDescriptionValue) : shortDescriptionValue,
    raw_parent: parent || null,
    raw_product: isVariation ? null : product,
    raw_variation: variation || null,
    meta_data: normalizeMetaData(item?.meta_data || base?.meta_data),
  };
}

async function buildProductsCsv() {
  const rows = [[
    "row_type",
    "product_id",
    "parent_id",
    "product_type",
    "status",
    "sku",
    "name",
    "permalink",
    "price",
    "regular_price",
    "sale_price",
    "stock_status",
    "stock_quantity",
    "categories",
    "tags",
    "featured_image",
    "attributes",
    "variation_attributes",
  ]];

  const products = await fetchAllPages("/products", { status: "any" });
  for (const product of products) {
    rows.push([
      "product",
      product.id,
      "",
      product.type,
      product.status,
      product.sku,
      product.name,
      product.permalink,
      product.price,
      product.regular_price,
      product.sale_price,
      product.stock_status,
      product.stock_quantity,
      joinNames(product.categories),
      joinNames(product.tags),
      product.images?.[0]?.src || "",
      joinAttributes(product.attributes),
      "",
    ]);

    if (product.type === "variable") {
      const variations = await fetchAllVariations(product.id);
      for (const variation of variations) {
        rows.push([
          "variation",
          variation.id,
          product.id,
          "variation",
          variation.status || product.status,
          variation.sku,
          product.name,
          product.permalink,
          variation.price,
          variation.regular_price,
          variation.sale_price,
          variation.stock_status,
          variation.stock_quantity,
          joinNames(product.categories),
          joinNames(product.tags),
          variation.image?.src || "",
          joinAttributes(product.attributes),
          joinAttributes(variation.attributes),
        ]);
      }
    }
  }

  return { csv: rowsToCsv(rows), productCount: products.length, rowCount: rows.length - 1 };
}

app.get("/api/workflows", (_req, res) => {
  res.json({ ok: true, workflows: workflowCatalog });
});

app.post("/api/auth/login", (req, res) => {
  const email = normalizeEmail(req.body?.email);
  const password = text(req.body?.password);

  if (!ALLOWED_LOGINS.has(email) || password !== SHARED_LOGIN_PASSWORD) {
    return res.status(401).json({ ok: false, error: "Invalid email or password." });
  }

  const token = crypto.randomUUID();
  const session = { email, createdAt: new Date().toISOString() };
  authSessions.set(token, session);

  res.json({
    ok: true,
    token,
    user: session,
  });
});

app.post("/api/auth/logout", (req, res) => {
  const token = text(req.get("x-auth-token"));
  if (token) {
    authSessions.delete(token);
  }
  res.json({ ok: true });
});

app.post("/api/sku/lookup", upload.any(), async (req, res) => {
  try {
    const skus = collectEntries(req, "skusText", "skusFile");
    if (!skus.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one SKU." });
    }

    const results = [];
    for (const sku of skus) {
      try {
        const item = await resolveProductBySku(sku);
        if (!item) {
          results.push({ found: false, sku });
          continue;
        }
        if (item.kind === "product") {
          results.push({
            found: true,
            type: "product",
            sku,
            id: item.product.id,
            name: item.product.name,
            price: item.product.price,
            regular_price: item.product.regular_price,
            sale_price: item.product.sale_price,
            stock_status: item.product.stock_status,
            stock_quantity: item.product.stock_quantity,
            permalink: item.product.permalink,
          });
        } else {
          results.push({
            found: true,
            type: "variation",
            sku,
            parent: { id: item.parent.id, name: item.parent.name, permalink: item.parent.permalink },
            variation: {
              id: item.variation.id,
              price: item.variation.price,
              regular_price: item.variation.regular_price,
              sale_price: item.variation.sale_price,
              stock_status: item.variation.stock_status,
              stock_quantity: item.variation.stock_quantity,
              attributes: item.variation.attributes,
            },
          });
        }
      } catch (error) {
        results.push({ found: false, sku, error: error.message || String(error) });
      }
    }

    res.json({
      ok: true,
      summary: { requested: skus.length, found: results.filter((item) => item.found).length },
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/barcodes/resolve", upload.any(), async (req, res) => {
  try {
    const barcodes = collectEntries(req, "barcodesText", "barcodesFile");
    if (!barcodes.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one barcode." });
    }
    const resolved = resolveBarcodesToSkus(barcodes);
    const csvRows = [
      ["barcode", "sku"],
      ...resolved.matches.map((item) => [item.barcode, item.sku]),
    ];

    res.json({
      ok: true,
      summary: { requested: barcodes.length, mapped: resolved.skus.length, missing: resolved.missing.length },
      downloads: [
        { filename: "skus.txt", mimeType: "text/plain", content: resolved.skus.join("\n") + (resolved.skus.length ? "\n" : "") },
        { filename: "skus.json", mimeType: "application/json", content: JSON.stringify(resolved.skus, null, 2) },
        { filename: "barcode-to-sku.csv", mimeType: "text/csv", content: rowsToCsv(csvRows) },
        { filename: "barcode-to-sku.json", mimeType: "application/json", content: JSON.stringify(resolved.matches, null, 2) },
      ],
      missingBarcodes: resolved.missing,
      results: resolved.matches,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/sku/descriptions", upload.any(), async (req, res) => {
  try {
    const skus = collectEntries(req, "skusText", "skusFile");
    const plainText = truthy(getField(req, "plainText", true));
    const skuToBarcodes = loadSkuToBarcodesMap();
    if (!skus.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one SKU." });
    }

    const results = [];
    for (const sku of skus) {
      const resolved = await resolveProductBySku(sku);
      if (!resolved) {
        results.push({
          code_bars: skuToBarcodes.get(sku) || [],
          sku,
          type: "not_found",
          productId: "",
          variationId: "",
          name: "",
          variation_name: "",
          featured_image: "",
          image_count: 0,
          images: [],
          description: "",
          short_description: "",
        });
        continue;
      }

      if (resolved.kind === "product") {
        results.push(
          productSnapshot({
            sku,
            codeBars: skuToBarcodes.get(sku) || [],
            plainText,
            product: resolved.product,
          }),
        );
        continue;
      }

      results.push(
        productSnapshot({
          sku,
          codeBars: skuToBarcodes.get(sku) || [],
          plainText,
          parent: resolved.parent,
          variation: resolved.variation,
        }),
      );
    }

    const rows = [
      [
        "sku",
        "code_bars",
        "type",
        "parent_sku",
        "product_type",
        "name",
        "variation_name",
        "slug",
        "price",
        "regular_price",
        "stock_status",
        "stock_quantity",
        "color",
        "categories",
        "tags",
        "image_1",
        "image_2",
        "image_3",
        "image_4",
        "image_5",
        "image_6",
        "attributes",
        "description",
        "short_description",
      ],
      
      ...results.map((item) => [
        item.sku,
        (item.code_bars || []).join(" | "),
        item.type,
        item.parent_sku,
        item.product_type,
        item.name,
        item.variation_name,
        item.slug,
        item.price,
        item.regular_price,
        item.stock_status,
        item.stock_quantity,
        item.color,
        item.category_names,
        item.tag_names,
        item.images?.[0]?.src || "",
        item.images?.[1]?.src || "",
        item.images?.[2]?.src || "",
        item.images?.[3]?.src || "",
        item.images?.[4]?.src || "",
        item.images?.[5]?.src || "",
        item.attributes,
        item.description,
        item.short_description,
      ]),
    ];

    res.json({
      ok: true,
      summary: { requested: skus.length, results: results.length },
      downloads: [
        { filename: "sku-descriptions.csv", mimeType: "text/csv", content: rowsToCsv(rows) },
        { filename: "sku-descriptions.json", mimeType: "application/json", content: JSON.stringify(results, null, 2) },
      ],
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/users/get", upload.any(), async (req, res) => {
  try {
    const query = text(getField(req, "query", ""));
    if (!query) {
      const customers = await fetchAllPages("/customers", {
        orderby: "id",
        order: "asc",
      });
      const results = customers.map(customerSnapshot);

      return res.json({
        ok: true,
        summary: {
          customers: results.length,
        },
        downloads: [
          {
            filename: "customers.csv",
            mimeType: "text/csv",
            content: buildCustomersCsv(customers),
          },
          {
            filename: "customers.json",
            mimeType: "application/json",
            content: JSON.stringify(results, null, 2),
          },
        ],
        results,
      });
    }

    let exactById = null;
    if (/^\d+$/.test(query)) {
      try {
        const { data } = await api.get(`/customers/${query}`);
        exactById = data;
      } catch {}
    }

    let customers = exactById ? [exactById] : [];
    if (!customers.length) {
      const { data } = await api.get("/customers", {
        params: { search: query, per_page: 100 },
      });
      customers = Array.isArray(data) ? data : [];
    }

    const filtered = customers.filter((customer) => customerMatches(customer, query));
    const results = (filtered.length ? filtered : customers).map(customerSnapshot);

    res.json({
      ok: true,
      summary: {
        query,
        matches: results.length,
      },
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/users/export", async (_req, res) => {
  try {
    const customers = await fetchAllPages("/customers", {
      orderby: "id",
      order: "asc",
    });
    const results = customers.map(customerSnapshot);

    res.json({
      ok: true,
      summary: {
        customers: results.length,
      },
      downloads: [
        {
          filename: "customers.csv",
          mimeType: "text/csv",
          content: buildCustomersCsv(customers),
        },
        {
          filename: "customers.json",
          mimeType: "application/json",
          content: JSON.stringify(results, null, 2),
        },
      ],
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/email/send-bulk", upload.any(), async (req, res) => {
  try {
    const subject = text(getField(req, "subject", ""));
    const content = getField(req, "content", "");
    const uploadedBannerFile = getFile(req, "bannerFile");
    const attachmentFile = getFile(req, "attachmentFile");
    const recipientsFile = text(getField(req, "recipientsFile", "customers.json")) || "customers.json";
    const batchSize = Math.max(1, Number(getField(req, "batchSize", "25")) || 25);
    const delayMs = Math.max(0, Number(getField(req, "delayMs", "1500")) || 0);
    const dryRun = truthy(getField(req, "dryRun", true));
    const stopOnError = truthy(getField(req, "stopOnError", true));

    if (!subject) {
      return res.status(400).json({ ok: false, error: "Subject is required." });
    }
    if (!text(content)) {
      return res.status(400).json({ ok: false, error: "Content is required." });
    }

    const inputUsers = collectEntries(req, "usersText", "usersFile");
    let recipients = unique(inputUsers.map((value) => text(value).toLowerCase()));
    if (!recipients.length) {
      recipients = loadRecipientsFromJson(recipientsFile);
    }
    if (!recipients.length) {
      return res.status(400).json({ ok: false, error: "Provide a list of users or upload a users file." });
    }

    const defaultBannerCandidates = ["banner.jpeg", "banner.jpg", "banner.png", "banner.webp", "banner.gif"];
    const defaultBannerPath = defaultBannerCandidates
      .map((name) => path.join(__dirname, name))
      .find((candidate) => fs.existsSync(candidate));
    const bannerFile =
      uploadedBannerFile ||
      (defaultBannerPath
        ? {
            originalname: path.basename(defaultBannerPath),
            mimetype: `image/${path.extname(defaultBannerPath).slice(1).toLowerCase()}`,
            buffer: fs.readFileSync(defaultBannerPath),
          }
        : null);

    const plainText = text(content);
    const htmlParts = [];
    if (bannerFile) {
      htmlParts.push('<div style="margin-bottom:24px;"><img src="cid:campaign-banner" alt="Banner" style="max-width:100%;height:auto;display:block;border:0;" /></div>');
    }
    htmlParts.push(
      `<div style="font-family:Arial,sans-serif;font-size:16px;line-height:1.6;color:#222;">${escapeHtml(content).replace(/\r?\n/g, "<br />")}</div>`,
    );
    const html = htmlParts.join("");

    const attachments = [];
    if (bannerFile) {
      attachments.push({
        filename: bannerFile.originalname,
        content: bannerFile.buffer,
        contentType: bannerFile.mimetype || undefined,
        cid: "campaign-banner",
      });
    }
    if (attachmentFile) {
      attachments.push({
        filename: attachmentFile.originalname,
        content: attachmentFile.buffer,
        contentType: attachmentFile.mimetype || undefined,
      });
    }

    if (dryRun) {
      return res.json({
        ok: true,
        summary: {
          dryRun: true,
          recipients: recipients.length,
          source: inputUsers.length ? "input" : `file:${recipientsFile}`,
          batchSize,
          delayMs,
          stopOnError,
          banner: bannerFile?.originalname || "",
          attachment: attachmentFile?.originalname || "",
        },
        preview: recipients.slice(0, 25),
      });
    }

    const transporter = createOvhTransporter();
    const { from, host, port } = getOvhMailerConfig();
    await transporter.verify();

    const sent = [];
    const failed = [];

    for (let index = 0; index < recipients.length; index += batchSize) {
      const batch = recipients.slice(index, index + batchSize);

      for (const email of batch) {
        try {
          const info = await transporter.sendMail({
            from,
            to: email,
            subject,
            text: text(plainText) || undefined,
            html: text(html) || undefined,
            attachments: attachments.length ? attachments : undefined,
          });

          sent.push({
            email,
            messageId: info?.messageId || "",
          });
        } catch (error) {
              const failure = {
            email,
            error: error.message || String(error),
            sentCount: sent.length,
          };
          failed.push(failure);

          if (stopOnError) {
            return res.status(500).json({
              ok: false,
              error: `Stopped on send error for ${email}: ${failure.error}`,
              firstFailedEmail: email,
              errorMessage: failure.error,
              summary: {
                host,
                port,
                recipients: recipients.length,
                source: inputUsers.length ? "input" : `file:${recipientsFile}`,
                sent: sent.length,
                failed: failed.length,
                stopped: true,
                firstFailedEmail: email,
                banner: bannerFile?.originalname || "",
                attachment: attachmentFile?.originalname || "",
              },
              sent,
              failed,
            });
          }
        }
      }

      if (index + batchSize < recipients.length && delayMs > 0) {
        await sleep(delayMs);
      }
    }

    res.json({
      ok: true,
      summary: {
        host,
        port,
        recipients: recipients.length,
        source: inputUsers.length ? "input" : `file:${recipientsFile}`,
        sent: sent.length,
        failed: failed.length,
        stopped: false,
        firstFailedEmail: failed[0]?.email || "",
      },
      sent,
      failed,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/sku/export", async (_req, res) => {
  try {
    const products = await fetchAllPages("/products", {
      status: "any",
      _fields: "id,sku,type",
    });
    const skuSet = new Set();
    for (const product of products) {
      if (product?.sku) skuSet.add(String(product.sku));
    }

    const variableIds = products.filter((product) => product.type === "variable").map((product) => product.id);
    for (const productId of variableIds) {
      const variations = await fetchAllPages(`/products/${productId}/variations`, { _fields: "sku" });
      for (const variation of variations) {
        if (variation?.sku) skuSet.add(String(variation.sku));
      }
    }

    const skus = [...skuSet].sort((a, b) => a.localeCompare(b));
    res.json({
      ok: true,
      summary: { products: products.length, skus: skus.length },
      downloads: [
        { filename: "skus.txt", mimeType: "text/plain", content: skus.join("\n") + (skus.length ? "\n" : "") },
        { filename: "skus.json", mimeType: "application/json", content: JSON.stringify(skus, null, 2) },
      ],
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/pricing/discount-from-barcodes", upload.any(), async (req, res) => {
  try {
    const barcodes = collectEntries(req, "barcodesText", "barcodesFile");
    const percentOff = Number(getField(req, "percentOff", "15"));
    const clear = truthy(getField(req, "clear", false));
    const dryRun = truthy(getField(req, "dryRun", true));
    if (!barcodes.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one barcode." });
    }
    if (!clear && (!Number.isFinite(percentOff) || percentOff <= 0 || percentOff >= 100)) {
      return res.status(400).json({ ok: false, error: "Percent Off must be between 0 and 100." });
    }

    const resolved = resolveBarcodesToSkus(barcodes);
    const results = [];
    const stats = { changed: 0, skipped: 0, failed: 0 };

    for (const sku of resolved.skus) {
      try {
        const ids = await findVariationBySku(sku);
        if (!ids) {
          results.push({ ok: false, sku, reason: "variation_not_found" });
          stats.skipped++;
          continue;
        }

        let current = (await api.get(`/products/${ids.productId}/variations/${ids.variationId}`)).data;
        if (clear) {
          if (!dryRun) {
            current = (
              await api.put(`/products/${ids.productId}/variations/${ids.variationId}`, {
                sale_price: "",
                date_on_sale_from_gmt: null,
                date_on_sale_to_gmt: null,
              })
            ).data;
          }
          results.push({
            ok: true,
            sku,
            action: dryRun ? "would_clear" : "cleared",
            productId: ids.productId,
            variationId: ids.variationId,
            sale_price: current.sale_price,
          });
          stats.changed++;
          continue;
        }

        const regular = Number(current?.regular_price || 0);
        if (regular <= 0) {
          results.push({ ok: false, sku, reason: "missing_regular_price" });
          stats.skipped++;
          continue;
        }

        const salePrice = Number((regular * (1 - percentOff / 100)).toFixed(2));
        if (!dryRun) {
          current = (
            await api.put(`/products/${ids.productId}/variations/${ids.variationId}`, {
              sale_price: String(salePrice),
            })
          ).data;
        }
        results.push({
          ok: true,
          sku,
          action: dryRun ? "would_discount" : "discounted",
          productId: ids.productId,
          variationId: ids.variationId,
          regular_price: current.regular_price,
          sale_price: current.sale_price || String(salePrice),
        });
        stats.changed++;
      } catch (error) {
        results.push({ ok: false, sku, error: error.message || String(error) });
        stats.failed++;
      }
    }

    res.json({
      ok: true,
      summary: { requested: barcodes.length, resolvedSkus: resolved.skus.length, missingBarcodes: resolved.missing.length, changed: stats.changed },
      missingBarcodes: resolved.missing,
      stats,
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/pricing/flat-price-from-barcodes", upload.any(), async (req, res) => {
  try {
    const barcodes = collectEntries(req, "barcodesText", "barcodesFile");
    const flatPrice = text(getField(req, "flatPrice", ""));
    const dryRun = truthy(getField(req, "dryRun", true));
    if (!barcodes.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one barcode." });
    }
    if (!flatPrice) {
      return res.status(400).json({ ok: false, error: "Flat Price is required." });
    }

    const resolved = resolveBarcodesToSkus(barcodes);
    const results = [];
    const stats = { changed: 0, skipped: 0, failed: 0 };

    for (const sku of resolved.skus) {
      try {
        const ids = await findVariationBySku(sku);
        if (!ids) {
          results.push({ ok: false, sku, reason: "variation_not_found" });
          stats.skipped++;
          continue;
        }

        let current = (await api.get(`/products/${ids.productId}/variations/${ids.variationId}`)).data;
        if (!dryRun) {
          current = (
            await api.put(`/products/${ids.productId}/variations/${ids.variationId}`, {
              sale_price: flatPrice,
              date_on_sale_from_gmt: null,
              date_on_sale_to_gmt: null,
            })
          ).data;
        }
        results.push({
          ok: true,
          sku,
          action: dryRun ? "would_apply_flat_price" : "flat_price_applied",
          productId: ids.productId,
          variationId: ids.variationId,
          regular_price: current.regular_price,
          sale_price: current.sale_price || flatPrice,
        });
        stats.changed++;
      } catch (error) {
        results.push({ ok: false, sku, error: error.message || String(error) });
        stats.failed++;
      }
    }

    res.json({
      ok: true,
      summary: { requested: barcodes.length, resolvedSkus: resolved.skus.length, missingBarcodes: resolved.missing.length, changed: stats.changed },
      missingBarcodes: resolved.missing,
      stats,
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/tag-products", upload.any(), async (req, res) => {
  try {
    const tagName = text(getField(req, "tagName", DEFAULT_TAG_NAME));
    const dryRun = truthy(getField(req, "dryRun", true));
    const skus = collectEntries(req, "skusText", "skusFile");
    const resolved = resolveBarcodesToSkus(collectEntries(req, "barcodesText", "barcodesFile"));
    const allSkus = unique([...skus, ...resolved.skus]);
    if (!allSkus.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one SKU or barcode." });
    }

    const tagId = await getOrCreateTagId(tagName);
    const results = [];
    let changed = 0;
    let alreadyTagged = 0;

    for (const sku of allSkus) {
      try {
        const resolvedProduct = await resolveProductBySku(sku);
        if (!resolvedProduct) {
          results.push({ ok: false, sku, reason: "product_not_found" });
          continue;
        }
        const product = resolvedProduct.kind === "variation" ? resolvedProduct.parent : resolvedProduct.product;
        const fullProduct = (await api.get(`/products/${product.id}`, { params: { _fields: "id,name,sku,tags" } })).data;
        if (productHasTag(fullProduct, tagId, tagName)) {
          results.push({ ok: true, sku, action: "already_tagged", productId: fullProduct.id });
          alreadyTagged++;
          continue;
        }

        if (!dryRun) {
          const tagIds = [...new Set([...(fullProduct.tags || []).map((tag) => tag.id), tagId])];
          await api.put(`/products/${fullProduct.id}`, { tags: tagIds.map((id) => ({ id })) });
        }

        results.push({ ok: true, sku, action: dryRun ? "would_tag" : "tagged", productId: fullProduct.id });
        changed++;
      } catch (error) {
        results.push({ ok: false, sku, error: error.message || String(error) });
      }
    }

    res.json({
      ok: true,
      summary: { requested: allSkus.length, changed, alreadyTagged, missingBarcodes: resolved.missing.length },
      missingBarcodes: resolved.missing,
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/remove-tag", upload.any(), async (req, res) => {
  try {
    const tagName = text(getField(req, "tagName", DEFAULT_TAG_NAME));
    const dryRun = truthy(getField(req, "dryRun", true));
    const skus = collectEntries(req, "skusText", "skusFile");
    const resolved = resolveBarcodesToSkus(collectEntries(req, "barcodesText", "barcodesFile"));
    const allSkus = unique([...skus, ...resolved.skus]);
    if (!allSkus.length) {
      return res.status(400).json({ ok: false, error: "Provide at least one SKU or barcode." });
    }

    const slug = slugifyTag(tagName);
    const { data: tags } = await api.get("/products/tags", {
      params: { search: tagName, per_page: 100, _fields: "id,name,slug" },
    });
    const tag = (tags || []).find(
      (item) => text(item?.name).toLowerCase() === tagName.toLowerCase() || text(item?.slug) === slug,
    );
    if (!tag) {
      return res.status(404).json({ ok: false, error: `Tag "${tagName}" not found.` });
    }

    const results = [];
    let changed = 0;
    let alreadyClear = 0;

    for (const sku of allSkus) {
      try {
        const resolvedProduct = await resolveProductBySku(sku);
        if (!resolvedProduct) {
          results.push({ ok: false, sku, reason: "product_not_found" });
          continue;
        }

        const product = resolvedProduct.kind === "variation" ? resolvedProduct.parent : resolvedProduct.product;
        const fullProduct = (await api.get(`/products/${product.id}`, { params: { _fields: "id,name,sku,tags" } })).data;
        if (!productHasTag(fullProduct, tag.id, tagName)) {
          results.push({ ok: true, sku, action: "tag_not_present", productId: fullProduct.id });
          alreadyClear++;
          continue;
        }

        if (!dryRun) {
          const nextTags = (fullProduct.tags || []).filter((item) => item.id !== tag.id);
          await api.put(`/products/${fullProduct.id}`, {
            tags: nextTags.map((item) => ({ id: item.id })),
          });
        }

        results.push({ ok: true, sku, action: dryRun ? "would_remove_tag" : "tag_removed", productId: fullProduct.id });
        changed++;
      } catch (error) {
        results.push({ ok: false, sku, error: error.message || String(error) });
      }
    }

    res.json({
      ok: true,
      summary: { requested: allSkus.length, changed, alreadyClear, missingBarcodes: resolved.missing.length },
      missingBarcodes: resolved.missing,
      results,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/orders/export", async (req, res) => {
  try {
    const range = buildDateRange(
      text(getField(req, "datePreset", "last-3-months")),
      text(getField(req, "dateFrom", "")),
      text(getField(req, "dateTo", "")),
    );
    const includeNotes = truthy(getField(req, "includeNotes", false));
    const includeRefunds = truthy(getField(req, "includeRefunds", false));
    const orders = await fetchOrders(range);
    const warnings = [];

    const downloads = [
      { filename: "orders.csv", mimeType: "text/csv", content: buildOrdersCsv(orders) },
      { filename: "order_items.csv", mimeType: "text/csv", content: buildOrderItemsCsv(orders) },
      { filename: "orders_skus.csv", mimeType: "text/csv", content: buildOrderSkusCsv(orders) },
    ];

    if (includeNotes) {
      const notesExport = await buildOrderNotesCsv(orders);
      downloads.push({ filename: "order_notes.csv", mimeType: "text/csv", content: notesExport.csv });
      warnings.push(...notesExport.failures);
    }
    if (includeRefunds) {
      const refundsExport = await buildOrderRefundsCsv(orders);
      downloads.push({ filename: "order_refunds.csv", mimeType: "text/csv", content: refundsExport.csv });
      warnings.push(...refundsExport.failures);
    }

    res.json({
      ok: true,
      summary: {
        range: range.label,
        orders: orders.length,
        downloads: downloads.map((item) => item.filename),
        warnings: warnings.length,
      },
      downloads,
      warnings,
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.post("/api/products/export", async (_req, res) => {
  try {
    const exported = await buildProductsCsv();
    res.json({
      ok: true,
      summary: { products: exported.productCount, rows: exported.rowCount },
      downloads: [
        { filename: "products.csv", mimeType: "text/csv", content: exported.csv },
      ],
    });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message || String(error) });
  }
});

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, port: PORT });
});

const clientDistPath = path.resolve(__dirname, "../client/dist");
if (fs.existsSync(path.join(clientDistPath, "index.html"))) {
  app.use(express.static(clientDistPath));
  app.get("/", (_req, res) => {
    res.sendFile(path.join(clientDistPath, "index.html"));
  });
} else {
  app.get("/", (_req, res) => {
    res.json({ ok: true, workflows: workflowCatalog.length });
  });
}

app.listen(PORT, () => {
  console.log(`Ready on http://localhost:${PORT}`);
});
