import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const responsePath = process.argv[2];
if (!responsePath) {
  throw new Error("PageDrop response path is required");
}

const response = JSON.parse(await readFile(responsePath, "utf8"));
if (response.status !== "success" || !response.data?.url || !response.data?.siteId) {
  throw new Error(response.message || "PageDrop deployment failed");
}

const scriptsDir = path.dirname(fileURLToPath(import.meta.url));
const outputPath = path.join(path.dirname(scriptsDir), "data", "deployment.json");
const publicMetadata = {
  provider: "PageDrop",
  siteId: response.data.siteId,
  url: response.data.url,
  deployedAt: new Date().toISOString(),
  expiresAt: response.data.expiresAt ?? null,
};

await writeFile(outputPath, `${JSON.stringify(publicMetadata, null, 2)}\n`, "utf8");
console.log(publicMetadata.url);

