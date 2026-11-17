import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";

const scriptDirectory = path.resolve("schedulr_apps_script");
const scriptFiles = fs
  .readdirSync(scriptDirectory)
  .filter((file) => file.endsWith(".gs"))
  .sort();

if (scriptFiles.length === 0) {
  throw new Error("No Apps Script files were found.");
}

for (const file of scriptFiles) {
  const fullPath = path.join(scriptDirectory, file);
  const source = fs.readFileSync(fullPath, "utf8");

  new vm.Script(source, {
    filename: fullPath,
  });

  console.log(`Syntax OK: ${file}`);
}