const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");

function runCommand(command) {
  try {
    return execSync(command, {
      cwd: projectRoot,
      encoding: "utf8",
      stdio: ["pipe", "pipe", "pipe"],
    });
  } catch (error) {
    return `${error.stdout || ""}${error.stderr || ""}${error.message || ""}`;
  }
}

describe("PPT Tool - M13 Text Replacement", () => {
  test("M13-TC-001: Replace text successfully", () => {
    const outputFile = path.join(outputDir, "m13-auto-replace-basic.pptx");

    const output = runCommand(
      `node src/index.js replace "QA Text Test" "QA Replaced Text" test-output/m03-text-basic.pptx --output test-output/m13-auto-replace-basic.pptx`
    );

    expect(output).toContain("In-place replacement successful");
    expect(fs.existsSync(outputFile)).toBe(true);
  });

  test("M13-TC-002: Inspect replaced PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m13-auto-replace-basic.pptx`
    );

    expect(output).toContain("QA Replaced Text");
  });

  test("M13-TC-003: Replace text globally successfully", () => {
    const outputFile = path.join(outputDir, "m13-auto-replace-global.pptx");

    const output = runCommand(
      `node src/index.js replace-global --input test-output/m03-text-basic.pptx --old "QA" --new "Quality Assurance" --flags g --output test-output/m13-auto-replace-global.pptx`
    );

    expect(output).toContain("Done");
    expect(fs.existsSync(outputFile)).toBe(true);
  });

  test("M13-TC-004: Inspect global replaced PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m13-auto-replace-global.pptx`
    );

    expect(output).toContain("Quality Assurance Text Test");
  });

  test("M13-TC-005: Handle missing argument validation", () => {
    const replaceOutput = runCommand(`node src/index.js replace`);
    const replaceGlobalOutput = runCommand(`node src/index.js replace-global`);

    expect(replaceOutput).toContain("Usage");
    expect(replaceGlobalOutput).toContain("Missing --old");
  });
});
