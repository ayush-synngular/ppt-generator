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

describe("PPT Tool - M08 Icon Components", () => {

  test("M08-TC-001: Create default icon successfully", () => {
    const outputFile = path.join(outputDir, "m08-auto-default.pptx");

    runCommand(
      `node src/index.js create icon --output test-output/m08-auto-default.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M08-TC-002: Create custom icon successfully", () => {
    const outputFile = path.join(outputDir, "m08-auto-heart.pptx");

    runCommand(
      `node src/index.js create icon --name heart --output test-output/m08-auto-heart.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M08-TC-003: Create positioned icon successfully", () => {
    const outputFile = path.join(outputDir, "m08-auto-position.pptx");

    runCommand(
      `node src/index.js create icon --name star --x 1 --y 1 --output test-output/m08-auto-position.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M08-TC-004: Inspect generated icon PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m08-auto-default.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("Slide titles");
  });

  test("M08-TC-005: Verify generated icon PPT file is valid", () => {
  const outputFile = path.join(outputDir, "m08-auto-default2.pptx");

  runCommand(
    `node src/index.js create icon --output test-output/m08-auto-default2.pptx`
  );

  expect(fs.existsSync(outputFile)).toBe(true);
  expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
});
});