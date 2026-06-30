const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");
const inputPpt = "test-output/m03-text-basic.pptx";
const imagePath = "node_modules/@jest/reporters/assets/jest_logo.png";

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

describe("PPT Tool - M14 Background & Existing Slide Image Editing", () => {
  test("M14-TC-001: Set slide background successfully", () => {
    const outputFile = path.join(outputDir, "m14-auto-background-red.pptx");

    const output = runCommand(
      `node src/index.js set-background --input ${inputPpt} --slide 1 --color FF0000 --output test-output/m14-auto-background-red.pptx`
    );

    expect(output).toContain("Done");
    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M14-TC-002: Add image to existing slide successfully", () => {
    const outputFile = path.join(outputDir, "m14-auto-image-basic.pptx");

    const output = runCommand(
      `node src/index.js add-image --input ${inputPpt} --slide 1 --path ${imagePath} --output test-output/m14-auto-image-basic.pptx`
    );

    expect(output).toContain("Done");
    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M14-TC-003: Add image with custom position and size successfully", () => {
    const outputFile = path.join(outputDir, "m14-auto-image-custom.pptx");

    const output = runCommand(
      `node src/index.js add-image --input ${inputPpt} --slide 1 --path ${imagePath} --x 4 --y 1 --w 2 --h 2 --output test-output/m14-auto-image-custom.pptx`
    );

    expect(output).toContain("Done");
    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M14-TC-004: Inspect edited PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m14-auto-image-basic.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("QA Text Test");
  });

  test("M14-TC-005: Handle missing parameter validation", () => {
    const bgOutput = runCommand(
      `node src/index.js set-background --input ${inputPpt} --slide 1`
    );

    const imageOutput = runCommand(
      `node src/index.js add-image --input ${inputPpt} --slide 1`
    );

    expect(bgOutput).toContain("Missing --color");
    expect(imageOutput).toContain("Missing --path");
  });
});
