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

describe("PPT Tool - M03 Text Components", () => {
  test("M03-TC-001: Create basic text PPT successfully", () => {
    const outputFile = path.join(outputDir, "m03-auto-text-basic.pptx");

    runCommand(
      `node src/index.js create text --title "QA Text Test" --content "This is text component testing" --output test-output/m03-auto-text-basic.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M03-TC-002: Inspect generated text PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m03-auto-text-basic.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("QA Text Test");
  });

  test("M03-TC-003: Create bold text PPT successfully", () => {
    const outputFile = path.join(outputDir, "m03-auto-text-bold.pptx");

    runCommand(
      `node src/index.js create text --title "Bold Text Test" --content "This content should be visible" --bold true --output test-output/m03-auto-text-bold.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M03-TC-004: Create custom font size text PPT successfully", () => {
    const outputFile = path.join(outputDir, "m03-auto-text-fontsize.pptx");

    runCommand(
      `node src/index.js create text --title "Font Size Test" --content "This content uses custom font size" --fontSize 36 --output test-output/m03-auto-text-fontsize.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M03-TC-005: Create positioned text PPT successfully", () => {
    const outputFile = path.join(outputDir, "m03-auto-text-position.pptx");

    runCommand(
      `node src/index.js create text --title "Position Test" --content "This text uses custom position" --x 1 --y 1 --output test-output/m03-auto-text-position.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });
});