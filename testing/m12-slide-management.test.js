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

describe("PPT Tool - M12 Slide Management", () => {

  test("M12-TC-001: Add slide successfully", () => {
    const output = runCommand(
      `node src/index.js add-slide --input test-output/m03-text-basic.pptx --text "Automation Added Slide" --output test-output/m12-auto-add-slide.pptx`
    );

    expect(output).toContain("Done");

    const inspect = runCommand(
      `node src/index.js inspect test-output/m12-auto-add-slide.pptx`
    );

    expect(inspect).toContain("Slide count: 2");
  });

  test("M12-TC-002: Delete slide successfully", () => {
    const output = runCommand(
      `node src/index.js delete-slide --input test-output/m12-auto-add-slide.pptx --slide 2 --output test-output/m12-auto-delete-slide.pptx`
    );

    expect(output).toContain("Done");

    const inspect = runCommand(
      `node src/index.js inspect test-output/m12-auto-delete-slide.pptx`
    );

    expect(inspect).toContain("Slide count: 1");
  });

  test("M12-TC-003: Duplicate slide successfully", () => {
    const output = runCommand(
      `node src/index.js duplicate-slide --input test-output/m03-text-basic.pptx --slide 1 --output test-output/m12-auto-duplicate-slide.pptx`
    );

    expect(output).toContain("Done");

    const inspect = runCommand(
      `node src/index.js inspect test-output/m12-auto-duplicate-slide.pptx`
    );

    expect(inspect).toContain("Slide count: 2");
  });

  test("M12-TC-004: Move slide successfully", () => {
    const output = runCommand(
      `node src/index.js move-slide --input test-output/m12-auto-duplicate-slide.pptx --from 2 --to 1 --output test-output/m12-auto-move-slide.pptx`
    );

    expect(output).toContain("Done");

    const inspect = runCommand(
      `node src/index.js inspect test-output/m12-auto-move-slide.pptx`
    );

    expect(inspect).toContain("Slide count: 2");
  });

  test("M12-TC-005: Verify slide count command", () => {
    const output = runCommand(
      `node src/index.js slide-count --input test-output/m12-auto-move-slide.pptx`
    );

    expect(output).toContain("Slide count: 2");
  });

});