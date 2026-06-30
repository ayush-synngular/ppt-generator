const { execSync } = require("child_process");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");

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

describe("PPT Tool - M16 Error Handling & Validation", () => {

  test("M16-TC-001: Validate missing component type", () => {
    const output = runCommand(
      `node src/index.js create`
    );

    expect(output).toContain("Missing component sub-type");
  });

  test("M16-TC-002: Validate missing table rows", () => {
    const output = runCommand(
      `node src/index.js create table`
    );

    expect(output).toContain("Missing data array row matrix");
  });

  test("M16-TC-003: Validate invalid chart JSON", () => {
  const fs = require("fs");
  const outputFile = path.join(projectRoot, "test-output", "m16-invalid-chart.pptx");

  if (fs.existsSync(outputFile)) {
    fs.unlinkSync(outputFile);
  }

  runCommand(
    `node src/index.js create chart --data "invalid-json" --output test-output/m16-invalid-chart.pptx`
  );

  expect(fs.existsSync(outputFile)).toBe(false);
});

  test("M16-TC-004: Validate missing image path", () => {
    const output = runCommand(
      `node src/index.js create image`
    );

    expect(output).toContain("Missing image location");
  });

  test("M16-TC-005: Validate invalid slide number", () => {
    const output = runCommand(
      `node src/index.js delete-slide --input test-output/m03-text-basic.pptx --slide 0 --output test-output/m16-invalid-slide.pptx`
    );

    expect(output).toContain("Missing or invalid --slide");
  });

  test("M16-TC-006: Validate missing export input", () => {
    const output = runCommand(
      `node src/index.js export`
    );

    expect(output).toContain("Missing input file");
  });

  test("M16-TC-007: Validate unsupported export format", () => {
    const output = runCommand(
      `node src/index.js export --input test-output/m03-text-basic.pptx --format csv --output test-output/m16-invalid-export.csv`
    );

    expect(output).toContain("Unsupported format");
  });

});
