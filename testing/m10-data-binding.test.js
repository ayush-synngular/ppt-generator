const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");
const dataDir = path.join(outputDir, "template-data");

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

beforeAll(() => {
  fs.mkdirSync(dataDir, { recursive: true });

  fs.writeFileSync(
    path.join(dataDir, "m10-auto-template.json"),
    JSON.stringify(
      {
        title: "{{title}}",
        subtitle: "{{subtitle}}",
        owner: "{{owner}}"
      },
      null,
      2
    )
  );

  fs.writeFileSync(
    path.join(dataDir, "m10-auto-data.json"),
    JSON.stringify(
      {
        title: "QA Automation Report",
        subtitle: "PPT Tool Testing",
        owner: "Diya Jain"
      },
      null,
      2
    )
  );
});

describe("PPT Tool - M10 Data Binding", () => {
  test("M10-TC-001: Bind template and data successfully", () => {
    const outputFile = path.join(outputDir, "m10-auto-bound-template.json");

    const output = runCommand(
      `node src/index.js bind --template test-output/template-data/m10-auto-template.json --data test-output/template-data/m10-auto-data.json --output test-output/m10-auto-bound-template.json`
    );

    expect(output).toContain("Data binding successful");
    expect(fs.existsSync(outputFile)).toBe(true);
  });

  test("M10-TC-002: Generate output JSON file successfully", () => {
    const outputFile = path.join(outputDir, "m10-auto-bound-template.json");

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M10-TC-003: Replace placeholders with data values successfully", () => {
    const outputFile = path.join(outputDir, "m10-auto-bound-template.json");
    const boundData = JSON.parse(fs.readFileSync(outputFile, "utf8"));

    expect(boundData.title).toBe("QA Automation Report");
    expect(boundData.subtitle).toBe("PPT Tool Testing");
    expect(boundData.owner).toBe("Diya Jain");
  });

  test("M10-TC-004: Handle missing template/data validation", () => {
    const output = runCommand("node src/index.js bind");

    expect(output).toContain("Missing template or data file");
  });
});