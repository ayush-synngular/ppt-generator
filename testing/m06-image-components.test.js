const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const projectRoot = path.resolve(__dirname, "..");
const outputDir = path.join(projectRoot, "test-output");
const assetDir = path.join(outputDir, "assets");
const imagePath = path.join(assetDir, "qa-image-test.png");

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
  fs.mkdirSync(assetDir, { recursive: true });

  if (!fs.existsSync(imagePath)) {
    const base64Png =
      "iVBORw0KGgoAAAANSUhEUgAAAZAAAAD6CAIAAADp6h5DAAAACXBIWXMAAAsTAAALEwEAmpwYAAADdUlEQVR4nO3bQW7bMBRA0XLuf+Z2uZQWVgQbC09YCSG4JJ7mDzvT6g0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOA9c7+qBwAAwLW3AgAgCwCAJAAAkAQAwLZb1Q8AwO7d5n42ns9nAQDg+NsBAJAFAEASAAApv1X9AADs3e7n4/l8FgAAjr8dAABZAAAkAQCQBADAtkvVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQCQBADAtkvVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAwHn6NwAAwN9zAADIAgAgCQCAX0/9AADs3e7n4/l8FgAAjr8dAABZAAAkAQDwn0rVDwDA7t3mfjaez2cBAOD42wEAkAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAvwGg2RZydXN0VAAAAABJRU5ErkJggg==";

    fs.writeFileSync(imagePath, Buffer.from(base64Png, "base64"));
  }
});

describe("PPT Tool - M06 Image Components", () => {
  test("M06-TC-001: Create image PPT successfully", () => {
    const outputFile = path.join(outputDir, "m06-auto-image-basic.pptx");

    runCommand(
      `node src/index.js create image --path test-output/assets/qa-image-test.png --output test-output/m06-auto-image-basic.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M06-TC-002: Create image PPT with custom position successfully", () => {
    const outputFile = path.join(outputDir, "m06-auto-image-position.pptx");

    runCommand(
      `node src/index.js create image --path test-output/assets/qa-image-test.png --x 5 --y 2 --output test-output/m06-auto-image-position.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M06-TC-003: Create image PPT with custom size successfully", () => {
    const outputFile = path.join(outputDir, "m06-auto-image-size.pptx");

    runCommand(
      `node src/index.js create image --path test-output/assets/qa-image-test.png --w 3 --h 2 --output test-output/m06-auto-image-size.pptx`
    );

    expect(fs.existsSync(outputFile)).toBe(true);
    expect(fs.statSync(outputFile).size).toBeGreaterThan(0);
  });

  test("M06-TC-004: Inspect generated image PPT successfully", () => {
    const output = runCommand(
      `node src/index.js inspect test-output/m06-auto-image-basic.pptx`
    );

    expect(output).toContain("Slide count: 1");
    expect(output).toContain("Slide titles");
  });

  test("M06-TC-005: Handle missing image path validation", () => {
    const output = runCommand(`node src/index.js create image`);

    expect(output).toContain("Missing image location");
  });
});
