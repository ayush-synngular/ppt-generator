# PPT Generator Architecture Audit

## 1. Current modules

### Root exports / entrypoint
- `src/index.js`

### Core modules
- `src/core/Presentation.js`
- `src/core/Slide.js`
- `src/core/SlideBuilder.js`
- `src/core/TemplateEngine.js`
- `src/core/DataBinder.js`
- `src/core/AutoLayoutEngine.js`
- `src/core/ValidationEngine.js`
- `src/core/Theme.js`

### Import / reader / modifier
- `src/import/PptReader.js`
- `src/import/PptModifier.js`
- `src/import/models/PresentationModel.js`
- `src/import/models/SlideModel.js`
- `src/import/models/ShapeModel.js`

### Elements
- `src/elements/text.js`
- `src/elements/image.js`
- `src/elements/table.js`
- `src/elements/chart.js`
- `src/elements/shape.js`
- `src/elements/icon.js`

### Layouts
- `src/layouts/title-slide.js`
- `src/layouts/two-column.js`
- `src/layouts/comparison.js`

### Components
- `src/components/title.js`
- `src/components/bullet-list.js`
- `src/components/metric-card.js`
- `src/components/chart-card.js`

### Templates
- `src/templates/ExecutiveSummaryTemplate.js`
- `src/templates/ProjectStatusTemplate.js`
- `src/templates/QuarterlyBusinessReviewTemplate.js`

### Utilities
- `src/utils/color.js`
- `src/utils/export.js`
- `src/utils/positioning.js`
- `src/utils/sizing.js`

### Constants
- `src/constants/fonts.js`
- `src/constants/themes.js`

### Examples
- `examples/test-ppt-reader.js`
- `examples/test-ppt-modifier.js`
- `examples/test-ppt-modifier-save.js`

## 2. Dependencies between modules

### Core generation path
- `src/index.js` exports the core API surface.
- `src/core/Presentation.js` depends on `pptxgenjs`.
- `src/core/SlideBuilder.js` depends on layout modules:
  - `src/layouts/title-slide.js`
  - `src/layouts/two-column.js`
  - `src/layouts/comparison.js`
- `src/core/TemplateEngine.js` depends on `src/core/DataBinder.js`.
- Layout modules and component modules depend on `src/elements/text.js` for rendering text.
- `src/components/chart-card.js` depends on `src/elements/chart.js`.
- Templates depend on layouts and components to build slide sequences.

### Import and modification path
- `src/import/PptReader.js` depends on:
  - `adm-zip`
  - `xml2js`
  - `src/import/models/PresentationModel.js`
  - `src/import/models/SlideModel.js`
  - `src/import/models/ShapeModel.js`
- `src/import/PptModifier.js` depends on:
  - `src/core/Presentation.js`
  - `src/elements/text.js`

### Shared relationships
- `src/index.js` re-exports `PptModifier` alongside generation APIs.
- `src/import/PptModifier.js` is the only import module that directly depends on the core generation module `Presentation.js`.

## 3. Duplicate functionality

- `src/core/Presentation.js` / `src/core/Slide.js` and `src/import/models/PresentationModel.js` / `src/import/models/SlideModel.js` represent two separate presentation/slide models.
- `PptModifier.save()` rebuilds a PPTX from imported JSON using text-only slide reconstruction, which overlaps with generation behavior in `Presentation` + slide builder layouts, but does not reuse that existing creation path.
- `TemplateEngine` and `SlideBuilder` both provide mechanisms for turning structured config into slides, creating overlapping slide generation semantics.
- Multiple modules provide text placement logic:
  - layout modules
  - component modules
  - `PptModifier.save()` hardcoded text layout
- `ValidationEngine` validates raw JSON slide/component config, but there is no equivalent validator for imported `PresentationModel` data.

## 4. Missing integration points

- `PptReader` produces an imported model that is not mapped into the same `Presentation`/`Slide` object graph used by generation.
- The core `Presentation` API and the import reader output are disconnected; there is no adapter layer to transform imported slides into native slide objects.
- `PptModifier` only supports text replacement and PPTX export by rebuilding text slides, so imported shapes, images, charts, tables, or theme metadata are dropped.
- Validation is focused on config-driven generation and does not validate imported presentations or imported shape metadata.
- `src/import/models` are only used by the reader, and the rest of the system does not consume those models.
- Templates and slide builder components are not used by the import path and cannot produce a unified imported/exported presentation.
- There is no explicit shared schema for slides or presentation contents across generation, import, and modification.

## 5. Which modules operate on Presentation objects

### Directly operate on `Presentation` instances
- `src/core/Presentation.js`
- `src/core/SlideBuilder.js`
- `src/core/TemplateEngine.js`
- `src/templates/ExecutiveSummaryTemplate.js`
- `src/templates/ProjectStatusTemplate.js`
- `src/templates/QuarterlyBusinessReviewTemplate.js`
- `src/components/title.js`
- `src/components/bullet-list.js`
- `src/components/metric-card.js`
- `src/components/chart-card.js`
- `src/layouts/title-slide.js`
- `src/layouts/two-column.js`
- `src/layouts/comparison.js`
- `src/import/PptModifier.js` (uses `new Presentation(...)` for export)

### Indirectly related to presentation objects
- `src/elements/text.js`
- `src/elements/image.js`
- `src/elements/table.js`
- `src/elements/chart.js`
- `src/elements/shape.js`
- `src/elements/icon.js`

## 6. Which modules operate on raw JSON

- `src/import/PptReader.js`
- `src/import/PptModifier.js` (reads/writes JSON structure via `toJSON()`)
- `src/import/models/PresentationModel.js`
- `src/import/models/SlideModel.js`
- `src/import/models/ShapeModel.js`
- `src/core/TemplateEngine.js` / `src/core/DataBinder.js`
- `src/core/ValidationEngine.js`
- `src/core/SlideBuilder.js`
- `src/templates/*` (template input data is raw JSON-like config)

## 7. Which modules are disconnected from the read-modify-write pipeline

### Read-only or import-only modules
- `src/import/PptReader.js`
- `src/import/models/PresentationModel.js`
- `src/import/models/SlideModel.js`
- `src/import/models/ShapeModel.js`

### Generation-only modules
- `src/core/Presentation.js`
- `src/core/Slide.js`
- `src/core/SlideBuilder.js`
- `src/core/TemplateEngine.js`
- `src/core/DataBinder.js`
- `src/core/AutoLayoutEngine.js`
- `src/core/ValidationEngine.js`
- `src/core/Theme.js`
- `src/elements/*`
- `src/layouts/*`
- `src/components/*`
- `src/templates/*`
- `src/utils/*`
- `src/constants/*`

### Partial connector
- `src/import/PptModifier.js` connects import and export only via text extraction and text-based reconstruction; the rest of the generation pipeline is not reused.

## 8. Recommendations to unify everything under a single Presentation Model

1. Create one canonical presentation model.
   - Consolidate `src/core/Presentation` and `src/import/models/PresentationModel` into a single shared model interface.
   - Use the same slide model across generation, import, and modification.

2. Define a shared slide schema.
   - Include slide metadata, text items, elements, shape bounds, theme references, and object IDs.
   - Avoid separate `Slide` vs `SlideModel` representations.

3. Build adapters instead of duplicate models.
   - Add an import adapter that transforms `PptReader` output into the native presentation model.
   - Add an export adapter that writes the canonical model to PPTX via `Presentation`/`pptxgenjs`.

4. Treat the read-modify-write pipeline as one flow.
   - `PptReader` should output the canonical model.
   - `PptModifier` should modify the canonical model directly, not raw JSON.
   - `Presentation` or a dedicated writer should consume the canonical model and preserve layout/elements.

5. Reuse existing generators for export.
   - Instead of reconstructing text-only slides in `PptModifier.save()`, reuse layout and element renderers to preserve more structure.
   - Provide a generic `PresentationWriter` or `PptExporter` that understands shared slide element types.

6. Add schema validation for imported presentations.
   - Extend `ValidationEngine` to validate both generation config and imported presentation objects.
   - This will make the read/modify/export flow safer and more consistent.

7. Incrementally expand import coverage.
   - Start by preserving text and slide order.
   - Then add shapes, themes, images, tables, and charts using the shared model.
   - Ensure each stage uses the same abstraction layer.

8. Simplify the public API.
   - Export a unified model layer from `src/index.js`.
   - Keep `PptReader`, `PptModifier`, and `Presentation` as adapters on the same underlying presentation representation.
