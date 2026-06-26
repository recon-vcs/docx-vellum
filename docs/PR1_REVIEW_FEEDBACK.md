# PR #1 Review Feedback

Target: https://github.com/recon-vcs/docx-vellum/pull/1

Scope checked so far:

- PR metadata: 123 files, +8341 / -6548, `main` -> `big-refactor`.
- Local build: `pnpm run build` passes.
- Local unit tests: `pnpm test` fails before most tests run because many tests import paths that no longer exist.

## 1. The test suite is not a usable quality gate

This PR moves core modules into the new `src/ooxml/...`, `src/rendering/...`, `src/shared/...` layout, but many unit tests still import the old paths. Example: `tests/unit/modern-page-splitter.spec.ts:2` imports `../../src/document/document`, `tests/unit/modern-page-splitter.spec.ts:4` imports `../../src/layout/modern-page-splitter`, and similar stale imports remain across pagination, length, utils, and parser tests.

`pnpm test` currently fails with 14 failed suites and only 2 passing suites. The failures are import-resolution failures, not behavioral assertions. That means the PR cannot claim the new parser/pagination/rendering behavior is protected by the added tests.

This should be fixed before judging the implementation quality. The right fix is not to patch one import at a time; the repo needs a deliberate module boundary/export policy for tests and internal code. If the new layout is intended, tests should import from the new stable internal paths or from explicit internal barrels. If compatibility paths are intended, add them intentionally and document that boundary.

## 2. The refactor splits files, but the renderer is still a stateful god object

`src/html-renderer-sync.ts` is smaller than before, but it still owns almost every rendering concern: document/style initialization, source path assignment, pagination, overflow splitting, header/footer rendering, notes, table context, tab stops, Konva image transformation state, and the callback wiring for every extracted renderer (`src/html-renderer-sync.ts:41`, `src/html-renderer-sync.ts:90`, `src/html-renderer-sync.ts:320`, `src/html-renderer-sync.ts:458`, `src/html-renderer-sync.ts:571`).

The extracted modules are mostly functions called through callback bags that close back over `HtmlRendererSync`. That is file extraction, not a clean architecture boundary. It keeps the hardest invariants implicit: `currentPage`, `currentPart`, `usedHeaderFooterParts`, `currentFootnoteIds`, `tableCtx`, and overflow state have to be valid by convention across many functions.

This is the main place where the implementation feels symptomatic. A more systematic design would make pagination/layout state an explicit object, make element rendering depend on a narrow rendering context, and keep DOM measurement/overflow as a separate phase or service. Without that, adding new element types will keep increasing callback surface area and shared mutable state.

## 3. Parser decomposition has the same facade/callback problem

`DocumentParser` now delegates to `paragraph-parser`, `run-parser`, `table-parser`, `style-parser`, `numbering-parser`, `drawing-parser`, and `math-parser`, but the central class still wires all parsing behavior through callback factories and still exposes broad methods for almost every OOXML concern (`src/ooxml/wordprocessingml/parsing/document-parser.ts:61`, `src/ooxml/wordprocessingml/parsing/document-parser.ts:138`, `src/ooxml/wordprocessingml/parsing/document-parser.ts:191`, `src/ooxml/wordprocessingml/parsing/document-parser.ts:199`, `src/ooxml/wordprocessingml/parsing/document-parser.ts:216`).

This gives the appearance of modularity while preserving a large implicit dependency graph. It will be hard to extend because adding one feature often means threading another callback through several layers instead of adding a cohesive parser for one OOXML concept.

The better direction is to define parser responsibilities around OOXML domains and shared parse context: relationships/options/xml utilities/style parsing should be explicit dependencies, not ad hoc callback bags. Then unit tests can target each domain parser without constructing the central parser facade.

## 4. Public API and documentation are inconsistent

The implementation returns `RenderResult` from both `renderDocument` and `renderSync` (`src/render.ts:14`, `src/render.ts:27`), but README still says `renderSync` returns the parsed `WordDocument` and the sample names the result `wordDocument` (`README.md:47`, `README.md:71`). The API table also lists `renderAsync`, but `src/docx-preview.ts` no longer exports `renderAsync`.

This is not a small docs typo. The PR changes the public surface while the docs describe a different one, so consumers cannot tell which contract is intentional. Before merging, decide whether `RenderResult` is the new API now or a future API, then align exports, README, type declarations, and compatibility notes.

## 5. Strict TypeScript is claimed, but important checks are still disabled

`tsconfig.json` sets `"strict": true`, then disables `noImplicitAny`, `strictNullChecks`, and `strictPropertyInitialization` (`tsconfig.json:10`, `tsconfig.json:15`, `tsconfig.json:16`, `tsconfig.json:18`). The code still relies on broad `any`, nullable fields, and convention-based initialization in central classes.

For this PR's size, that is risky because the refactor moved many files and introduced new boundaries, exactly where nullability and implicit any mistakes matter. If full strictness is not feasible yet, the PR should at least prevent new modules from depending on the old looseness. Otherwise the new architecture starts with the same type-safety debt as the legacy code.

## 6. Work-in-progress artifacts are mixed into the PR

`CLAUDE.md` and `real/zz.png` look like local review/workflow artifacts, not library source or test fixtures. `CLAUDE.md` contains agent instructions and a typo-like note about comparing `fixtures/a.docx` to `real/zz.png` (`CLAUDE.md:1`, `CLAUDE.md:3`). If a visual golden comparison is intended, it should be represented as an actual test fixture with a documented assertion flow. If not, these files should not be in the PR.

This matters because the PR is already very large. Mixing implementation, architecture movement, tests, docs, local instructions, and visual scratch assets makes it harder to review what is intentional.

## 7. New rendering helpers do not have a clean lifecycle

The new `RenderResult` always creates a `DomOverlayLayer`, which installs a `MutationObserver`, optional `ResizeObserver`, and a `window.resize` listener (`src/render-result.ts:75`, `src/render-result.ts:81`, `src/render-result.ts:93`, `src/render-result.ts:99`). Cleanup requires callers to know they must call `result.overlay.dispose()`.

That is a public API/lifecycle change, but existing examples ignore the return value (`README.md:62`), and browser tests type `renderSync` as returning `Promise<unknown>` and discard it. In practice this can leak observers/listeners every render, especially in an app that repeatedly previews documents into the same container.

The same issue exists in the drawing pipeline: `createKonva` appends a div with fixed id `konva-container` and creates a Konva stage using that global id (`src/rendering/dom/elements/drawing-renderer.ts:38`, `src/rendering/dom/elements/drawing-renderer.ts:40`, `src/rendering/dom/elements/drawing-renderer.ts:43`). Multiple renderers on the same page or repeated renders can collide, and there is no obvious `stage.destroy()` lifecycle.

This should be designed as an explicit renderer session/resource lifecycle. Either rendering must clean up previous internal resources itself, or the returned result must own all disposables and the docs/tests must enforce disposal. Fixed DOM ids should be avoided for library internals.

## 8. The added line-count test is not a meaningful architecture guard

`tests/unit/line-count.spec.ts` only enforces that each production `.ts` file is at most 800 lines. That check passes even though `src/html-renderer-sync.ts` is still 723 lines and still owns most rendering responsibilities. It also passes while the main unit test suite is unusable because of stale imports.

This is a weak proxy for maintainability. It can encourage mechanical file splitting while leaving the same coupling in place. If the goal is a systematic refactor, the quality gate should assert real boundaries: no stale compatibility imports, no renderer module reaching into parser internals, no global DOM ids, explicit disposable lifecycle, and focused tests around pagination/rendering contracts.
