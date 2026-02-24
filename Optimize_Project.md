# The solo developer's playbook for two linked TypeScript repos

**yalc (not npm link), golden file tests, Release Please, and a surgically focused CLAUDE.md** form the foundation of an effective solo workflow for maintaining a public npm framework alongside a private Electron consumer app. The critical insight across all seven areas researched: every tool must earn its place by saving more time than it costs to maintain. For a single developer, the highest-ROI stack is yalc for local linking, `@xarsh/ooxml-validator` in every test, Renovate for cross-repo dependency flow, and Claude Code with progressive-disclosure documentation that gives AI genuine domain understanding.

This report covers local development linking, testing strategy, quality infrastructure, CI/CD, AI-assisted development, dependency management, and codebase health — each with specific configurations ready to implement.

---

## yalc wins the local linking battle for Electron

The choice of local development linking tool has outsized impact because electron-builder **breaks with symlinks**. Multiple electron-builder GitHub issues (#956, #3386, #6290) document ENOENT errors during asar packaging when symlinks exist in `node_modules`. This eliminates `npm link` entirely.

**yalc with `add` mode** (not `link` mode) is the recommended approach. It copies package files into a `.yalc/` directory and injects a `file:` reference — no symlinks touch `node_modules`. The feedback loop is fast: `tsc --watch` rebuilds the framework, `nodemon` watches `dist/` and runs `yalc push --changed`, and the consumer app picks up changes in **~2 seconds**.

| Approach         | Symlinks | electron-builder safe | Feedback speed | Setup effort |
| ---------------- | -------- | --------------------- | -------------- | ------------ |
| **yalc add**     | None     | ✅ Yes                | ~2 seconds     | Low          |
| npm pack + file: | None     | ✅ Yes                | ~15 seconds    | Low          |
| verdaccio        | None     | ✅ Yes                | ~15 seconds    | Medium-high  |
| **npm link**     | **Yes**  | ❌ **Breaks**         | Instant        | Very low     |

The daily workflow uses two terminals. Terminal 1 runs `npm run dev:full` in the framework (concurrently running `tsc --watch` and `yalc push`). Terminal 2 runs the Vite dev server in the Electron app after a one-time `yalc add docxmlater`. Changes propagate automatically.

**Before packaging for distribution**, always run `yalc remove --all && npm install` to restore the real npm version. Add this as a `prebuild:dist` script. The `.yalc/` directory and `yalc.lock` belong in `.gitignore` — accidentally committing a `file:.yalc/` reference will break CI.

For TypeScript declarations and source maps to work across the link, the framework's `tsconfig.json` needs `declaration: true`, `declarationMap: true`, `sourceMap: true`, and `inlineSources: true`. Include `src/**/*.ts` in the package.json `files` field so "Go to Definition" in the Electron app reaches the original TypeScript source.

**npm pack serves a different purpose**: pre-release validation. Before publishing a new framework version, `npm pack` creates a tarball identical to what npm would download. Install it in the app with `npm install ../docxmlater/docxmlater-10.0.3.tgz` and run the full test suite. This catches packaging issues (missing files, wrong exports) that yalc won't reveal.

---

## An inverted testing pyramid catches corruption bugs

For document processing libraries, the standard testing pyramid is **inverted**. Corruption bugs almost never occur in isolated units — they emerge at integration boundaries when XML fragments combine into invalid documents. The recommended ratio: **30% unit tests, 40% integration/golden file tests, 15% validation tests, 10% round-trip tests, 5% property-based tests**.

### Golden file testing is the highest-ROI pattern

The docxtemplater project's core testing approach — and the one most relevant to docxmlater — is golden file comparison. Generate a document, unzip both actual and expected `.docx` files, pretty-print each XML file, and compare strings. This catches corruption while tolerating whitespace differences.

```typescript
async function compareDocx(actual: Buffer, expected: Buffer): Promise<void> {
  const actualZip = await JSZip.loadAsync(actual);
  const expectedZip = await JSZip.loadAsync(expected);

  const actualFiles = Object.keys(actualZip.files).sort();
  const expectedFiles = Object.keys(expectedZip.files).sort();
  expect(actualFiles).toEqual(expectedFiles);

  for (const filename of actualFiles) {
    if (filename.endsWith('.xml') || filename.endsWith('.rels')) {
      const actualXml = normalizeXml(await actualZip.file(filename).async('text'));
      const expectedXml = normalizeXml(await expectedZip.file(filename).async('text'));
      expect(actualXml).toEqual(expectedXml);
    }
  }
}
```

Non-deterministic parts need normalization before comparison: strip `dcterms:created` and `dcterms:modified` timestamps, replace GUIDs with `{NORMALIZED-GUID}`, normalize `w:rsid*` revision attributes. Update golden files intentionally with `UPDATE_GOLDEN=true npx jest tests/golden/` and review the git diff.

### Every generated document should pass OOXML validation

**`@xarsh/ooxml-validator`** is the best Node.js option for schema validation. It uses a pre-built .NET binary (no runtime required), validates against official ECMA-376 schemas through Microsoft 365 version, and returns detailed errors with XPath locations.

```typescript
import { validateFile } from '@xarsh/ooxml-validator';

it('produces valid OOXML', async () => {
  const buffer = await generateDocument();
  fs.writeFileSync('/tmp/test.docx', buffer);
  const result = await validateFile('/tmp/test.docx');
  expect(result.ok).toBe(true);
});
```

**Integrate validation into every integration test** — not as a separate suite, but as a standard assertion. This single change will catch more corruption bugs than any other testing investment.

### Property-based tests with fast-check find edge cases humans miss

Property-based testing shines for invariants: "any valid text content round-trips correctly," "relationship IDs are always unique," "generated documents always pass OOXML validation." fast-check generates hundreds of random inputs per test run, finding edge cases in Unicode handling, extreme nesting depths, and empty-string boundary conditions.

### The "one bug, one regression test" discipline

Following docxtemplater and Apache POI's pattern: every corruption bug fix gets a dedicated regression test in `tests/regression/` that reproduces the exact scenario. Name the file after the issue number: `issue-234-missing-content-types.test.ts`. These tests are permanent — never delete them.

**Critical areas to cover in corruption tests**, ordered by risk:

- **`[Content_Types].xml`** — missing content types silently corrupt documents
- **Relationship files** — orphaned relationships or missing targets
- **Namespace declarations** — missing xmlns attributes
- **Element ordering** — OOXML schemas enforce strict child element order (even when Word tolerates violations)
- **Required attributes** — missing required attributes on elements

### Mutation testing: focused, not comprehensive

Full Stryker mutation testing on 75K lines would take **hours to days** and is impractical for a solo developer. Instead, run mutation testing only on the ~5-10 most critical files (packaging, relationship management, content types, ZIP assembly) before releases. Use `incremental: true` to speed up subsequent runs.

---

## Quality infrastructure that earns its keep

### ESLint: start with recommendedTypeChecked, not strict

For a 75K-line codebase with **zero existing linting**, `@typescript-eslint/recommended-type-checked` plus `stylistic-type-checked` is the right starting point. The `strict-type-checked` preset is explicitly documented as "not stable under semver" — rules and options may change in minor versions. Start lenient and tighten over time.

**Migration strategy for an unlinted codebase**: Set all rules to `warn` first. Run `eslint --format json` to get a baseline violation count. Use `eslint --fix` for auto-fixable rules. Then use `eslint --max-warnings N` in CI with a **decreasing threshold** over time — ratchet quality upward without blocking all development.

The rules that matter most for OOXML processing: **`no-floating-promises`** (critical for async JSZip operations), **`strict-boolean-expressions`** (prevents truthy/falsy bugs with XML string values like `""` being falsy), and **`no-fallthrough`** (for switch-case XML element dispatching).

### TypeScript strict mode: noUncheckedIndexedAccess is non-negotiable

**`noUncheckedIndexedAccess`** is the single most impactful TypeScript setting for XML processing code. When accessing `element.attributes['w:val']`, it adds `| undefined` to the return type, forcing null checks. Without it, `Record<string, T>` returns `T` instead of `T | undefined` — a silent corruption vector when XML attributes are missing.

Also enable `exactOptionalPropertyTypes` (distinguishes "property not set" from "property set to undefined" — critical for document fidelity) and `noPropertyAccessFromIndexSignature` (forces bracket notation for dynamic XML property access, making it explicit).

### Release Please automates changelogs with minimal overhead

For a solo developer who wants conventional commits → auto changelog → auto npm publish, **Release Please** wins. It parses git history, creates a release PR automatically, and you merge when ready. No changeset files to create (like Changesets requires), no complex plugin configuration (like semantic-release demands).

Changesets is designed for monorepos and teams — overkill for solo. semantic-release is powerful but the plugin configuration burden is significant. `commit-and-tag-version` (maintained fork of deprecated standard-version) is a simpler local-only alternative but doesn't create GitHub Releases.

### Knip catches dead code other tools miss

Knip detects unused files, exports, dependencies, types, and enum members. For a standalone library, configure it with `entry: ['src/index.ts']` and gradually adopt: start with `--include dependencies` (easiest wins), then add `unlisted`, then `exports`, then everything. Run in CI with `knip --no-progress`.

### dependency-cruiser enforces architecture without a team

Define rules like "utils/ should not import from domain code" and "no circular dependencies." The rules file lives in version control and CI enforces them permanently. This prevents the framework from growing unmaintainably complex — **the architecture enforces itself**.

For cognitive complexity thresholds, **20-25 is reasonable for XML/OOXML parsing code** (standard recommendation is 15, but XML parsing involves inherent branching on element types). Use the ESLint `complexity` rule at `warn` severity initially.

### Bundle size monitoring with size-limit

size-limit (by the PostCSS/autoprefixer author) simulates real consumer bundling and reports size. Set a budget in `package.json`, and the `andresz1/size-limit-action` GitHub Action posts size change comments on every PR. This prevents accidental bloat from sneaking in.

---

## CI/CD: four workflow files total

### docxmlater gets two workflows

**ci.yml** runs on every push and PR: lint, TypeScript type check, and Jest tests across a matrix of Node 18/20/22 × Ubuntu/Windows (DOCX users are often on Windows). Coverage uploads to Codecov.

**publish.yml** triggers on version tags (`v*`). It builds, runs the full test suite, and publishes to npm with `--provenance --access public`. The `id-token: write` permission is required for Sigstore provenance signing. Use either npm Trusted Publishing (requires npm CLI ≥11.5.1 / Node 24) or a classic `NPM_TOKEN` automation token.

```yaml
# Key publish step
- run: npm publish --provenance --access public
  env:
    NODE_AUTH_TOKEN: ${{ secrets.NPM_TOKEN }}
```

The `registry-url` parameter in `actions/setup-node` is **required** even though npmjs is the default — it triggers `.npmrc` creation. The `package.json` must include a `repository` field pointing to the GitHub repo for provenance to work.

### dochub-app gets two workflows plus Renovate

**ci.yml** runs lint and tests on Ubuntu only (save Windows runner minutes for release builds). Electron tests on headless Linux need `xvfb-run` or the `coactions/setup-xvfb` action.

**build.yml** triggers on tags or manual dispatch. Builds the MSI on `windows-latest`. Add `"postinstall": "electron-builder install-app-deps"` to `package.json` — this uses `@electron/rebuild` internally to rebuild better-sqlite3 and canvas against Electron's Node.js ABI. Known gotcha: WiX `light.exe` can fail with ICE validation errors on GitHub Actions runners; add `additionalWixArgs: ["-sval"]` to the MSI config as a workaround.

**Cost management**: The private repo gets 2,000 free GitHub Actions minutes/month. Windows runners bill at **2x multiplier**. Running lint/test on Linux only (~5 min/run × 30 pushes = 150 min) plus occasional Windows MSI builds (~10 min × 2x × 2 releases = 40 min) stays well under budget. Use `concurrency` groups with `cancel-in-progress: true` to avoid wasted runs.

### Cross-repo triggers: Renovate plus optional repository_dispatch

**Renovate** is the primary mechanism. Install the Renovate GitHub App on dochub-app. Configure it to automerge docxmlater updates immediately:

```json
{
  "packageRules": [
    {
      "matchPackageNames": ["docxmlater"],
      "automerge": true,
      "automergeType": "branch",
      "schedule": ["at any time"],
      "minimumReleaseAge": "0 days"
    }
  ]
}
```

With `automergeType: "branch"`, Renovate merges directly without creating a PR — zero noise. It detects new npm versions typically within an hour.

For **immediate** testing after publish, add a `repository_dispatch` step to docxmlater's publish workflow that curls the GitHub API to trigger a dochub-app workflow. This requires a fine-grained PAT scoped to the dochub-app repo with `contents: write` permission. The default `GITHUB_TOKEN` cannot trigger cross-repo dispatches.

**Branch strategy for solo developer**: Trunk-based development. Push directly to `main` for small changes. Use short-lived feature branches with PRs for larger changes. Tag releases from `main`. No `develop` branch — unnecessary overhead for one person.

---

## AI-assisted development that overcomes the OOXML knowledge gap

### CLAUDE.md should be under 100 lines with pointers elsewhere

The fundamental insight from community research: Claude Code's system prompt already contains ~50 instructions. Frontier LLMs reliably follow **150-200 instructions total**. Your CLAUDE.md competes for that budget. Furthermore, Claude Code wraps CLAUDE.md with a system note saying it "may or may not be relevant" — meaning bloated files get actively ignored.

Structure CLAUDE.md around **WHAT/WHY/HOW** in under 100 lines:

```markdown
# docxmlater - OOXML/DOCX Processing Framework

## What This Is

TypeScript framework for reading/writing OOXML (.docx) files per ECMA-376.

## Commands

- `npm run build` / `npm test` / `npm run typecheck`

## Architecture

- src/core/: XML parsing, OPC package handling
- src/wml/: WordprocessingML element models
- src/api/: Public API surface

## Critical Rules

- All XML element classes extend BaseElement
- Changes to core/ require full test suite
- New WML elements need: model, serializer, deserializer, tests
- See @agent_docs/ for deep context
```

### Progressive disclosure through agent_docs/ is the highest-impact change

Instead of stuffing everything into CLAUDE.md, create an `agent_docs/` directory with focused documents that Claude reads **on demand**:

- **`architecture.md`** — Module dependency graph, key abstractions and their relationships
- **`ooxml-glossary.md`** — ECMA-376 terminology (OPC, WML, content types, relationships)
- **`change-patterns.md`** — "If you modify X, you must also change Y" (the **single most impactful doc** for preventing broken changes)
- **`testing-guide.md`** — How to write tests following project patterns
- **`anti-patterns.md`** — Things that look correct but break subtly
- **`ooxml-wml-reference.md`** — Curated ECMA-376 spec snippets by feature area

The `change-patterns.md` file deserves special attention. Document every multi-step change pattern: "Adding a new WML element requires: (1) model class, (2) serializer, (3) deserializer, (4) element factory registration, (5) index.ts export, (6) tests." This prevents Claude from making partial changes that compile but break at runtime.

### Claude Code plus Cursor is the optimal combo

**Claude Code** handles complex autonomous tasks: multi-file refactoring, test generation, spec-compliant implementation, code review. Its 200K token context window and SWE-bench score of **80.9%** make it the strongest tool for reasoning about large codebases. **Cursor** handles interactive editing with fast tab completion (powered by Supermaven). Together they cost ~$40-60/month.

Key Claude Code practices for a 75K-line codebase:

- **Use `/clear` aggressively** — every new task starts fresh to prevent context pollution
- **Use Plan Mode** (Shift+Tab twice) before implementation for complex OOXML changes
- **Use thinking levels strategically**: "ultrathink" for complex spec interpretation, regular for routine changes
- **Use `--add-dir ../dochub-app`** to work across both repos simultaneously
- **Set up hooks** in `.claude/settings.json` to auto-run Prettier and TypeScript type checking after every edit

### Custom skills replace repetitive prompting

Create `.claude/skills/` for common operations:

- **`/new-element`**: Reads change-patterns.md, relevant spec section, creates model/serializer/deserializer/tests
- **`/add-test`**: Reads existing test patterns, generates tests following project conventions
- **`/review`**: Acts as skeptical senior engineer reviewing all changes in git status
- **`/ooxml-lookup`**: Fetches relevant ECMA-376 spec context for the current task

### Keep MCP servers minimal — two is enough

Community consensus: "If you're using more than 20K tokens of MCPs, you're crippling Claude." Every MCP server consumes context budget. Two servers provide genuine value:

- **GitHub MCP** — PR management, issue tracking directly from Claude Code
- **Context7 MCP** — Fetches current library documentation (solves stale training data for third-party APIs)

An optional high-value investment: **build a custom OOXML spec MCP server** that accepts an element name (e.g., "w:p") and returns the ECMA-376 definition. The TypeScript MCP SDK makes this achievable in ~100 lines. The spec is freely downloadable from ecma-international.org and available in web-friendly format at c-rex.net.

---

## Dependency management between the repos

### Exact pinning plus Renovate automerge

Use **exact version pinning** in dochub-app (`"docxmlater": "10.0.2"`, no `^` or `~`). You control both sides — exact pinning means you know precisely which version is deployed. Combined with Renovate automerge, you get the best of both worlds: each version bump is a discrete, visible event that auto-merges only if CI tests pass.

### Emergency hotfix workflow

When docxmlater breaks dochub-app in production, the playbook has three escalation levels:

**Level 1 — Rollback** (under 5 minutes): Change the version to the last known good release, `npm install`, rebuild, deploy.

**Level 2 — Temporary patch** (if rollback isn't possible): Install `patch-package`, edit the framework directly in `node_modules`, run `npx patch-package docxmlater` to create a persistent patch file. Commit the patch. This buys time for a proper fix.

**Level 3 — Hotfix release**: Branch from the broken tag, fix, `npm version patch`, `npm publish`, then update the consumer. Renovate will also auto-detect the fix.

### Pre-release testing before publishing

Use **npm dist-tags** for staged releases: `npm version prerelease --preid=rc` creates `10.0.3-rc.0`, then `npm publish --tag next` publishes to the `next` tag without affecting `latest`. Install in dochub-app with `npm install docxmlater@next`, run the full test suite, and if everything passes, publish the real release.

---

## Preventing architectural decay as a solo maintainer

### The recommended implementation sequence

These tools have the highest ROI for a solo developer, ordered by impact-per-hour-invested:

1. **Week 1**: Install `@xarsh/ooxml-validator` and add validation to every integration test. Set up ESLint flat config with `recommendedTypeChecked` (all rules as `warn`). Configure Release Please GitHub Action.
2. **Week 2**: Enable `noUncheckedIndexedAccess` in tsconfig (fix resulting errors — each one is a potential bug found). Implement golden file testing infrastructure for the 5-10 most common document types. Set up CI workflows (4 YAML files total across both repos).
3. **Week 3**: Run Knip for dead code cleanup. Set up dependency-cruiser rules. Install Renovate on dochub-app. Configure size-limit.
4. **Week 4**: Rewrite CLAUDE.md to <100 lines. Create `agent_docs/` with architecture, glossary, change-patterns, and testing guide. Set up Claude Code skills and hooks.

This sequence front-loads **bug prevention** (validation, strict types) before **automation** (CI, Renovate) before **AI optimization** (documentation, skills). Each week builds on the previous one, and any week's work stands on its own if you stop there.

---

## Conclusion

The core tension for a solo developer maintaining two linked repos is **reliability versus speed**. yalc resolves this for local development — fast feedback without the symlink disasters that plague electron-builder. For testing, the inverted pyramid with mandatory OOXML validation on every generated document will catch more corruption bugs than doubling the unit test count. The four-file CI/CD setup (two workflows per repo) plus Renovate handles the cross-repo coordination that would otherwise require manual attention on every release.

The AI-assisted development findings carry a counterintuitive lesson: **less documentation is more effective** for AI tools. A focused 100-line CLAUDE.md with pointers to deep-context files outperforms a 1,000-line knowledge dump that gets ignored. The `change-patterns.md` concept — explicitly documenting multi-file change requirements — addresses the developer's core frustration that "AI tools don't understand the project" better than any amount of architectural documentation.

The total tooling cost is modest: ~$40-60/month for Claude Code plus Cursor, zero for all CI/CD (public repo is free, private repo stays well under 2,000 minutes), and zero for Renovate, Codecov, and all the linting/testing tools. The time investment to implement everything is roughly **four focused weeks**, after which the infrastructure maintains itself and the AI tools become genuinely useful collaborators rather than context-blind code generators.
