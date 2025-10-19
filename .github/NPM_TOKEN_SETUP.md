# GitHub Actions NPM_TOKEN Setup Guide

## The Issue

The GitHub Actions workflow is failing because the `NPM_TOKEN` secret is not configured.

**Error**: The workflow cannot authenticate with npm registry.

## Solution: Add NPM_TOKEN to GitHub Secrets

### Step 1: Get Your npm Access Token

You already have an npm token since you can publish manually. To create a new automation token:

1. Go to [npmjs.com](https://www.npmjs.com/) and log in
2. Click your profile icon → **Access Tokens**
3. Click **Generate New Token** → **Classic Token**
4. Select **Automation** type (not Publish)
5. Name it: `GitHub Actions - docXMLater`
6. Click **Generate Token**
7. **COPY THE TOKEN** (you won't see it again!)

### Step 2: Add Token to GitHub Secrets

1. Go to your GitHub repository: https://github.com/ItMeDiaTech/docXMLater
2. Click **Settings** tab
3. Click **Secrets and variables** → **Actions** (left sidebar)
4. Click **New repository secret**
5. Enter details:
   - **Name**: `NPM_TOKEN` (exactly this - case sensitive!)
   - **Secret**: Paste your npm token from Step 1
6. Click **Add secret**

### Step 3: Verify Setup

After adding the secret, you can test it:

**Option 1: Push a new tag**
```bash
git tag v0.23.2-test
git push origin v0.23.2-test
```

**Option 2: Re-run the failed workflow**
1. Go to: https://github.com/ItMeDiaTech/docXMLater/actions
2. Click on the failed "Publish to npm" workflow
3. Click **Re-run all jobs**

### Step 4: Verify Publication

Once the workflow succeeds, verify the package:

```bash
npm view docxmlater@0.23.1 version
```

## Current Status

**Manual publication completed**: v0.23.1 is live on npm!

The manual publish worked because you're logged in locally (`npm whoami` shows `diatech`).

The GitHub Actions workflow needs the `NPM_TOKEN` secret to authenticate from the CI environment.

## Workflow File Reference

The workflow is located at: `.github/workflows/publish.yml`

It triggers when you push tags matching `v*.*.*` pattern (e.g., `v0.23.1`).

The relevant line that needs the token:
```yaml
- name: Publish to npm
  run: npm publish --provenance --access public
  env:
    NODE_AUTH_TOKEN: ${{ secrets.NPM_TOKEN }}  # ← This needs to be set
```

## Troubleshooting

### Workflow still fails after adding token

Check that:
- Token name is exactly `NPM_TOKEN` (case sensitive)
- Token type is **Automation** (not Publish or Read-only)
- Token hasn't expired
- You have publish rights to the `docxmlater` package

### How to check if secret is set

You can't view the secret value (it's encrypted), but you can see if it exists:

1. Go to: Settings → Secrets and variables → Actions
2. Look for `NPM_TOKEN` in the "Repository secrets" list
3. If it's there, it's set ✅

### Alternative: Manual Publishing

If you prefer not to use automated publishing, you can always publish manually:

```bash
# Update version
npm version patch  # or minor/major

# Publish manually
npm publish

# Push changes
git push && git push --tags
```

## Security Note

**Never commit npm tokens to git!**

The token is stored securely in GitHub Secrets and is only accessible during workflow runs.
