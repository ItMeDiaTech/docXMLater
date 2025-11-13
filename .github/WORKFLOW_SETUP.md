# GitHub Actions Workflow Setup

This repository includes an automated npm publishing workflow that triggers when you create a new git tag.

## How It Works

1. **Create a new tag** (e.g., `v0.23.0`)
2. **Push the tag** to GitHub
3. **GitHub Actions** automatically:
   - Checks out the code
   - Installs dependencies
   - Runs tests
   - Builds the package
   - Publishes to npm

## Setup Instructions

### 1. Create npm Access Token

1. Go to [npmjs.com](https://www.npmjs.com/) and log in
2. Click your profile icon → **Access Tokens**
3. Click **Generate New Token** → **Classic Token**
4. Select **Automation** type
5. Copy the token (you won't see it again!)

### 2. Add Token to GitHub Secrets

1. Go to your GitHub repository
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret**
4. Name: `NPM_TOKEN`
5. Value: Paste your npm token
6. Click **Add secret**

### 3. Usage

Now you can publish new versions with just two commands:

```bash
# Update version in package.json, commit, and create tag
npm version patch  # or minor, or major
git push && git push --tags

# GitHub Actions will automatically:
# - Run tests
# - Build the package
# - Publish to npm
```

## Manual Process (Alternative)

If you prefer to publish manually (current process):

```bash
# Update version
npm version patch

# Commit and push
git push && git push --tags

# Publish manually
npm publish
```

## Workflow File

The workflow is defined in `.github/workflows/publish.yml`

Key features:
- Triggers on tags matching `v*.*.*` pattern
- Runs on Ubuntu latest
- Requires Node.js 18
- Runs tests before publishing
- Uses npm provenance for enhanced security
- Requires NPM_TOKEN secret

## Security

- The workflow uses `npm publish --provenance` for enhanced security
- Requires `id-token: write` permission for provenance
- Token is stored securely in GitHub Secrets
- Never commit the token to the repository

## Troubleshooting

### Workflow doesn't trigger

- Make sure tag follows `v*.*.*` format (e.g., `v0.22.0`)
- Check that you pushed the tag: `git push --tags`
- Verify workflow file is in `.github/workflows/` directory

### Tests fail

- Workflow will abort if tests fail
- Check the Actions tab on GitHub for error details
- Fix issues locally and create a new tag

### Publishing fails

- Verify NPM_TOKEN is set in GitHub Secrets
- Check token hasn't expired
- Ensure you have publish rights to the package

## Example Release Process

```bash
# 1. Make your changes
git add .
git commit -m "feat: add new feature"

# 2. Update version and create tag
npm version minor  # Creates commit + tag automatically

# 3. Push everything
git push && git push --tags

# 4. GitHub Actions takes over!
# Check: https://github.com/ItMeDiaTech/docXMLater/actions
```

## Benefits of Automation

- **Consistency**: Same build process every time
- **Speed**: No manual steps required
- **Security**: Provenance attestation
- **Reliability**: Tests must pass before publishing
- **Convenience**: One command to release
