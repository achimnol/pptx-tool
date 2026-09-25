# CLAUDE.md

## General Rules

* Keep the generated API schema files in sync when changing the API or the package version.
  - Files: `frontend/openapi.json`, `frontend/src/api/schema.d.ts`
  - Command: `pnpm -C frontend gen:openapi` and `pnpm -C frontend gen:api`.
* Release process
  - Bump the version numbers in `pyproject.toml`, `frontend/package.json`, `uv.lock`
  - Ensure updates of the API schema files mentioning the package version.
  - Commit with the message: `chore: Release vX.Y.Z`
  - Create an annotated tag on it and push.
