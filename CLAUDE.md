# CLAUDE.md

## Keeping the generated API files in sync

`frontend/openapi.json` and `frontend/src/api/schema.d.ts` are generated from the backend and
checked by the Frontend CI job with `git diff --exit-code`.  Regenerate them whenever you change
the API models **or the package version** (the schema records it in `info.version`):

```console
$ pnpm -C frontend gen:openapi
$ pnpm -C frontend gen:api
```

## Releasing

A version bump touches `pyproject.toml`, `frontend/package.json`, `uv.lock` (`uv lock`) and
`frontend/openapi.json` (regenerate as above).  Commit all four in the `chore: Release vX.Y.Z`
commit before tagging it; see the "Releasing" section of README.md.
