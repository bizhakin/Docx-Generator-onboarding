# Workspace

## Overview

pnpm workspace monorepo using TypeScript. Each package manages its own dependencies.
Also contains a standalone Python Flask app for contract document generation.

## Stack

- **Monorepo tool**: pnpm workspaces
- **Node.js version**: 24
- **Package manager**: pnpm
- **TypeScript version**: 5.9
- **API framework**: Express 5
- **Database**: PostgreSQL + Drizzle ORM
- **Validation**: Zod (`zod/v4`), `drizzle-zod`
- **API codegen**: Orval (from OpenAPI spec)
- **Build**: esbuild (CJS bundle)

## Flask Contract Generator App

Located in `/flask-app/`. A standalone Python Flask web app.

### Stack
- **Language**: Python 3.11
- **Framework**: Flask
- **Document generation**: python-docx
- **Templates**: `flask-app/contract_templates/` (.docx files)
- **Config**: `flask-app/contracts.json` (defines contract types and fields)

### Running
- Workflow: `Start application` → `cd flask-app && python3 app.py`
- Runs on port 5000

### Adding new contract types
1. Add a new entry to `flask-app/contracts.json` with `id`, `name`, `template` filename, and `fields` array
2. Create a `.docx` template file in `flask-app/contract_templates/` using `{{PLACEHOLDER}}` syntax
3. The `{{DATE}}` placeholder is always auto-filled with today's date

### Template placeholders
Use `{{FIELD_ID}}` in `.docx` templates. Example: `{{CLIENT_NAME}}`, `{{DATE}}`.

## Key Commands (TypeScript monorepo)

- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- `pnpm --filter @workspace/api-server run dev` — run API server locally

See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details.
