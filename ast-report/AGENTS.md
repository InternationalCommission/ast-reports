# IC Project Report — Agent Guidelines

## Cloudflare Workers

**STOP.** Your knowledge of Cloudflare Workers APIs and limits may be outdated. Always retrieve current documentation before any Workers, KV, R2, D1, Durable Objects, Queues, Vectorize, AI, or Agents SDK task.

### Docs
- https://developers.cloudflare.com/workers/
- MCP: `https://docs.mcp.cloudflare.com/mcp`

For all limits and quotas, retrieve from the product's `/platform/limits/` page. e.g. `/workers/platform/limits`

### Node.js Compatibility
https://developers.cloudflare.com/workers/runtime-apis/nodejs/

### Errors
- **Error 1102** (CPU/Memory exceeded): Retrieve limits from `/workers/platform/limits/`
- **All errors**: https://developers.cloudflare.com/workers/observability/errors/

### Product Docs
Retrieve API references and limits from:
`/kv/` · `/r2/` · `/d1/` · `/durable-objects/` · `/queues/` · `/vectorize/` · `/workers-ai/` · `/agents/`

---

## Commands

| Command | Purpose |
|---------|---------|
| `npm run dev` or `npx wrangler dev` | Local development server |
| `npm run deploy` or `npx wrangler deploy` | Deploy to Cloudflare production |
| `npx wrangler types` | Generate TypeScript types from bindings |

Run `wrangler types` after changing bindings in `wrangler.jsonc`.

### Running a Single Test
There are no automated tests in this project. Test manually using the form at `form.html` and admin at `admin.html`.

---

## Code Style Guidelines

### Formatting (Prettier + EditorConfig)
- **Print width:** 140 characters
- **Quotes:** Single quotes
- **Semicolons:** Always
- **Indentation:** Tabs (not spaces)
- **Line endings:** LF
- **Charset:** UTF-8
- **Trailing whitespace:** Trimmed
- **Final newline:** Yes

Config: `.prettierrc` and `.editorconfig`

### JavaScript Conventions
- Use ES modules (`export default`, `import` not needed since Workers use module syntax)
- Use async/await over Promise chains
- Use `const` by default, `let` only when reassignment needed
- Always use meaningful variable/function names (camelCase)
- Use `PascalCase` for exported default objects (the Worker handler)
- Prefix private module-level variables with underscore: `_jwksCache`

### Error Handling
- Wrap async operations in try/catch blocks
- Return meaningful error messages in responses
- Log errors with `console.error()` including context
- Use `Promise.allSettled()` when multiple parallel operations can partially fail
- Prefer specific error messages over generic ones

### Naming
- Functions: `camelCase`, descriptive verbs (e.g., `handleSubmit`, `getAccessToken`)
- Constants: `UPPER_SNAKE_CASE` for config values (e.g., `JWKS_TTL_MS`)
- Module-level cache/state: `_prefixWithUnderscore` (e.g., `_cachedIds`)
- HTTP methods: uppercase (GET, POST, etc.)

### Imports/Dependencies
- Minimal external dependencies (Wrangler only for dev)
- Use built-in Web APIs: `fetch`, `crypto.subtle`, `URL`, `URLSearchParams`
- No npm packages needed for this Worker

### Type Safety
- This is plain JavaScript (no TypeScript in this project)
- Use JSDoc comments for complex functions if needed
- Validate env variables exist before using them
- Use nullish coalescing (`??`) for fallback values

### Response Format
- Use the `corsResponse()` helper for all responses
- Always return JSON-serializable bodies
- Use appropriate HTTP status codes:
  - `200` - Success
  - `204` - No content (OPTIONS preflight)
  - `400` - Bad request (if applicable)
  - `401` - Unauthorized
  - `403` - Forbidden
  - `404` - Not found
  - `500` - Server error
  - `207` - Multi-status (partial success)

### Security
- Never log secrets or tokens
- Validate `ALLOWED_ORIGIN` for CORS
- Verify Azure AD JWT signature against JWKS
- Store all secrets via `wrangler secret put`, never in code
- Validate audience and issuer on tokens

### SharePoint/Graph API
- Use Microsoft Graph API for list operations
- Use SharePoint REST API for file uploads (triggers indexing)
- Cache site/list IDs in module-level variables when possible
- Handle missing columns gracefully (use `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`)

### Email
- Use Microsoft Graph `/sendMail` endpoint
- Support test mode via `TEST_MODE` env var
- Send to `TEST_EMAIL_RECIPIENT` in test mode
- Always CC coordinator if email provided

---

## Project Structure

```
ast-report/
├── src/
│   └── worker.js          # Main Worker entry point
├── wrangler.jsonc         # Worker configuration
├── package.json           # npm scripts
├── .prettierrc            # Code formatting rules
├── .editorconfig          # Editor settings
├── form.html              # Public submission form (frontend)
├── admin.html             # Admin viewer (frontend)
└── AGENTS.md              # This file
```

---

## Environment Variables (Secrets)

Set via `wrangler secret put`:
- `AZURE_TENANT_ID` - Azure AD tenant ID
- `AZURE_CLIENT_ID` - App registration for SharePoint/Graph
- `AZURE_CLIENT_SECRET` - Client secret fallback (deprecated, use certificate)
- `AZURE_CLIENT_CERTIFICATE` - PEM private key or PFX for SharePoint authentication (recommended)
- `AZURE_CLIENT_CERTIFICATE_PASSWORD` - Password for PFX certificate (optional for PEM)
- `ADMIN_CLIENT_ID` - App registration for admin SPA
- `SHAREPOINT_SITE_URL` - e.g. https://org.sharepoint.com/sites/site
- `SHAREPOINT_LIST_NAME` - Target list name
- `SHAREPOINT_FOLDER_PATH` - Server-relative folder for photos
- `EMAIL_SENDER` - M365 mailbox to send from
- `EMAIL_RECIPIENT` - Where confirmation emails go
- `ALLOWED_ORIGIN` - Website origin for CORS
- `TEST_MODE` - "true" to send emails to TEST_EMAIL_RECIPIENT
- `TEST_EMAIL_RECIPIENT` - Override recipient in test mode
- `EDIT_TOKEN_SECRET` - Secret key for signing edit tokens
- `SUPER_ADMIN_GROUP_ID` - Azure AD group ID for Super Admin role
- `READWRITE_GROUP_ID` - Azure AD group ID for Read/Write role
- `READONLY_GROUP_ID` - Azure AD group ID for Read-only role

---

## Role-Based Access Control

### Roles and Permissions

| Role | View | Edit | Recycle | Restore | Perm. Delete |
|------|------|------|---------|---------|-------------|
| Read-only | ✓ | ✗ | ✗ | ✗ | ✗ |
| Read/Write | ✓ | ✓ | ✓ | ✗ | ✗ |
| Super Admin | ✓ | ✓ | ✓ | ✓ | ✓ |

### Azure AD Group IDs (Current)

| Role | Group ID |
|------|----------|
| Super Admin | `a21edf30-1be1-4a2c-b13b-411e130d41e2` |
| Read/Write | `2a94d6d7-85d8-482b-92a1-7dda91f3daba` |
| Read-only | `ba1c0013-5cbb-4b5a-b625-19ae165953e5` |

### API Endpoints

| Method | Endpoint | Role Required | Description |
|--------|----------|---------------|-------------|
| GET | `/reports` | Any authenticated | List active reports |
| GET | `/reports/recycle-bin` | SuperAdmin | List recycled reports |
| GET | `/reports/:id` | Any authenticated | Get single report |
| PATCH | `/reports/:id` | ReadWrite+ | Update report |
| POST | `/reports/:id/recycle` | ReadWrite+ | Move to recycle bin |
| POST | `/reports/:id/restore` | SuperAdmin | Restore from recycle bin |
| DELETE | `/reports/:id/permanent` | SuperAdmin | Permanently delete |

### Recycle Bin Feature

Reports can be soft-deleted (recycled) instead of permanently deleted:
- Recycled reports are hidden from the main list
- They appear in the Recycle Bin (SuperAdmin only)
- Reports can be restored or permanently deleted
- SharePoint column `IsRecycled` tracks status