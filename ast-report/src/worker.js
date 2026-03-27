/**
 * IC Project Report — Cloudflare Worker
 *
 * POST /              — Submit a report (public, CORS-restricted)
 * GET  /reports       — Fetch all reports for admin viewer (requires Azure AD JWT)
 * GET  /reports/recycle-bin — Fetch recycled reports (requires SuperAdmin role)
 * GET  /reports/:id   — Fetch single report by SharePoint item ID
 * POST /reports/:id/edit-token — Generate edit token (requires Azure AD JWT)
 * PATCH /reports/:id  — Update report (requires valid edit token + ReadWrite/SuperAdmin role)
 * POST /reports/:id/recycle — Move report to recycle bin (requires ReadWrite/SuperAdmin role)
 * POST /reports/:id/restore — Restore report from recycle bin (requires SuperAdmin role)
 * DELETE /reports/:id/permanent — Permanently delete report (requires SuperAdmin role)
 * GET  /reports/:id/photos — List photos in report's folder
 *
 * Required Environment Variables (wrangler secret put ...):
 *   AZURE_TENANT_ID         - Azure AD tenant ID
 *   AZURE_CLIENT_ID         - App registration client ID (backend + API scope audience)
 *   AZURE_CLIENT_SECRET     - Client secret for Key Vault/ROPC access
 *   AZURE_KEY_VAULT_URL     - Azure Key Vault URL (e.g. https://yourvault.vault.azure.net/)
 *   AZURE_KEY_VAULT_CERT_NAME - Name of certificate in Key Vault
 *   AZURE_CLIENT_CERTIFICATE_PASSWORD - Password for PFX certificate (optional for PEM)
 *   ADMIN_CLIENT_ID         - App registration client ID for the admin SPA
 *                             (can be the same as AZURE_CLIENT_ID if using one app)
 *   SHAREPOINT_SITE_URL     - e.g. https://yourorg.sharepoint.com/sites/yoursite
 *   SHAREPOINT_LIST_NAME    - Target list name, e.g. "IC Project Reports"
 *   SHAREPOINT_FOLDER_PATH  - Server-relative folder for photo uploads
 *   EMAIL_SENDER            - Licensed M365 mailbox to send from
 *   EMAIL_RECIPIENT         - Where confirmation emails go
 *   ALLOWED_ORIGIN          - Your website origin for CORS (form + admin)
 *   EDIT_TOKEN_SECRET       - Secret key for signing edit tokens (wrangler secret put)
 *   SUPER_ADMIN_GROUP_ID    - Azure AD group ID for Super Admin role
 *   READWRITE_GROUP_ID      - Azure AD group ID for Read/Write role
 *   READONLY_GROUP_ID       - Azure AD group ID for Read-only role
 *   SERVICE_ACCOUNT_USERNAME - SharePoint service account email for file uploads (via ROPC)
 *   SERVICE_ACCOUNT_PASSWORD - SharePoint service account password (via ROPC)
 *   POWER_AUTOMATE_WEBHOOK_URL - Power Automate HTTP trigger webhook URL (for file uploads)
 *   UPLOAD_METHOD          - 'powerautomate' or 'sharepoint' (default: powerautomate)
 */

export default {
  async fetch(request, env) {
    const url    = new URL(request.url);
    const path   = url.pathname.replace(/\/$/, "") || "/";
    const method = request.method;

    if (method === "OPTIONS") return corsResponse(null, 204, env);

    if (method === "POST" && path === "/")          return handleSubmit(request, env);
    if (method === "GET"  && path === "/reports")   return handleGetReports(request, env, url);
    if (method === "GET"  && path === "/reports/recycle-bin") return handleGetRecycleBin(request, env);
    if (method === "POST" && path.match(/^\/reports\/[^/]+\/edit-token$/)) {
      const id = path.split("/reports/")[1].replace("/edit-token", "");
      return handleEditToken(request, env, id);
    }
    if (method === "PATCH" && path.match(/^\/reports\/[^/]+$/)) {
      const id = path.split("/reports/")[1];
      return handleUpdateReport(request, env, id);
    }
    if (method === "POST" && path.match(/^\/reports\/[^/]+\/recycle$/)) {
      const id = path.split("/reports/")[1].replace("/recycle", "");
      return handleRecycleReport(request, env, id);
    }
    if (method === "POST" && path.match(/^\/reports\/[^/]+\/restore$/)) {
      const id = path.split("/reports/")[1].replace("/restore", "");
      return handleRestoreReport(request, env, id);
    }
    if (method === "DELETE" && path.match(/^\/reports\/[^/]+\/permanent$/)) {
      const id = path.split("/reports/")[1].replace("/permanent", "");
      return handlePermanentDelete(request, env, id);
    }
    if (method === "GET"  && path.match(/^\/reports\/[^/]+\/photos$/)) {
      const id = path.split("/reports/")[1].replace("/photos", "");
      return handleGetPhotos(request, env, id);
    }
    if (method === "GET"  && path.startsWith("/reports/")) {
      return handleGetReport(request, env, path.split("/reports/")[1]);
    }

    return corsResponse({ error: "Not found" }, 404, env);
  },
};

// ────────────────────────────────────────────────────────────────────────────
// POST / — Submit a new report
// ────────────────────────────────────────────────────────────────────────────

async function handleSubmit(request, env) {
  console.log(">>> HANDLING SUBMIT");
  const origin = request.headers.get("Origin") || "";
  const allowedOrigins = (env.ALLOWED_ORIGIN || "").split(",").map(o => o.trim());
  allowedOrigins.push("http://localhost:8080");
  if (env.ALLOWED_ORIGIN && !allowedOrigins.includes(origin)) {
    return corsResponse({ error: "Forbidden origin" }, 403, env);
  }

  try {
    const formData = await request.formData();
    const fields   = extractFields(formData);
    const photos   = formData.getAll("photos").concat(formData.getAll("photo"));
    console.log(`[handleSubmit] Photos found: ${photos.length}`);

    const graphToken = await getAccessToken(env);

    // Run all three in parallel
    const [listItemResult, uploadResults, emailResult] = await Promise.allSettled([
      createSharePointListItem(fields, env, graphToken),
      uploadPhotos(photos, fields, env),
      sendConfirmationEmail(fields, env, graphToken),
    ]);

    // If both list item and uploads succeeded, patch the folder URL back onto
    // the list item so the admin viewer can find the photos later
    if (listItemResult.status === "fulfilled" && uploadResults.status === "fulfilled") {
      const { folderWebUrl, folderServerRelativePath, driveId, folderItemId } = uploadResults.value;
      const listItemId = listItemResult.value?.id;
      if (folderWebUrl && folderServerRelativePath && listItemId) {
        await patchPhotoFolder(listItemId, folderWebUrl, folderServerRelativePath, env, graphToken, {
          PhotoDriveId: driveId,
          PhotoFolderItemId: folderItemId,
        }).catch(e => console.warn("Could not patch photo folder URL:", e.message));
      }
    }

    const errors = [];
    if (listItemResult.status === "rejected")
      errors.push({ step: "sharepoint_list", message: listItemResult.reason?.message });
    if (uploadResults.status === "rejected")
      errors.push({ step: "file_upload", message: uploadResults.reason?.message });
    if (emailResult.status === "rejected")
      errors.push({ step: "email", message: emailResult.reason?.message });

    // Add per-file upload errors if some uploads failed but not all
    if (uploadResults.status === "fulfilled" && uploadResults.value?.errors?.length > 0) {
      uploadResults.value.errors.forEach(e => {
        errors.push({ step: "file_upload", file: e.file, message: e.error });
      });
    }

    const listItemId  = listItemResult.status  === "fulfilled" ? listItemResult.value?.id   : null;
    const uploadValue = uploadResults.status   === "fulfilled" ? uploadResults.value         : null;

  return corsResponse(
    {
      success:       errors.length === 0,
      message:       errors.length === 0 ? "Report submitted successfully." : "Report submitted with some issues.",
      listItemId:    listItemId ?? null,
      uploadedFiles: uploadValue?.uploaded ?? [],
      uploadErrors:  uploadValue?.errors ?? [],
      errors,
    },
    errors.length === 0 ? 200 : 207,
    env
  );
  } catch (err) {
    console.error("Submit error:", err);
    return corsResponse({ success: false, error: err.message }, 500, env);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// GET /reports — Return paginated list of all report items
// ────────────────────────────────────────────────────────────────────────────

async function handleGetReports(request, env, url) {
  const authError = await validateAzureToken(request, env);
  if (authError) return corsResponse({ error: authError }, 401, env);

  try {
    const token = await getAccessToken(env);
    const { siteId, listId } = await resolveListIds(env, token);

    const top    = url.searchParams.get("top")    || "50";
    const cursor = url.searchParams.get("cursor") || null;

    // Prefer header: lets Graph query non-indexed columns without erroring.
    // Results may occasionally be inconsistent on very large lists, but is
    // fine for typical report volumes.
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
    };

    // Fetch all fields without $select — avoids 400s from missing columns.
    // No $orderby — SubmittedAt is not indexed; we sort client-side below.
    // Filter to exclude recycled items.
    const endpoint = cursor
      ? decodeURIComponent(cursor)
      : `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`
          + `?expand=fields`
          + `&$filter='Is_x0020_Recycled' eq false`
          + `&$top=${top}`;

    let res;
    let graphFilterFailed = false;
    try {
      res = await graphFetch(endpoint, { headers });
    } catch (filterError) {
      // If filter fails (e.g., IsRecycled column not found), fetch all and filter client-side
      console.error("Filter failed, falling back to client-side filtering:", filterError.message);
      graphFilterFailed = true;
      const fallbackEndpoint = cursor
        ? decodeURIComponent(cursor)
        : `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`
            + `?expand=fields`
            + `&$top=${top}`;
      res = await graphFetch(fallbackEndpoint, { headers });
      // Log available fields for debugging
      if (res.value?.length > 0) {
        const firstItemFields = res.value[0].fields;
        const fieldNames = firstItemFields ? Object.keys(firstItemFields) : [];
        console.log("Available fields:", JSON.stringify(fieldNames));
        console.log("First item raw fields:", JSON.stringify(firstItemFields));
      }
    }

    // Sort descending by SubmittedAt client-side since the column isn't indexed
    let items = (res.value || [])
      .map(item => normalizeItem(item));

    // If Graph filter failed, check if IsRecycled exists and filter client-side
    if (graphFilterFailed && items.length > 0) {
      const rawFields = res.value[0]?.fields;
      if (rawFields && ('Is Recycled' in rawFields || 'IsRecycled' in rawFields)) {
        console.log("Is Recycled found in fields - applying client-side filter");
        items = items.filter(item => !item.isRecycled);
      } else {
        console.warn("Is Recycled field NOT found in SharePoint - showing all items");
      }
    } else if (!graphFilterFailed) {
      // Graph filter worked, but still filter in case of inconsistencies
      items = items.filter(item => !item.isRecycled);
    }

    items.sort((a, b) => new Date(b.submittedAt || 0).getTime() - new Date(a.submittedAt || 0).getTime());

    // Convert SharePoint's nextLink into a worker-relative cursor URL so the
    // admin never calls SharePoint directly and all requests stay authenticated.
    const spNextLink = res["@odata.nextLink"] || null;
    const workerNextLink = spNextLink
      ? `${new URL(request.url).origin}/reports?cursor=${encodeURIComponent(spNextLink)}`
      : null;

    return corsResponse({ items, nextLink: workerNextLink }, 200, env);
  } catch (err) {
    console.error("GetReports error:", err);
    return corsResponse({ error: err.message }, 500, env);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// GET /reports/:id — Return a single report item
// ────────────────────────────────────────────────────────────────────────────

async function handleGetReport(request, env, id) {
  const authError = await validateAzureToken(request, env);
  if (authError) return corsResponse({ error: authError }, 401, env);

  try {
    const token = await getAccessToken(env);
    const { siteId, listId } = await resolveListIds(env, token);

    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
    };
    const item = await graphFetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${id}?expand=fields`,
      { headers }
    );

    return corsResponse(normalizeItem(item), 200, env);
  } catch (err) {
    console.error("GetReport error:", err);
    return corsResponse({ error: err.message }, 500, env);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// POST /reports/:id/edit-token — Generate edit token for a report
// ────────────────────────────────────────────────────────────────────────────

async function handleEditToken(request, env, id) {
  const authError = await validateAzureToken(request, env);
  if (authError) return corsResponse({ error: authError }, 401, env);

  if (!env.EDIT_TOKEN_SECRET) {
    return corsResponse({ error: "Edit token secret not configured" }, 500, env);
  }

  try {
    const expiryMs = 30 * 60 * 1000; // 30 minutes
    const expires = Date.now() + expiryMs;
    const dataToSign = `${id}:${expires}`;
    const signature = await signData(dataToSign, env.EDIT_TOKEN_SECRET);
    
    return corsResponse({
      token: signature,
      expires: expires,
      reportId: id
    }, 200, env);
  } catch (err) {
    console.error("EditToken error:", err);
    return corsResponse({ error: err.message }, 500, env);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Token signing and validation utilities
// ────────────────────────────────────────────────────────────────────────────

async function signData(data, secret) {
  const encoder = new TextEncoder();
  const keyData = encoder.encode(secret);
  const messageData = encoder.encode(data);
  const key = await crypto.subtle.importKey(
    "raw", keyData,
    { name: "HMAC", hash: "SHA-256" },
    false, ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, messageData);
  const hashArray = Array.from(new Uint8Array(signature));
  return hashArray.map(b => b.toString(16).padStart(2, "0")).join("");
}

async function validateEditToken(reportId, token, expires, secret) {
  if (!token || !expires || !secret) return false;
  if (Date.now() > expires) return false;
  
  const dataToSign = `${reportId}:${expires}`;
  const expectedToken = await signData(dataToSign, secret);
  
  return token === expectedToken;
}

// ────────────────────────────────────────────────────────────────────────────
// PATCH /reports/:id — Update an existing report
// ────────────────────────────────────────────────────────────────────────────

async function handleUpdateReport(request, env, id) {
  if (!env.EDIT_TOKEN_SECRET) {
    return corsResponse({ error: "Edit token secret not configured" }, 500, env);
  }

  const roleCheck = await requireRole(request, env, ["SuperAdmin", "ReadWrite"]);
  if (roleCheck.error) return corsResponse({ error: roleCheck.error }, 403, env);

  try {
    const formData = await request.formData();
    const token = formData.get("editToken");
    const expires = formData.get("editExpires");
    
    const isValid = await validateEditToken(
      id,
      token,
      expires ? parseInt(expires, 10) : 0,
      env.EDIT_TOKEN_SECRET
    );
    
    if (!isValid) {
      return corsResponse({ error: "Invalid or expired edit token" }, 401, env);
    }

    const fields = extractFields(formData);
    const photos = formData.getAll("photos");

    const graphToken = await getAccessToken(env);

    const updateResult = await updateSharePointListItem(id, fields, env, graphToken);

    let photoResult = { uploaded: [], errors: [] };
    if (photos.length > 0) {
      photoResult = await uploadPhotos(photos, fields, env);
    }

    return corsResponse({
      success: true,
      message: "Report updated successfully",
      listItemId: id,
      uploadedFiles: photoResult.uploaded,
      uploadErrors: photoResult.errors,
    }, 200, env);
  } catch (err) {
    console.error("UpdateReport error:", err);
    return corsResponse({ success: false, error: err.message }, 500, env);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Azure AD JWT validation
// ────────────────────────────────────────────────────────────────────────────

// Module-level JWKS cache (lives for the duration of the Worker isolate)
const _jwksCache = { keys: null, fetchedAt: 0 };
const JWKS_TTL_MS = 60 * 60 * 1000; // 1 hour

async function getJwks(tenantId) {
  const now = Date.now();
  if (_jwksCache.keys && (now - _jwksCache.fetchedAt) < JWKS_TTL_MS) {
    return _jwksCache.keys;
  }
  const res  = await fetch(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`);
  const data = await res.json();
  _jwksCache.keys      = data.keys;
  _jwksCache.fetchedAt = now;
  return data.keys;
}

/**
 * Validate an Azure AD Bearer JWT.
 * Returns null on success, or an error string describing the failure.
 *
 * Checks:
 *  - Token is present and well-formed
 *  - Not expired
 *  - Issuer matches our tenant (v2.0 endpoint)
 *  - Audience matches ADMIN_CLIENT_ID or api://ADMIN_CLIENT_ID
 *  - RSA-SHA256 signature verified against Azure AD's public JWKS
 */
async function validateAzureToken(request, env) {
  const authHeader = request.headers.get("Authorization") || "";
  if (!authHeader.startsWith("Bearer ")) return "Missing Bearer token";

  const token = authHeader.slice(7).trim();
  const parts = token.split(".");
  if (parts.length !== 3) return "Malformed JWT";

  let header, payload;
  try {
    header  = JSON.parse(b64urlDecode(parts[0]));
    payload = JSON.parse(b64urlDecode(parts[1]));
  } catch {
    return "Could not decode JWT";
  }

  // Expiry
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) {
    return "Token expired";
  }

  // Issuer — accept both Azure AD v1.0 (sts.windows.net) and v2.0 (login.microsoftonline.com)
  const validIssuers = [
    `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/v2.0`,
    `https://sts.windows.net/${env.AZURE_TENANT_ID}/`,
  ];
  if (!validIssuers.includes(payload.iss)) {
    return `Invalid issuer: ${payload.iss}`;
  }

  // Audience — accept both bare GUID and api:// URI forms
  const adminClientId = env.ADMIN_CLIENT_ID || env.AZURE_CLIENT_ID;
  const validAudiences = [adminClientId, `api://${adminClientId}`];
  if (!validAudiences.includes(payload.aud)) {
    return `Invalid audience: ${payload.aud}`;
  }

  // Signature
  try {
    const keys    = await getJwks(env.AZURE_TENANT_ID);
    const jwk     = keys.find(k => k.kid === header.kid);
    if (!jwk) return `No matching key for kid=${header.kid}`;

    const cryptoKey = await crypto.subtle.importKey(
      "jwk", jwk,
      { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
      false, ["verify"]
    );

    const signingInput = new TextEncoder().encode(`${parts[0]}.${parts[1]}`);
    const signature    = b64urlToBytes(parts[2]);
    const valid        = await crypto.subtle.verify("RSASSA-PKCS1-v1_5", cryptoKey, signature, signingInput);

    if (!valid) return "Invalid signature";
  } catch (err) {
    console.error("JWT sig verification error:", err);
    return "Signature verification failed";
  }

  console.log("[DEBUG] Token payload:", JSON.stringify(payload, null, 2));
  return null; // ✓ valid
}

function getUserRoles(payload, env) {
  const groups = payload.groups || [];
  const roles = [];
  console.log("[DEBUG] Token groups claim:", groups);
  console.log("[DEBUG] Expected Group IDs - SuperAdmin:", env.SUPER_ADMIN_GROUP_ID, "ReadWrite:", env.READWRITE_GROUP_ID, "ReadOnly:", env.READONLY_GROUP_ID);
  if (groups.includes(env.SUPER_ADMIN_GROUP_ID)) roles.push("SuperAdmin");
  if (groups.includes(env.READWRITE_GROUP_ID)) roles.push("ReadWrite");
  if (groups.includes(env.READONLY_GROUP_ID)) roles.push("ReadOnly");
  console.log("[DEBUG] Resolved roles:", roles);
  return roles;
}

async function requireRole(request, env, requiredRoles) {
  const authHeader = request.headers.get("Authorization") || "";
  if (!authHeader.startsWith("Bearer ")) return { error: "Missing Bearer token", roles: [] };

  const token = authHeader.slice(7).trim();
  const parts = token.split(".");
  if (parts.length !== 3) return { error: "Malformed JWT", roles: [] };

  let header, payload;
  try {
    header  = JSON.parse(b64urlDecode(parts[0]));
    payload = JSON.parse(b64urlDecode(parts[1]));
  } catch {
    return { error: "Could not decode JWT", roles: [] };
  }

  const roles = getUserRoles(payload, env);
  const hasRole = requiredRoles.some(role => roles.includes(role));

  if (!hasRole) {
    return { error: `Access denied. Required roles: ${requiredRoles.join(" or ")}`, roles };
  }

  return { error: null, roles };
}

function b64urlDecode(str) {
  const padded = str.replace(/-/g, "+").replace(/_/g, "/");
  const rem    = padded.length % 4;
  const padded2 = rem ? padded + "=".repeat(4 - rem) : padded;
  return atob(padded2);
}

function b64urlToBytes(str) {
  const bin = b64urlDecode(str);
  return Uint8Array.from(bin, c => c.charCodeAt(0));
}

// ────────────────────────────────────────────────────────────────────────────
// Field extraction from FormData
// ────────────────────────────────────────────────────────────────────────────

function extractFields(formData) {
  const f = (key) => formData.get(key) ?? "";

  let testimonies = [];
  const rawTestimonies = f("testimoniesJson");
  if (rawTestimonies) {
    try {
      const parsed = JSON.parse(rawTestimonies);
      if (Array.isArray(parsed)) {
        testimonies = parsed
          .filter(t => t && (t.author || t.text))
          .map((t, i) => ({
            index:  i + 1,
            author: String(t.author || "").trim(),
            text:   String(t.text   || "").trim(),
          }));
      }
    } catch (e) {
      console.warn("Could not parse testimoniesJson:", e.message);
    }
  }

  return {
    coordinatorName:  f("coordinatorName"),
    coordinatorEmail: f("coordinatorEmail"),
    projectTitle:    f("projectTitle"),
    city:            f("city"),
    country:         f("country"),
    area:            f("area"),
    projectDateFrom: f("projectDateFrom"),
    projectDateTo:   f("projectDateTo"),
    introduction:    f("introduction"),
    churchesParticipated:       f("churchesParticipated"),
    nationalParticipants:       f("nationalParticipants"),
    usaParticipants:             f("usaParticipants"),
    otherCountriesParticipants: f("otherCountriesParticipants"),
    peopleHeardGospel:          f("peopleHeardGospel"),
    professionsOfFaith:         f("professionsOfFaith"),
    rededications:              f("rededications"),
    baptisms:                   f("baptisms"),
    newChurchesPlanted:         f("newChurchesPlanted"),
    testimonies,
    testimoniesJson: JSON.stringify(testimonies),
    totalFundsSent:             f("totalFundsSent"),
    spentOnMaterials:           f("spentOnMaterials"),
    ticketsCost:                f("ticketsCost"),
    fuelCost:                   f("fuelCost"),
    accommodationCost:          f("accommodationCost"),
    foodCost:                   f("foodCost"),
    financialHelpParticipants:  f("financialHelpParticipants"),
    numParticipantsHelp:        f("numParticipantsHelp"),
    ralliesExpenses:            f("ralliesExpenses"),
    ralliesDescription:        f("ralliesDescription"),
    additionalExpenses:          f("additionalExpenses"),
    additionalNeedDescription:   f("additionalNeedDescription"),
    submittedAt: new Date().toISOString(),
  };
}

// ────────────────────────────────────────────────────────────────────────────
// SharePoint helpers
// ────────────────────────────────────────────────────────────────────────────

const sharePointFields = [
  "Title","City","Country","Area","ProjectDateFrom","ProjectDateTo","Introduction",
  "ChurchesParticipated","NationalParticipants","USAParticipants",
  "OtherCountriesParticipants","PeopleHeardGospel",
  "ProfessionsOfFaith","Rededications","Baptisms","NewChurchesPlanted",
  "Testimonies",
  "TotalFundsSent","SpentOnMaterials","TicketsCost","FuelCost",
  "AccommodationCost","FoodCost","FinancialHelpParticipants","NumParticipantsHelp",
  "RalliesExpenses","RalliesDescription","AdditionalExpenses","AdditionalNeedDescription",
  "CoordinatorName","CoordinatorEmail","SubmittedAt","PhotoFolderServerRelativePath",
];

let _cachedIds = null;
async function resolveListIds(env, token) {
  if (_cachedIds) return _cachedIds;
  const headers  = { Authorization: `Bearer ${token}`, Accept: "application/json" };
  const hostname = new URL(env.SHAREPOINT_SITE_URL).hostname;
  const sitePath = new URL(env.SHAREPOINT_SITE_URL).pathname;

  const siteRes  = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`, { headers });
  const listsRes = await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteRes.id}/lists?$filter=displayName eq '${encodeURIComponent(env.SHAREPOINT_LIST_NAME)}'&$select=id`,
    { headers }
  );
  if (!listsRes.value?.length) throw new Error(`List "${env.SHAREPOINT_LIST_NAME}" not found.`);
  _cachedIds = { siteId: siteRes.id, listId: listsRes.value[0].id };
  return _cachedIds;
}

function normalizeItem(item) {
  const f = item.fields || {};
  let testimonies = [];
  if (f.Testimonies) {
    try {
      const parsed = JSON.parse(f.Testimonies);
      if (Array.isArray(parsed)) testimonies = parsed;
    } catch {
      testimonies = [{ index: 1, author: "", text: f.Testimonies }];
    }
  }
  return {
    id:           item.id,
    createdAt:    item.createdDateTime,
    projectTitle:               f.Title,
    city:                      f.City,
    country:                   f.Country,
    area:                      f.Area,
    projectDateFrom:            f.ProjectDateFrom,
    projectDateTo:              f.ProjectDateTo,
    introduction:               f.Introduction,
    churchesParticipated:       f.ChurchesParticipated,
    nationalParticipants:       f.NationalParticipants,
    usaParticipants:            f.USAParticipants,
    otherCountriesParticipants: f.OtherCountriesParticipants,
    peopleHeardGospel:          f.PeopleHeardGospel,
    professionsOfFaith:         f.ProfessionsOfFaith,
    rededications:              f.Rededications,
    baptisms:                   f.Baptisms,
    newChurchesPlanted:         f.NewChurchesPlanted,
    testimonies,
    totalFundsSent:             f.TotalFundsSent,
    spentOnMaterials:           f.SpentOnMaterials,
    ticketsCost:                f.TicketsCost,
    fuelCost:                   f.FuelCost,
    accommodationCost:          f.AccommodationCost,
    foodCost:                   f.FoodCost,
    financialHelpParticipants:  f.FinancialHelpParticipants,
    numParticipantsHelp:        f.NumParticipantsHelp,
    ralliesExpenses:            f.RalliesExpenses,
    ralliesDescription:         f.RalliesDescription,
    additionalExpenses:         f.AdditionalExpenses,
    additionalNeedDescription:  f.AdditionalNeedDescription,
    coordinatorName:            f.CoordinatorName,
    coordinatorEmail:           f.CoordinatorEmail,
    submittedAt:                f.SubmittedAt,
    // Photo folder info — populated after upload
    photoFolderUrl:                  f.PhotoFolderUrl || null,
    photoFolderServerRelativePath:   f.PhotoFolderServerRelativePath || null,
    photoDriveId:                    f.PhotoDriveId || null,
    photoFolderItemId:               f.PhotoFolderItemId || null,
    // Recycle bin status (column is "Is Recycled" with space, internal name may vary)
    isRecycled:                 f['Is Recycled'] ?? f.IsRecycled ?? false,
  };
}

// ── Patch photo folder info back onto the list item after upload ──────────────
async function patchPhotoFolder(itemId, folderWebUrl, folderServerRelativePath, env, token, extraFields = {}) {
	const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Accept: "application/json" };
	const { siteId, listId } = await resolveListIds(env, token);
	await graphFetch(
		`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
		{
			method: "PATCH",
			headers,
			body: JSON.stringify({
				PhotoFolderUrl: folderWebUrl,
				PhotoFolderServerRelativePath: folderServerRelativePath,
				...extraFields,
			}),
		}
	);
}

// ── GET /reports/:id/photos — list photos in a report's folder ────────────────
async function handleGetPhotos(request, env, id) {
	const authError = await validateAzureToken(request, env);
	if (authError) return corsResponse({ error: authError }, 401, env);

	try {
		const graphToken = await getAccessToken(env);
		const { siteId, listId } = await resolveListIds(env, graphToken);
		const headers = { Authorization: `Bearer ${graphToken}`, Accept: "application/json" };

		const item = await graphFetch(
			`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${id}?expand=fields($select=PhotoFolderUrl,PhotoFolderServerRelativePath)`,
			{ headers }
		);

		const folderUrl = item.fields?.PhotoFolderUrl || null;
		const folderServerRelativePath = item.fields?.PhotoFolderServerRelativePath || null;

		if (!folderServerRelativePath) {
			return corsResponse({ photos: [], folderUrl }, 200, env);
		}

		const sharePointToken = await getUserToken(env, 'sharepoint');
		const spoHeaders = { Authorization: `Bearer ${sharePointToken}`, Accept: 'application/json;odata=verbose' };
		const spoSiteUrl = env.SHAREPOINT_SITE_URL;
		const encodedPath = folderServerRelativePath.split('/').map(p => encodeURIComponent(p)).join('/');
		
		const filesRes = await fetch(
			`${spoSiteUrl}/_api/web/getfolderbyserverrelativeurl('${encodedPath}')/files`,
			{ headers: spoHeaders }
		);
		
		let files = [];
		if (filesRes.ok) {
			const filesData = await filesRes.json();
			files = filesData.d?.results || filesData.value || [];
		}
		
		const photos = files.map(file => {
			const origin = new URL(env.SHAREPOINT_SITE_URL).origin;
			const serverRelativeUrl = file.ServerRelativeUrl || file.name;
			const webUrl = serverRelativeUrl ? `${origin}${serverRelativeUrl}` : null;
			
			return {
				id: file.UniqueId || file.Name,
				name: file.Name,
				webUrl,
				thumbnail: webUrl,
				lastModified: file.TimeLastModified || null,
				size: file.Length || null,
			};
		});

		return corsResponse({ photos, folderUrl }, 200, env);
	} catch (err) {
		console.error("GetPhotos error:", err);
		return corsResponse({ error: err.message }, 500, env);
	}
}

// ────────────────────────────────────────────────────────────────────────────
// GET /reports/recycle-bin — Return all recycled reports (SuperAdmin only)
// ────────────────────────────────────────────────────────────────────────────

async function handleGetRecycleBin(request, env) {
	const authError = await validateAzureToken(request, env);
	if (authError) return corsResponse({ error: authError }, 401, env);

	const roleCheck = await requireRole(request, env, ["SuperAdmin"]);
	if (roleCheck.error) return corsResponse({ error: roleCheck.error }, 403, env);

	try {
		const token = await getAccessToken(env);
		const { siteId, listId } = await resolveListIds(env, token);

		const headers = {
			Authorization: `Bearer ${token}`,
			Accept: "application/json",
			Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
		};

		const endpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`
			+ `?expand=fields`
			+ `&$filter='Is_x0020_Recycled' eq true`;

		let res;
		let graphFilterFailed = false;
		try {
			res = await graphFetch(endpoint, { headers });
		} catch (filterError) {
			console.error("Recycle bin filter failed, falling back to client-side filtering:", filterError.message);
			graphFilterFailed = true;
			const fallbackEndpoint = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`
				+ `?expand=fields`
				+ `&$top=5000`;
			res = await graphFetch(fallbackEndpoint, { headers });
			if (res.value?.length > 0) {
				const firstItemFields = res.value[0].fields;
				const fieldNames = firstItemFields ? Object.keys(firstItemFields) : [];
				console.log("Available fields:", JSON.stringify(fieldNames));
			}
		}

		let items = (res.value || [])
			.map(item => normalizeItem(item));

		if (graphFilterFailed && items.length > 0) {
			const rawFields = res.value[0]?.fields;
			if (rawFields && ('Is Recycled' in rawFields || 'IsRecycled' in rawFields)) {
				console.log("Is Recycled found - applying client-side filter for recycle bin");
				items = items.filter(item => item.isRecycled);
			} else {
				console.warn("Is Recycled field NOT found - showing all items in recycle bin");
			}
		} else if (!graphFilterFailed) {
			items = items.filter(item => item.isRecycled);
		}

		items.sort((a, b) => new Date(b.submittedAt || 0).getTime() - new Date(a.submittedAt || 0).getTime());

		return corsResponse({ items, roles: roleCheck.roles }, 200, env);
	} catch (err) {
		console.error("GetRecycleBin error:", err);
		return corsResponse({ error: err.message }, 500, env);
	}
}

// ────────────────────────────────────────────────────────────────────────────
// POST /reports/:id/recycle — Move report to recycle bin (ReadWrite+)
// ────────────────────────────────────────────────────────────────────────────

async function handleRecycleReport(request, env, id) {
	const authError = await validateAzureToken(request, env);
	if (authError) return corsResponse({ error: authError }, 401, env);

	const roleCheck = await requireRole(request, env, ["SuperAdmin", "ReadWrite"]);
	if (roleCheck.error) return corsResponse({ error: roleCheck.error }, 403, env);

	try {
		const token = await getAccessToken(env);
		const { siteId, listId } = await resolveListIds(env, token);

		const headers = {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
			Accept: "application/json",
		};

		await graphFetch(
			`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${id}/fields`,
			{
				method: "PATCH",
				headers,
				body: JSON.stringify({ 'Is Recycled': true }),
			}
		);

		return corsResponse({ success: true, message: "Report moved to recycle bin" }, 200, env);
	} catch (err) {
		console.error("RecycleReport error:", err);
		return corsResponse({ error: err.message }, 500, env);
	}
}

// ────────────────────────────────────────────────────────────────────────────
// POST /reports/:id/restore — Restore report from recycle bin (SuperAdmin only)
// ────────────────────────────────────────────────────────────────────────────

async function handleRestoreReport(request, env, id) {
	const authError = await validateAzureToken(request, env);
	if (authError) return corsResponse({ error: authError }, 401, env);

	const roleCheck = await requireRole(request, env, ["SuperAdmin"]);
	if (roleCheck.error) return corsResponse({ error: roleCheck.error }, 403, env);

	try {
		const token = await getAccessToken(env);
		const { siteId, listId } = await resolveListIds(env, token);

		const headers = {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
			Accept: "application/json",
		};

		await graphFetch(
			`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${id}/fields`,
			{
				method: "PATCH",
				headers,
				body: JSON.stringify({ 'Is Recycled': false }),
			}
		);

		return corsResponse({ success: true, message: "Report restored from recycle bin" }, 200, env);
	} catch (err) {
		console.error("RestoreReport error:", err);
		return corsResponse({ error: err.message }, 500, env);
	}
}

// ────────────────────────────────────────────────────────────────────────────
// DELETE /reports/:id/permanent — Permanently delete report (SuperAdmin only)
// ────────────────────────────────────────────────────────────────────────────

async function handlePermanentDelete(request, env, id) {
	const authError = await validateAzureToken(request, env);
	if (authError) return corsResponse({ error: authError }, 401, env);

	const roleCheck = await requireRole(request, env, ["SuperAdmin"]);
	if (roleCheck.error) return corsResponse({ error: roleCheck.error }, 403, env);

	try {
		const token = await getAccessToken(env);
		const { siteId, listId } = await resolveListIds(env, token);

		const headers = {
			Authorization: `Bearer ${token}`,
			Accept: "application/json",
		};

		await graphFetch(
			`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${id}`,
			{
				method: "DELETE",
				headers,
			}
		);

		return corsResponse({ success: true, message: "Report permanently deleted" }, 200, env);
	} catch (err) {
		console.error("PermanentDelete error:", err);
		return corsResponse({ error: err.message }, 500, env);
	}
}

async function createSharePointListItem(fields, env, token) {
	const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Accept: "application/json" };
	const { siteId, listId } = await resolveListIds(env, token);

	const createRes = await graphFetch(
		`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
		{
			method: "POST",
			headers,
			body: JSON.stringify({
				fields: {
					Title:                       fields.projectTitle || "IC Project Report",
					City:                        fields.city,
					Country:                     fields.country,
					Area:                        fields.area,
					ProjectDateFrom:             fields.projectDateFrom || null,
					ProjectDateTo:               fields.projectDateTo   || null,
					Introduction:                fields.introduction,
					ChurchesParticipated:        toNum(fields.churchesParticipated),
					NationalParticipants:        toNum(fields.nationalParticipants),
					USAParticipants:             toNum(fields.usaParticipants),
					OtherCountriesParticipants:  toNum(fields.otherCountriesParticipants),
					PeopleHeardGospel:           toNum(fields.peopleHeardGospel),
					ProfessionsOfFaith:          toNum(fields.professionsOfFaith),
					Rededications:               toNum(fields.rededications),
					Baptisms:                    toNum(fields.baptisms),
					NewChurchesPlanted:          toNum(fields.newChurchesPlanted),
					Testimonies:                 fields.testimoniesJson,
					TotalFundsSent:              toNum(fields.totalFundsSent),
					SpentOnMaterials:            toNum(fields.spentOnMaterials),
					TicketsCost:                 toNum(fields.ticketsCost),
					FuelCost:                    toNum(fields.fuelCost),
					AccommodationCost:           toNum(fields.accommodationCost),
					FoodCost:                    toNum(fields.foodCost),
					FinancialHelpParticipants:   toNum(fields.financialHelpParticipants),
					NumParticipantsHelp:         toNum(fields.numParticipantsHelp),
					RalliesExpenses:             toNum(fields.ralliesExpenses),
					RalliesDescription:          fields.ralliesDescription,
					AdditionalExpenses:          toNum(fields.additionalExpenses),
					AdditionalNeedDescription:   fields.additionalNeedDescription,
					CoordinatorName:             fields.coordinatorName,
					CoordinatorEmail:            fields.coordinatorEmail,
					SubmittedAt:                 fields.submittedAt,
				},
			}),
		}
	);
	return { id: createRes.id };
}

async function updateSharePointListItem(itemId, fields, env, token) {
	const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Accept: "application/json" };
	const { siteId, listId } = await resolveListIds(env, token);

	const updateRes = await graphFetch(
		`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
		{
			method: "PATCH",
			headers,
			body: JSON.stringify({
				Title:                       fields.projectTitle || "IC Project Report",
				City:                        fields.city,
				Country:                     fields.country,
				Area:                        fields.area,
				ProjectDateFrom:             fields.projectDateFrom || null,
				ProjectDateTo:               fields.projectDateTo   || null,
				Introduction:                fields.introduction,
				ChurchesParticipated:        toNum(fields.churchesParticipated),
				NationalParticipants:        toNum(fields.nationalParticipants),
				USAParticipants:             toNum(fields.usaParticipants),
				OtherCountriesParticipants:  toNum(fields.otherCountriesParticipants),
				PeopleHeardGospel:           toNum(fields.peopleHeardGospel),
				ProfessionsOfFaith:          toNum(fields.professionsOfFaith),
				Rededications:               toNum(fields.rededications),
				Baptisms:                    toNum(fields.baptisms),
				NewChurchesPlanted:          toNum(fields.newChurchesPlanted),
				Testimonies:                 fields.testimoniesJson,
				TotalFundsSent:              toNum(fields.totalFundsSent),
				SpentOnMaterials:            toNum(fields.spentOnMaterials),
				TicketsCost:                 toNum(fields.ticketsCost),
				FuelCost:                    toNum(fields.fuelCost),
				AccommodationCost:           toNum(fields.accommodationCost),
				FoodCost:                    toNum(fields.foodCost),
				FinancialHelpParticipants:   toNum(fields.financialHelpParticipants),
				NumParticipantsHelp:         toNum(fields.numParticipantsHelp),
				RalliesExpenses:             toNum(fields.ralliesExpenses),
				RalliesDescription:          fields.ralliesDescription,
				AdditionalExpenses:          toNum(fields.additionalExpenses),
				AdditionalNeedDescription:   fields.additionalNeedDescription,
				CoordinatorName:             fields.coordinatorName,
				CoordinatorEmail:            fields.coordinatorEmail,
			}),
		}
	);
	return { id: itemId };
}

// ────────────────────────────────────────────────────────────────────────────
// Photo upload (to SharePoint via Power Automate or direct REST API)
// ────────────────────────────────────────────────────────────────────────────

async function uploadPhotos(photoFiles, fields, env) {
	console.log(`[uploadPhotos] Received ${photoFiles.length} file(s) in photoFiles`);
	const validFiles = photoFiles.filter(file => file?.size > 0 && file?.name);
	console.log(`[uploadPhotos] Valid files after filter: ${validFiles.length}`);
	if (!validFiles.length) return { uploaded: [], errors: [], folderWebUrl: null, folderServerRelativePath: null };

	const usePowerAutomate = env.UPLOAD_METHOD !== 'sharepoint' && env.POWER_AUTOMATE_WEBHOOK_URL;
	console.log(`[uploadPhotos] Using Power Automate: ${usePowerAutomate}`);
	console.log(`[uploadPhotos] Starting upload of ${validFiles.length} file(s)`);

	const sitePath = new URL(env.SHAREPOINT_SITE_URL).pathname;
	const parentFolderPath = env.SHAREPOINT_FOLDER_PATH || '';
	
	let baseFolderPath;
	if (parentFolderPath.startsWith(sitePath)) {
		baseFolderPath = parentFolderPath;
	} else {
		baseFolderPath = normalizePath(`${sitePath}/${parentFolderPath}`);
	}
	
	console.log(`[uploadPhotos] Base folder path: ${baseFolderPath}`);

	const folderName = buildFolderName(fields.projectTitle, fields.submittedAt);
	console.log(`[uploadPhotos] Folder name: ${folderName}`);

	const folderServerRelativePath = normalizePath(`${baseFolderPath}/${folderName}`);
	
	console.log(`[uploadPhotos] Full folder path: ${folderServerRelativePath}`);

	const uploaded = [];
	const errors = [];

	for (let i = 0; i < validFiles.length; i++) {
		const file = validFiles[i];
		const fileName = sanitizeFileName(file.name);
		console.log(`[uploadPhotos] Uploading file ${i + 1}/${validFiles.length}: ${fileName} (${file.size} bytes)`);

		try {
			const arrayBuffer = await file.arrayBuffer();
			const buffer = new Uint8Array(arrayBuffer);

			let fileInfo;
			if (usePowerAutomate) {
				fileInfo = await uploadPhotoPowerAutomate({
					webhookUrl: env.POWER_AUTOMATE_WEBHOOK_URL,
					siteUrl: env.SHAREPOINT_SITE_URL,
					folderPath: folderServerRelativePath,
					fileName: fileName,
					contentType: file.type || 'application/octet-stream',
					buffer,
					projectTitle: folderName,
				});
			} else {
				const sharePointToken = await getUserToken(env, 'sharepoint');
				fileInfo = await uploadPhotoSPO({
					siteUrl: env.SHAREPOINT_SITE_URL,
					sharePointToken,
					folderServerRelativePath,
					fileName: fileName,
					contentType: file.type || 'application/octet-stream',
					buffer,
				});
			}

			console.log(`[uploadPhotos] File uploaded successfully: ${fileName}`);
			uploaded.push({
				name: fileName,
				webUrl: fileInfo.webUrl,
			});
		} catch (error) {
			console.error(`[uploadPhotos] Upload failed for ${fileName}:`, error.message);
			errors.push({ file: fileName, error: error.message });
		}
	}

	console.log(`[uploadPhotos] Complete: ${uploaded.length} uploaded, ${errors.length} errors`);

	const folderWebUrl = `${new URL(env.SHAREPOINT_SITE_URL).origin}${folderServerRelativePath}`;

	return {
		uploaded,
		errors,
		folderWebUrl,
		folderServerRelativePath,
	};
}

async function uploadPhotoPowerAutomate({ webhookUrl, siteUrl, folderPath, fileName, contentType, buffer, projectTitle }) {
	console.log(`[uploadPhotoPowerAutomate] ${fileName}: ${buffer.length} bytes via Power Automate`);
	
	// Convert buffer to base64 without spreading (avoid stack overflow for large files)
	let binary = '';
	for (let i = 0; i < buffer.length; i++) {
		binary += String.fromCharCode(buffer[i]);
	}
	const base64Content = btoa(binary);
	
	// Strip site path from folderPath - SharePoint only needs relative path
	const sitePath = new URL(siteUrl).pathname;
	let relativeFolderPath = folderPath;
	if (folderPath.startsWith(sitePath)) {
		relativeFolderPath = folderPath.substring(sitePath.length);
	}
	
	console.log(`[uploadPhotoPowerAutomate] Original folderPath: ${folderPath}`);
	console.log(`[uploadPhotoPowerAutomate] Relative folderPath: ${relativeFolderPath}`);
	
	const payload = {
		fileName: fileName,
		content: base64Content,
		contentType: contentType,
		folderPath: relativeFolderPath,
		siteUrl: siteUrl,
		projectTitle: projectTitle,
	};
	
	console.log(`[uploadPhotoPowerAutomate] Sending to webhook: ${webhookUrl.substring(0, 50)}...`);
	
	const response = await fetch(webhookUrl, {
		method: 'POST',
		headers: {
			'Content-Type': 'application/json',
		},
		body: JSON.stringify(payload),
	});
	
	console.log(`[uploadPhotoPowerAutomate] Response status: ${response.status}`);
	
	if (!response.ok) {
		const error = await response.text();
		console.error('[uploadPhotoPowerAutomate] Power Automate upload failed:', error);
		throw new Error(`Power Automate upload failed (${response.status}): ${error}`);
	}
	
	const result = await response.json().catch(() => ({}));
	console.log(`[uploadPhotoPowerAutomate] Upload successful:`, result);
	
	return {
		webUrl: result.webUrl || `${new URL(siteUrl).origin}${folderPath}/${fileName}`,
	};
}

async function ensureSharePointFolder({ siteUrl, sharePointToken, folderPath }) {
	console.log(`[ensureSharePointFolder] Checking if folder exists: ${folderPath}`);
	
	// Strip site path prefix since SharePoint REST API is already at that site
	const sitePath = new URL(siteUrl).pathname;
	let relativePath = folderPath;
	if (relativePath.startsWith(sitePath)) {
		relativePath = relativePath.slice(sitePath.length);
		if (!relativePath.startsWith('/')) relativePath = '/' + relativePath;
	}
	
	// Encode spaces but preserve slashes
	const encodedPath = relativePath.split('/').map(p => encodeURIComponent(p)).join('/');
	console.log(`[ensureSharePointFolder] Original path: ${folderPath}`);
	console.log(`[ensureSharePointFolder] Relative path: ${relativePath}`);
	console.log(`[ensureSharePointFolder] Encoded path: ${encodedPath}`);
	
	const checkResponse = await fetch(
		`${siteUrl}/_api/web/getfolderbyserverrelativeurl('${encodedPath}')`,
		{
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				Accept: 'application/json;odata=verbose',
			},
		}
	);

	if (checkResponse.ok) {
		console.log(`[ensureSharePointFolder] Folder exists: ${relativePath}`);
		return folderPath;
	}

	console.log(`[ensureSharePointFolder] Creating folder: ${relativePath}`);
	const requestDigest = await getSharePointRequestDigest(siteUrl, sharePointToken);

	// Use relative path for folder creation
	const createUrl = `${siteUrl}/_api/web/folders/addUsingPath(decodedUrl='${relativePath}')`;
	console.log(`[ensureSharePointFolder] Using URL: ${createUrl}`);

	const createResponse = await fetch(createUrl, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
			Accept: 'application/json;odata=verbose',
			'X-RequestDigest': requestDigest,
		},
	});

	if (!createResponse.ok && createResponse.status !== 409) {
		const errorText = await createResponse.text();
		console.error('[ensureSharePointFolder] Create failed:', errorText);
		throw new Error(`Folder create failed (${createResponse.status}): ${errorText}`);
	}

	console.log(`[ensureSharePointFolder] Folder created/exists: ${folderPath}`);
	return folderPath;
}

async function getSharePointRequestDigest(siteUrl, token) {
	console.log('[getSharePointRequestDigest] Requesting context info...');
	const response = await fetch(`${siteUrl}/_api/contextinfo`, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${token}`,
			Accept: 'application/json;odata=verbose',
		},
	});

	if (!response.ok) {
		const errorText = await response.text();
		console.error('[getSharePointRequestDigest] Failed:', errorText);
		throw new Error(`Context info failed (${response.status}): ${errorText}`);
	}

	const payload = await response.json();
	const digest = payload?.d?.GetContextWebInformation?.FormDigestValue;
	if (!digest) throw new Error('SharePoint request digest missing');
	return digest;
}

async function uploadPhotoSPO({ siteUrl, sharePointToken, folderServerRelativePath, fileName, contentType, buffer }) {
	console.log(`[uploadPhotoSPO] ${fileName}: ${buffer.length} bytes via SharePoint REST API`);
	console.log(`[uploadPhotoSPO] folderServerRelativePath: ${folderServerRelativePath}`);
	console.log(`[uploadPhotoSPO] siteUrl: ${siteUrl}`);
	console.log(`[uploadPhotoSPO] contentType: ${contentType}`);

	const fullFolderPath = folderServerRelativePath.endsWith('/') 
		? folderServerRelativePath.slice(0, -1)
		: folderServerRelativePath;
	
	const requestDigest = await getSharePointRequestDigest(siteUrl, sharePointToken);
	console.log(`[uploadPhotoSPO] Got digest: ${requestDigest.substring(0, 50)}...`);
	
	// Strip site path prefix since SharePoint REST API is already at that site
	const sitePath = new URL(siteUrl).pathname;
	let relativePath = fullFolderPath;
	if (relativePath.startsWith(sitePath)) {
		relativePath = relativePath.slice(sitePath.length);
		if (!relativePath.startsWith('/')) relativePath = '/' + relativePath;
	}
	
	// Encode spaces but preserve slashes
	const encodedFolderPath = relativePath.split('/').map(p => encodeURIComponent(p)).join('/');
	console.log(`[uploadPhotoSPO] Original path: ${fullFolderPath}`);
	console.log(`[uploadPhotoSPO] Relative path: ${relativePath}`);
	console.log(`[uploadPhotoSPO] encodedFolderPath: ${encodedFolderPath}`);
	
	const fileUrl = `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${encodedFolderPath}')/files/addUsingPath(decodedurl='${fileName}',overwrite=true)`;
	console.log(`[uploadPhotoSPO] Upload URL: ${fileUrl}`);
	
	console.log(`[uploadPhotoSPO] Token preview: ${sharePointToken.substring(0, 50)}...`);

	const response = await fetch(fileUrl, {
		method: 'POST',
		body: buffer,
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
			'Content-Type': contentType,
			'Content-Length': buffer.length,
			'X-RequestDigest': requestDigest,
		},
	});

	console.log(`[uploadPhotoSPO] Response status: ${response.status}`);
	console.log(`[uploadPhotoSPO] Response headers:`, Object.fromEntries(response.headers.entries()));

	if (!response.ok && response.status !== 200 && response.status !== 201) {
		const error = await response.text();
		console.error('[uploadPhotoSPO] SharePoint REST API upload failed:', error);
		throw new Error(`SharePoint REST API upload failed (${response.status}): ${error}`);
	}

	const resultText = await response.text();
	const origin = new URL(siteUrl).origin;
	let webUrl = `${origin}${fullFolderPath}/${fileName}`;
	
	try {
		const parser = new DOMParser();
		const xmlDoc = parser.parseFromString(resultText, "text/xml");
		const uriNode = xmlDoc.querySelector("id");
		if (uriNode) {
			const uri = uriNode.textContent;
			const nameMatch = uri.match(/decodedurl='([^']+)'/);
			if (nameMatch) {
				webUrl = `${siteUrl}${fullFolderPath}/${nameMatch[1]}`;
			}
		}
	} catch (e) {
		console.log('[uploadPhotoSPO] Could not parse response, using default webUrl');
	}
	
	console.log(`[uploadPhotoSPO] Upload successful: ${webUrl}`);

	return {
		ServerRelativeUrl: `${fullFolderPath}/${fileName}`,
		webUrl: webUrl,
	};
}

async function getSharePointToken(env) {
	const scope = 'https://graph.microsoft.com/.default';

	if (env.AZURE_KEY_VAULT_URL && env.AZURE_KEY_VAULT_CERT_NAME) {
		console.log('[getSharePointToken] Using Azure Key Vault certificate');
		return getSharePointTokenWithKeyVault(env, scope);
	}

	if (env.AZURE_CLIENT_CERTIFICATE) {
		console.log('[getSharePointToken] Using inline certificate authentication');
		return getSharePointTokenWithCert(env, scope);
	}

	if (env.AZURE_CLIENT_SECRET) {
		console.log('[getSharePointToken] Using client secret authentication');
		return getSharePointTokenWithSecret(env, scope);
	}

	throw new Error('AZURE_KEY_VAULT_URL, AZURE_CLIENT_CERTIFICATE, or AZURE_CLIENT_SECRET is required');
}

async function getUserToken(env, scope) {
	const scopeLabel = scope === 'graph' ? 'Graph' : 'SharePoint';
	console.log(`[getUserToken] Getting ${scopeLabel} user token via ROPC flow`);
	
	if (!env.SERVICE_ACCOUNT_USERNAME || !env.SERVICE_ACCOUNT_PASSWORD) {
		throw new Error('SERVICE_ACCOUNT_USERNAME and SERVICE_ACCOUNT_PASSWORD are required for user authentication');
	}

	let tokenScope;
	if (scope === 'graph') {
		tokenScope = 'https://graph.microsoft.com/.default';
	} else {
		const hostname = new URL(env.SHAREPOINT_SITE_URL).hostname;
		tokenScope = `https://${hostname}/.default`;
	}
	
	console.log(`[getUserToken] Scope: ${tokenScope}`);

	const body = new URLSearchParams({
		client_id: env.AZURE_CLIENT_ID,
		client_secret: env.AZURE_CLIENT_SECRET || '',
		scope: `${tokenScope} offline_access`,
		username: env.SERVICE_ACCOUNT_USERNAME,
		password: env.SERVICE_ACCOUNT_PASSWORD,
		grant_type: 'password'
	});

	const res = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
		body
	});

	const data = await res.json();

	if (!res.ok) {
		console.error('[getUserToken] Token failed:', data);
		throw new Error(`User token failed: ${data.error_description || data.error}`);
	}
	
	// Decode token to check scopes
	try {
		const tokenParts = data.access_token.split('.');
		const payload = JSON.parse(atob(tokenParts[1]));
		console.log(`[getUserToken] ${scopeLabel} token scopes:`, payload.scp || payload.roles || 'no scopes');
	} catch (e) {
		console.log('[getUserToken] Could not decode token payload');
	}

	console.log(`[getUserToken] ${scopeLabel} user token obtained successfully`);
	return data.access_token;
}

async function getSharePointTokenWithKeyVault(env, scope) {
	console.log('[getSharePointToken] Fetching certificate from Key Vault...');

	const keyVaultToken = await getKeyVaultAccessToken(env);
	const certUrl = `${env.AZURE_KEY_VAULT_URL}/certificates/${env.AZURE_KEY_VAULT_CERT_NAME}/?api-version=7.0`;

	const certResponse = await fetch(certUrl, {
		headers: {
			Authorization: `Bearer ${keyVaultToken}`,
			'Content-Type': 'application/json',
		},
	});

	if (!certResponse.ok) {
		const error = await certResponse.text();
		console.error('[getSharePointToken] Key Vault certificate fetch failed:', error);
		throw new Error(`Key Vault certificate fetch failed (${certResponse.status}): ${error}`);
	}

	const certData = await certResponse.json();
	const certX5t = certData.x5t;

	if (!certX5t) {
		throw new Error('Certificate thumbprint (x5t) not found in Key Vault certificate');
	}

	console.log('[getSharePointToken] Got certificate thumbprint from Key Vault:', certX5t);

	const now = Math.floor(Date.now() / 1000);
	const header = {
		alg: 'RS256',
		typ: 'JWT',
		x5t: certX5t,
	};

	const jti = generateUUID();
	const payload = {
		aud: `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/v2.0`,
		exp: now + 3600,
		iss: env.AZURE_CLIENT_ID,
		jti,
		nbf: now,
		sub: env.AZURE_CLIENT_ID,
	};

	const signingInput = base64UrlEncode(JSON.stringify(header)) + '.' + base64UrlEncode(JSON.stringify(payload));
	console.log('[getSharePointToken] Signing JWT with Key Vault...');

	const signature = await signWithKeyVault(env, keyVaultToken, signingInput);
	const assertion = signingInput + '.' + signature;

	console.log('[getSharePointToken] Requesting SharePoint token with Key Vault certificate assertion');

	const response = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
		body: new URLSearchParams({
			grant_type: 'client_credentials',
			client_id: env.AZURE_CLIENT_ID,
			client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
			client_assertion: assertion,
			scope,
		}),
	});

	if (!response.ok) {
		const error = await response.text();
		console.error('[getSharePointToken] SharePoint token request failed:', error);
		throw new Error(`SharePoint token fetch failed (${response.status}): ${error}`);
	}

	const result = await response.json();
	console.log('[getSharePointToken] Successfully obtained SharePoint token via Key Vault');
	return result.access_token;
}

async function getKeyVaultAccessToken(env) {
	const scope = 'https://vault.azure.net/.default';

	if (env.AZURE_CLIENT_SECRET) {
		console.log('[getKeyVaultAccessToken] Using client secret for Key Vault');
		const response = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
			body: new URLSearchParams({
				grant_type: 'client_credentials',
				client_id: env.AZURE_CLIENT_ID,
				client_secret: env.AZURE_CLIENT_SECRET,
				scope,
			}),
		});

		if (!response.ok) {
			const error = await response.text();
			throw new Error(`Key Vault token failed (${response.status}): ${error}`);
		}

		const result = await response.json();
		return result.access_token;
	}

	throw new Error('AZURE_CLIENT_SECRET is required for Key Vault access');
}

async function signWithKeyVault(env, accessToken, dataToSign) {
	const signUrl = `${env.AZURE_KEY_VAULT_URL}/keys/${env.AZURE_KEY_VAULT_CERT_NAME}/sign?api-version=7.0`;

	const dataBytes = new TextEncoder().encode(dataToSign);
	const hashBuffer = await crypto.subtle.digest('SHA-256', dataBytes);
	const hashArray = new Uint8Array(hashBuffer);
	const hashBase64Url = btoa(String.fromCharCode(...hashArray))
		.replace(/\+/g, '-')
		.replace(/\//g, '_')
		.replace(/=+$/, '');

	const requestBody = {
		alg: 'RS256',
		value: hashBase64Url,
	};

	const response = await fetch(signUrl, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${accessToken}`,
			'Content-Type': 'application/json',
		},
		body: JSON.stringify(requestBody),
	});

	if (!response.ok) {
		const error = await response.text();
		console.error('[signWithKeyVault] Key Vault sign operation failed:', error);
		throw new Error(`Key Vault sign failed (${response.status}): ${error}`);
	}

	const result = await response.json();
	return result.value;
}

function base64UrlDecode(str) {
	const base64 = str.replace(/-/g, '+').replace(/_/g, '/');
	const padded = base64 + '=='.slice(0, (4 - base64.length % 4) % 4);
	return Uint8Array.from(atob(padded), c => c.charCodeAt(0));
}

async function getSharePointTokenWithSecret(env, scope) {
	const response = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
		body: new URLSearchParams({
			grant_type: 'client_credentials',
			client_id: env.AZURE_CLIENT_ID,
			client_secret: env.AZURE_CLIENT_SECRET,
			scope,
		}),
	});

	if (!response.ok) {
		const error = await response.text();
		console.error('[getSharePointToken] Secret auth failed:', error);
		throw new Error(`Token fetch failed (${response.status}): ${error}`);
	}

	const payload = await response.json();
	return payload.access_token;
}

async function getSharePointTokenWithCert(env, scope) {
	try {
		const now = Math.floor(Date.now() / 1000);

		const thumbprint = await getCertThumbprint(env.AZURE_CLIENT_CERTIFICATE);
		console.log('[getSharePointToken] Certificate thumbprint:', thumbprint);

		const header = {
			alg: 'RS256',
			typ: 'JWT',
			x5t: thumbprint,
		};

		const jti = generateUUID();
		console.log('[getSharePointToken] Generated JWT ID:', jti);

		const payload = {
			aud: `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/v2.0`,
			exp: now + 3600,
			iss: env.AZURE_CLIENT_ID,
			jti,
			nbf: now,
			sub: env.AZURE_CLIENT_ID,
		};

		const signingInput = base64UrlEncode(JSON.stringify(header)) + '.' + base64UrlEncode(JSON.stringify(payload));
		console.log('[getSharePointToken] Signing JWT...');

		const signature = await signJwt(signingInput, env.AZURE_CLIENT_CERTIFICATE, env.AZURE_CLIENT_CERTIFICATE_PASSWORD);
		const assertion = signingInput + '.' + signature;

		console.log('[getSharePointToken] Requesting token with certificate assertion');

		const response = await fetch(`https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`, {
			method: 'POST',
			headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
			body: new URLSearchParams({
				grant_type: 'client_credentials',
				client_id: env.AZURE_CLIENT_ID,
				client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
				client_assertion: assertion,
				scope,
			}),
		});

		if (!response.ok) {
			const error = await response.text();
			console.error('[getSharePointToken] Certificate auth failed:', error);
			throw new Error(`Certificate token fetch failed (${response.status}): ${error}`);
		}

		const result = await response.json();
		console.log('[getSharePointToken] Certificate auth successful, token received');
		return result.access_token;
	} catch (error) {
		console.error('[getSharePointToken] Certificate auth error:', error.message, error.stack);
		throw error;
	}
}

async function getCertThumbprint(pemCert) {
	const certBase64 = pemCert
		.replace(/-----BEGIN CERTIFICATE-----/, '')
		.replace(/-----END CERTIFICATE-----/, '')
		.replace(/\s/g, '');

	const certDer = Uint8Array.from(atob(certBase64), c => c.charCodeAt(0));

	const hashBuffer = await crypto.subtle.digest('SHA-1', certDer);
	const thumbprintBytes = new Uint8Array(hashBuffer);
	return btoa(String.fromCharCode(...thumbprintBytes));
}

function generateUUID() {
	const bytes = new Uint8Array(16);
	crypto.getRandomValues(bytes);
	bytes[6] = (bytes[6] & 0x0f) | 0x40;
	bytes[8] = (bytes[8] & 0x3f) | 0x80;
	const hex = Array.from(bytes).map(b => b.toString(16).padStart(2, '0')).join('');
	return `${hex.slice(0,8)}-${hex.slice(8,12)}-${hex.slice(12,16)}-${hex.slice(16,20)}-${hex.slice(20)}`;
}

async function signJwt(input, privateKeyPem, password) {
	let keyData;
	let keyType = 'unknown';

	if (privateKeyPem.includes('-----BEGIN')) {
		if (privateKeyPem.includes('ENCRYPTED PRIVATE KEY')) {
			console.log('[signJwt] Processing ENCRYPTED PEM private key');
			if (!password) {
				throw new Error('AZURE_CLIENT_CERTIFICATE_PASSWORD is required for encrypted private key');
			}
			keyData = await decryptPemPrivateKey(privateKeyPem, password);
			keyType = 'encrypted-pem';
		} else if (privateKeyPem.includes('PRIVATE KEY')) {
			console.log('[signJwt] Processing unencrypted PEM private key');
			keyData = decodePemPrivateKey(privateKeyPem);
			keyType = 'pem';
		} else {
			throw new Error('Unknown PEM format in private key');
		}
	} else {
		console.log('[signJwt] Processing base64-encoded PFX/PKCS12');
		if (!password) {
			throw new Error('AZURE_CLIENT_CERTIFICATE_PASSWORD is required for PFX file');
		}
		keyData = await extractKeyFromPfx(privateKeyPem, password);
		keyType = 'pfx';
	}

	console.log(`[signJwt] Key type: ${keyType}, length: ${keyData.length}`);

	try {
		const key = await crypto.subtle.importKey(
			'pkcs8',
			keyData,
			{ name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
			false,
			['sign']
		);

		const encoder = new TextEncoder();
		const signature = await crypto.subtle.sign(
			'RSASSA-PKCS1-v1_5',
			key,
			encoder.encode(input)
		);

		return base64UrlEncode(new Uint8Array(signature));
	} catch (importError) {
		console.error('[signJwt] Key import failed:', importError.message);
		console.error('[signJwt] First 20 bytes (hex):', Array.from(keyData.slice(0, 20)).map(b => b.toString(16).padStart(2, '0')).join(' '));
		throw importError;
	}
}

function decodePemPrivateKey(pem) {
	let pemContents = pem
		.replace(/-----BEGIN PRIVATE KEY-----/, '')
		.replace(/-----END PRIVATE KEY-----/, '')
		.replace(/-----BEGIN RSA PRIVATE KEY-----/, '')
		.replace(/-----END RSA PRIVATE KEY-----/, '')
		.replace(/\s/g, '');

	return Uint8Array.from(atob(pemContents), c => c.charCodeAt(0));
}

async function decryptPemPrivateKey(pem, password) {
	const pemContents = pem
		.replace(/-----BEGIN ENCRYPTED PRIVATE KEY-----/, '')
		.replace(/-----END ENCRYPTED PRIVATE KEY-----/, '')
		.replace(/\s/g, '');

	const encryptedDer = Uint8Array.from(atob(pemContents), c => c.charCodeAt(0));
	const pwdBytes = new TextEncoder().encode(password);

	const key = await crypto.subtle.importKey(
		'raw',
		await crypto.subtle.digest('SHA-256', pwdBytes),
		{ name: 'AES-CBC' },
		false,
		['decrypt']
	);

	const iv = encryptedDer.slice(0, 16);
	const ciphertext = encryptedDer.slice(16);

	const decrypted = await crypto.subtle.decrypt(
		{ name: 'AES-CBC', iv },
		key,
		ciphertext
	);

	return new Uint8Array(decrypted);
}

async function extractKeyFromPfx(pfxBase64, password) {
	console.log('[extractKeyFromPfx] Decoding PFX...');
	const pfxDer = Uint8Array.from(atob(pfxBase64), c => c.charCodeAt(0));
	console.log('[extractKeyFromPfx] PFX length:', pfxDer.length, 'bytes');

	const pwdBytes = new TextEncoder().encode(password);

	const pkcs12Info = parsePkcs12(pfxDer);
	console.log('[extractKeyFromPfx] Parsed PFX, keybags found:', pkcs12Info.keyBags.length);

	for (const bag of pkcs12Info.keyBags) {
		if (bag.encrypted) {
			try {
				const decrypted = await pkcs12Decrypt(bag.data, pwdBytes);
				if (decrypted) {
					console.log('[extractKeyFromPfx] Successfully decrypted key bag');
					return decrypted;
				}
			} catch (e) {
				console.log('[extractKeyFromPfx] Decryption failed:', e.message);
			}
		} else {
			console.log('[extractKeyFromPfx] Found unencrypted key bag');
			return bag.data;
		}
	}

	for (const bag of pkcs12Info.keyBags) {
		try {
			const decrypted = await pkcs12Decrypt(bag.data, new Uint8Array(0));
			if (decrypted) {
				console.log('[extractKeyFromPfx] Successfully decrypted with empty password');
				return decrypted;
			}
		} catch (e) {
		}
	}

	throw new Error('Failed to extract private key from PFX. Check your password.');
}

function parsePkcs12(der) {
	const result = { keyBags: [] };
	let offset = 0;

	if (der[offset] !== 0x30 || der[offset + 1] !== 0x82) {
		throw new Error('Invalid PFX structure');
	}
	offset += 4;

	const parseLength = (data, pos) => {
		let len = data[pos];
		if (len & 0x80) {
			const numBytes = len & 0x7f;
			len = 0;
			for (let i = 0; i < numBytes; i++) {
				len = (len << 8) | data[pos + 1 + i];
			}
			return { length: len, next: pos + 1 + numBytes, headerLen: 1 + numBytes };
		}
		return { length: len, next: pos + 1, headerLen: 1 };
	};

	const pfxSeqLen = parseLength(der, offset);
	const pfxEnd = offset + pfxSeqLen.headerLen + pfxSeqLen.length;
	offset = pfxSeqLen.next;

	while (offset < pfxEnd) {
		if (der[offset] !== 0x30) break;
		const seqLen = parseLength(der, offset);
		const seqEnd = offset + seqLen.headerLen + seqLen.length;
		offset = seqLen.next;

		if (offset >= seqEnd) break;

		const numSetItems = der[offset++];
		for (let i = 0; i < numSetItems && offset < seqEnd; i++) {
			if (der[offset] !== 0x30) break;
			const setSeqLen = parseLength(der, offset);
			const setEnd = offset + setSeqLen.headerLen + setSeqLen.length;
			offset = setSeqLen.next;

			while (offset < setEnd) {
				if (der[offset] !== 0x30) break;
				const itemLen = parseLength(der, offset);
				const itemEnd = offset + itemLen.headerLen + itemLen.length;
				offset = itemLen.next;

				if (offset >= itemEnd) break;

				const oidLen = der[offset++];
				const oid = der.slice(offset, offset + oidLen);
				offset += oidLen;

				const isKeyBag = oid.length >= 7 && 
					oid[oid.length - 5] === 0x2a && oid[oid.length - 4] === 0x86 && 
					oid[oid.length - 3] === 0x48 && oid[oid.length - 2] === 0x0f && oid[oid.length - 1] === 0x01;

				if (der[offset++] !== 0x04) continue;
				const octetLen = parseLength(der, offset);
				const octetEnd = offset + octetLen.headerLen + octetLen.length;
				offset = octetLen.next;

				const bagContents = der.slice(offset, octetEnd);
				offset = octetEnd;

				const isEncrypted = bagContents[0] !== 0x30;
				result.keyBags.push({ data: new Uint8Array(bagContents), encrypted: isEncrypted });
			}
			offset = setEnd;
		}
		offset = seqEnd;
	}

	return result;
}

async function pkcs12Decrypt(encryptedData, password) {
	const SALT = new Uint8Array([0x30, 0xa5, 0x83, 0x02, 0xa2, 0x01, 0x34]);
	const ITERATIONS = 100;

	const deriveKey = async (password, salt, keyLen) => {
		let key = new Uint8Array(keyLen);
		let block = new Uint8Array(salt.length + 2);
		block.set(salt);

		for (let i = 0; i < Math.ceil(keyLen / 20); i++) {
			block[salt.length] = ((i + 1) >> 8) & 0xff;
			block[salt.length + 1] = (i + 1) & 0xff;

			const combined = new Uint8Array([...block, ...password]);
			let hash = await crypto.subtle.digest('SHA-1', combined);

			for (let j = 1; j < ITERATIONS; j++) {
				hash = await crypto.subtle.digest('SHA-1', hash);
			}

			const result = new Uint8Array(hash);
			const offset = i * 20;
			const copyLen = Math.min(20, keyLen - offset);
			key.set(result.slice(0, copyLen), offset);
		}

		return key;
	};

	const key = await deriveKey(password, SALT, 32);
	const iv = await deriveKey(password, SALT, 16);

	const aesKey = await crypto.subtle.importKey(
		'raw',
		key,
		'AES-CBC',
		false,
		['decrypt']
	);

	const decrypted = await crypto.subtle.decrypt(
		{ name: 'AES-CBC', iv },
		aesKey,
		encryptedData
	);

	const result = new Uint8Array(decrypted);
	const paddingLen = result[result.length - 1];
	return result.slice(0, result.length - paddingLen);
}

function base64UrlEncode(data) {
	const bytes = typeof data === 'string' ? new TextEncoder().encode(data) : data;
	let binary = '';
	for (let i = 0; i < bytes.length; i++) {
		binary += String.fromCharCode(bytes[i]);
	}
	return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}



function buildFolderName(projectTitle, submittedAt) {
	const safeTitle = String(projectTitle || 'IC Report').replace(/[^a-zA-Z0-9 _-]/g, '').trim() || 'IC Report';
	const safeDate = new Date(submittedAt);
	const datePart = Number.isNaN(safeDate.getTime()) ? new Date().toISOString().slice(0, 10) : safeDate.toISOString().slice(0, 10);
	return `${safeTitle} - ${datePart}`;
}

function sanitizeFileName(fileName) {
	return String(fileName || 'upload.bin').replace(/[^a-zA-Z0-9._-]/g, '_');
}

function normalizePath(pathValue) {
	const normalized = String(pathValue || '').trim().replace(/\\/g, '/').replace(/\/+/g, '/');
	if (!normalized) return '';
	return normalized.startsWith('/') ? normalized : `/${normalized}`;
}

// ────────────────────────────────────────────────────────────────────────────
// Confirmation email
// ────────────────────────────────────────────────────────────────────────────

async function sendConfirmationEmail(fields, env, token) {
  const isTestMode = env.TEST_MODE === "true";

  const totalCoordTrip =
    (toNum(fields.ticketsCost)       || 0) +
    (toNum(fields.fuelCost)          || 0) +
    (toNum(fields.accommodationCost) || 0) +
    (toNum(fields.foodCost)          || 0);

  const dateRange = fields.projectDateFrom && fields.projectDateTo
    ? `${fields.projectDateFrom} – ${fields.projectDateTo}`
    : fields.projectDateFrom || "";

  const testimoniesHtml = (fields.testimonies || [])
    .map(t => `
      <div style="border-left:3px solid #c8a96e;padding-left:16px;margin-bottom:18px;">
        ${t.author ? `<p style="font-weight:600;color:#1a3a5c;margin-bottom:6px;">${t.author}</p>` : ""}
        <p style="white-space:pre-wrap;color:#333;">${t.text}</p>
      </div>`)
    .join("");

  // ── Test mode banner (injected at top of email body) ──────────────────────
  const testBanner = isTestMode ? `
  <div style="background:#f5a623;padding:14px 40px;border-bottom:3px solid #c8820a;">
    <p style="margin:0;font-family:Georgia,serif;font-size:14px;font-weight:bold;color:#1a1a1a;">
      ⚠ TEST MODE — This email would normally go to: ${env.EMAIL_RECIPIENT}
    </p>
  </div>` : "";

  const emailBody = `
<html><body style="font-family:Georgia,serif;color:#1a1a1a;max-width:680px;margin:0 auto;">
  ${testBanner}
  <div style="background:#1a3a5c;padding:32px 40px;">
    <h1 style="color:#fff;margin:0;font-size:22px;letter-spacing:1px;">IC PROJECT REPORT RECEIVED</h1>
    <p style="color:#a8c4e0;margin:8px 0 0;">Submitted ${new Date(fields.submittedAt).toLocaleString("en-US",{dateStyle:"long",timeStyle:"short"})}</p>
  </div>
  <div style="padding:32px 40px;background:#f9f7f4;">
    <h2 style="color:#1a3a5c;border-bottom:2px solid #c8a96e;padding-bottom:8px;">${fields.projectTitle || "IC Project Report"}</h2>
    <p style="color:#555;margin-top:-8px;margin-bottom:20px;">${fields.city || ""}${fields.country ? `, ${fields.country}` : ""}${fields.area ? ` &nbsp;·&nbsp; ${fields.area}` : ""}${dateRange ? ` &nbsp;·&nbsp; ${dateRange}` : ""}</p>
    <h3 style="color:#1a3a5c;margin-top:24px;">Introduction</h3>
    <p style="white-space:pre-wrap;">${fields.introduction}</p>
    <h3 style="color:#1a3a5c;margin-top:28px;border-bottom:1px solid #ddd;padding-bottom:6px;">Statistics</h3>
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      ${statRow("# of Churches Who Participated",         fields.churchesParticipated)}
      ${statRow("# of National Project Participants",     fields.nationalParticipants)}
      ${statRow("# of USA Participants",                  fields.usaParticipants)}
      ${statRow("# of Participants From Other Countries", fields.otherCountriesParticipants)}
      ${statRow("# of People Who Heard the Gospel",       fields.peopleHeardGospel)}
      ${statRow("# of Professions of Faith",              fields.professionsOfFaith)}
      ${statRow("# of Rededications to Christ",           fields.rededications)}
      ${statRow("# of Baptisms",                          fields.baptisms)}
      ${statRow("# of New Churches Planted",              fields.newChurchesPlanted)}
    </table>
    ${testimoniesHtml ? `<h3 style="color:#1a3a5c;margin-top:28px;">Testimonies (${fields.testimonies.length})</h3>${testimoniesHtml}` : ""}
    <h3 style="color:#1a3a5c;margin-top:28px;border-bottom:1px solid #ddd;padding-bottom:6px;">Financial Report</h3>
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      ${moneyRow("Total Funds Sent by IC",                fields.totalFundsSent, true)}
      ${moneyRow("Spent on Materials",                    fields.spentOnMaterials)}
      ${moneyRow("Coordinator Trips — Tickets",           fields.ticketsCost)}
      ${moneyRow("Coordinator Trips — Fuel",              fields.fuelCost)}
      ${moneyRow("Coordinator Trips — Accommodation",     fields.accommodationCost)}
      ${moneyRow("Coordinator Trips — Food",              fields.foodCost)}
      ${moneyRow("Coordinator Trips — Total",             totalCoordTrip.toFixed(2))}
      ${moneyRow("Financial Help to Participants",        fields.financialHelpParticipants)}
      ${statRow("# of Participants Receiving Help",       fields.numParticipantsHelp)}
      ${moneyRow("Rallies Expenses",                      fields.ralliesExpenses)}
      ${fields.ralliesDescription ? statRow("Rallies Description", fields.ralliesDescription) : ""}
      ${moneyRow("Additional Expenses",                   fields.additionalExpenses)}
      ${fields.additionalNeedDescription ? statRow("Additional Need", fields.additionalNeedDescription) : ""}
    </table>
  </div>
  <div style="padding:20px 40px;background:#1a3a5c;color:#a8c4e0;font-size:12px;">
    <p style="margin:0;">Submitted by ${fields.coordinatorName || "coordinator"} · IC Project Report System${isTestMode ? " · TEST MODE" : ""}</p>
  </div>
</body></html>`;

  // ── Routing: test mode sends only to TEST_EMAIL_RECIPIENT, no CC ──────────
  const recipient = isTestMode ? env.TEST_EMAIL_RECIPIENT : env.EMAIL_RECIPIENT;
  const subject   = isTestMode
    ? `[TEST] IC Project Report: ${fields.projectTitle || "New Submission"} — ${new Date().toLocaleDateString("en-US")}`
    : `IC Project Report: ${fields.projectTitle || "New Submission"} — ${new Date().toLocaleDateString("en-US")}`;

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(env.EMAIL_SENDER)}/sendMail`,
    {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        message: {
          subject,
          body:         { contentType: "HTML", content: emailBody },
          toRecipients: [{ emailAddress: { address: recipient } }],
          // In test mode, suppress CC so coordinators don't receive test emails
          ...(!isTestMode && fields.coordinatorEmail && {
            ccRecipients: [{ emailAddress: { address: fields.coordinatorEmail, name: fields.coordinatorName || undefined } }],
          }),
        },
        saveToSentItems: true,
      }),
    }
  );
  if (!res.ok) throw new Error(`sendMail failed (${res.status}): ${await res.text()}`);
  return { sent: true, testMode: isTestMode, recipient };
}

// ────────────────────────────────────────────────────────────────────────────
// Utilities
// ────────────────────────────────────────────────────────────────────────────

async function getAccessToken(env) {
  console.log('>>> GET ACCESS TOKEN: graph');
  const tenantId = env.AZURE_TENANT_ID;
  const clientId = env.AZURE_CLIENT_ID;
  const clientSecret = env.AZURE_CLIENT_SECRET;
  const graphScope = 'https://graph.microsoft.com/.default';

  console.log(`>>> Token scope: ${graphScope}`);
  return fetchTokenV2(tenantId, clientId, clientSecret, graphScope);
}

async function fetchTokenV2(tenantId, clientId, clientSecret, scope) {
  const res = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope,
      }),
    }
  );
  if (!res.ok) throw new Error(`Token fetch failed (${res.status}): ${await res.text()}`);
  return (await res.json()).access_token;
}

async function graphFetch(url, options = {}) {
  const res = await fetch(url, options);
  if (!res.ok) throw new Error(`Graph ${res.status} at ${url}: ${await res.text().catch(() => "")}`);
  return res.json();
}

function toNum(val) {
  if (val === null || val === undefined || val === "") return null;
  const n = parseFloat(val);
  return isNaN(n) ? null : n;
}

function corsResponse(body, status, env) {
  return new Response(body ? JSON.stringify(body) : null, {
    status,
    headers: {
      "Access-Control-Allow-Origin":  env?.ALLOWED_ORIGIN || "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
      "Content-Type": "application/json",
    },
  });
}

function statRow(label, value) {
  return `<tr style="border-bottom:1px solid #eee;">
    <td style="padding:7px 4px;color:#555;">${label}</td>
    <td style="padding:7px 4px;font-weight:600;text-align:right;">${value || "—"}</td>
  </tr>`;
}

function moneyRow(label, value, bold = false) {
  const n = parseFloat(value);
  const display = isNaN(n) ? "—" : `USD $${n.toLocaleString("en-US", { minimumFractionDigits: 2 })}`;
  return `<tr style="border-bottom:1px solid #eee;${bold ? "background:#f0ebe2;" : ""}">
    <td style="padding:7px 4px;color:#555;${bold ? "font-weight:700;" : ""}">${label}</td>
    <td style="padding:7px 4px;font-weight:${bold ? "700" : "600"};text-align:right;">${display}</td>
  </tr>`;
}
