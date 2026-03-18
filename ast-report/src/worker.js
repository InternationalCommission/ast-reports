/**
 * IC Project Report — Cloudflare Worker
 *
 * POST /          — Submit a report (public, CORS-restricted)
 * GET  /reports   — Fetch all reports for admin viewer (requires Azure AD JWT)
 * GET  /reports/:id — Fetch single report by SharePoint item ID
 *
 * Required Environment Variables (wrangler secret put ...):
 *   AZURE_TENANT_ID         - Azure AD tenant ID
 *   AZURE_CLIENT_ID         - App registration client ID (backend + API scope audience)
 *   AZURE_CLIENT_SECRET     - App registration client secret (used for SharePoint/Graph)
 *   ADMIN_CLIENT_ID         - App registration client ID for the admin SPA
 *                             (can be the same as AZURE_CLIENT_ID if using one app)
 *   SHAREPOINT_SITE_URL     - e.g. https://yourorg.sharepoint.com/sites/yoursite
 *   SHAREPOINT_LIST_NAME    - Target list name, e.g. "IC Project Reports"
 *   SHAREPOINT_FOLDER_PATH  - Server-relative folder for photo uploads
 *   EMAIL_SENDER            - Licensed M365 mailbox to send from
 *   EMAIL_RECIPIENT         - Where confirmation emails go
 *   ALLOWED_ORIGIN          - Your website origin for CORS (form + admin)
 */

export default {
  async fetch(request, env) {
    const url    = new URL(request.url);
    const path   = url.pathname.replace(/\/$/, "") || "/";
    const method = request.method;

    if (method === "OPTIONS") return corsResponse(null, 204, env);

    if (method === "POST" && path === "/")          return handleSubmit(request, env);
    if (method === "GET"  && path === "/reports")   return handleGetReports(request, env, url);
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
  const origin = request.headers.get("Origin") || "";
  if (env.ALLOWED_ORIGIN && origin !== env.ALLOWED_ORIGIN) {
    return corsResponse({ error: "Forbidden origin" }, 403, env);
  }

  try {
    const formData = await request.formData();
    const fields   = extractFields(formData);
    const photos   = formData.getAll("photos");

    const token = await getAccessToken(env);

    // Run all three in parallel
    const [listItemResult, uploadResults, emailResult] = await Promise.allSettled([
      createSharePointListItem(fields, env, token),
      uploadPhotos(photos, fields.projectTitle, env, token),
      sendConfirmationEmail(fields, env, token),
    ]);

    // If both list item and uploads succeeded, patch the folder URL back onto
    // the list item so the admin viewer can find the photos later
    if (listItemResult.status === "fulfilled" && uploadResults.status === "fulfilled") {
      const { folderWebUrl, driveId, folderItemId } = uploadResults.value;
      const listItemId = listItemResult.value?.id;
      if (folderWebUrl && listItemId) {
        await patchPhotoFolder(listItemId, folderWebUrl, driveId, folderItemId, env, token)
          .catch(e => console.warn("Could not patch photo folder URL:", e.message));
      }
    }

    const errors = [];
    if (listItemResult.status === "rejected")
      errors.push({ step: "sharepoint_list", message: listItemResult.reason?.message });
    if (uploadResults.status === "rejected")
      errors.push({ step: "file_upload",     message: uploadResults.reason?.message });
    if (emailResult.status === "rejected")
      errors.push({ step: "email",           message: emailResult.reason?.message });

    const listItemId  = listItemResult.status  === "fulfilled" ? listItemResult.value?.id   : null;
    const uploadValue = uploadResults.status   === "fulfilled" ? uploadResults.value         : null;

    return corsResponse(
      {
        success:       errors.length === 0,
        message:       errors.length === 0 ? "Report submitted successfully." : "Report submitted with some issues.",
        listItemId:    listItemId ?? null,
        uploadedFiles: uploadValue?.uploaded ?? [],
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
    const endpoint = cursor
      ? decodeURIComponent(cursor)
      : `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`
          + `?expand=fields`
          + `&$top=${top}`;

    const res   = await graphFetch(endpoint, { headers });

    // Sort descending by SubmittedAt client-side since the column isn't indexed
    const items = (res.value || [])
      .map(item => normalizeItem(item))
      .sort((a, b) => new Date(b.submittedAt || 0).getTime() - new Date(a.submittedAt || 0).getTime());

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

  return null; // ✓ valid
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
    location:        f("location"),
    projectDateFrom: f("projectDateFrom"),
    projectDateTo:   f("projectDateTo"),
    introduction:    f("introduction"),
    churchesParticipated:        f("churchesParticipated"),
    localities:                  f("localities"),
    nationalParticipants:        f("nationalParticipants"),
    usaParticipants:             f("usaParticipants"),
    otherCountriesParticipants:  f("otherCountriesParticipants"),
    totalVisits:                 f("totalVisits"),
    peopleHeardGospel:           f("peopleHeardGospel"),
    professionsOfFaith:          f("professionsOfFaith"),
    rededications:               f("rededications"),
    baptisms:                    f("baptisms"),
    newChurchesPlanted:          f("newChurchesPlanted"),
    testimonies,
    testimoniesJson: JSON.stringify(testimonies),
    totalFundsSent:              f("totalFundsSent"),
    spentOnMaterials:            f("spentOnMaterials"),
    ticketsCost:                 f("ticketsCost"),
    fuelCost:                    f("fuelCost"),
    accommodationCost:           f("accommodationCost"),
    foodCost:                    f("foodCost"),
    financialHelpParticipants:   f("financialHelpParticipants"),
    numParticipantsHelp:         f("numParticipantsHelp"),
    ralliesExpenses:             f("ralliesExpenses"),
    ralliesDescription:          f("ralliesDescription"),
    additionalExpenses:          f("additionalExpenses"),
    additionalNeedDescription:   f("additionalNeedDescription"),
    submittedAt: new Date().toISOString(),
  };
}

// ────────────────────────────────────────────────────────────────────────────
// SharePoint helpers
// ────────────────────────────────────────────────────────────────────────────

const sharePointFields = [
  "Title","Location","ProjectDateFrom","ProjectDateTo","Introduction",
  "ChurchesParticipated","Localities","NationalParticipants","USAParticipants",
  "OtherCountriesParticipants","TotalVisits","PeopleHeardGospel",
  "ProfessionsOfFaith","Rededications","Baptisms","NewChurchesPlanted",
  "Testimonies",
  "TotalFundsSent","SpentOnMaterials","TicketsCost","FuelCost",
  "AccommodationCost","FoodCost","FinancialHelpParticipants","NumParticipantsHelp",
  "RalliesExpenses","RalliesDescription","AdditionalExpenses","AdditionalNeedDescription",
  "CoordinatorName","CoordinatorEmail","SubmittedAt",
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
    location:                   f.Location,
    projectDateFrom:            f.ProjectDateFrom,
    projectDateTo:              f.ProjectDateTo,
    introduction:               f.Introduction,
    churchesParticipated:       f.ChurchesParticipated,
    localities:                 f.Localities,
    nationalParticipants:       f.NationalParticipants,
    usaParticipants:            f.USAParticipants,
    otherCountriesParticipants: f.OtherCountriesParticipants,
    totalVisits:                f.TotalVisits,
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
    photoFolderUrl:             f.PhotoFolderUrl   || null,
    photoDriveId:               f.PhotoDriveId     || null,
    photoFolderItemId:          f.PhotoFolderItemId|| null,
  };
}

// ── Patch photo folder info back onto the list item after upload ──────────────
async function patchPhotoFolder(itemId, folderWebUrl, driveId, folderItemId, env, token) {
  const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Accept: "application/json" };
  const { siteId, listId } = await resolveListIds(env, token);
  await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
    {
      method: "PATCH",
      headers,
      body: JSON.stringify({
        PhotoFolderUrl:    folderWebUrl,
        PhotoDriveId:      driveId,
        PhotoFolderItemId: folderItemId,
      }),
    }
  );
}

// ── GET /reports/:id/photos — list photos in a report's folder ────────────────
async function handleGetPhotos(request, env, id) {
  const authError = await validateAzureToken(request, env);
  if (authError) return corsResponse({ error: authError }, 401, env);

  try {
    const token = await getAccessToken(env);
    const { siteId, listId } = await resolveListIds(env, token);
    const headers = { Authorization: `Bearer ${token}`, Accept: "application/json" };

    // Fetch the list item to get the folder drive/item IDs
    const item = await graphFetch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${id}?expand=fields($select=PhotoDriveId,PhotoFolderItemId,PhotoFolderUrl)`,
      { headers }
    );

    const driveId      = item.fields?.PhotoDriveId;
    const folderItemId = item.fields?.PhotoFolderItemId;
    const folderUrl    = item.fields?.PhotoFolderUrl;

    if (!driveId || !folderItemId) {
      return corsResponse({ photos: [], folderUrl: null }, 200, env);
    }

    // List children of the folder
    const childrenRes = await graphFetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderItemId}/children`
        + `?$select=id,name,webUrl,thumbnails,file`,
      { headers }
    );

    // Request thumbnails in the same call for faster rendering
    const photos = await Promise.all(
      (childrenRes.value || [])
        .filter(f => f.file) // files only, not subfolders
        .map(async f => {
          let thumbnail = null;
          try {
            const thumbRes = await graphFetch(
              `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${f.id}/thumbnails/0/medium`,
              { headers }
            );
            thumbnail = thumbRes.url || null;
          } catch { /* thumbnail unavailable */ }
          return { id: f.id, name: f.name, webUrl: f.webUrl, thumbnail };
        })
    );

    return corsResponse({ photos, folderUrl }, 200, env);
  } catch (err) {
    console.error("GetPhotos error:", err);
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
          Location:                    fields.location,
          ProjectDateFrom:             fields.projectDateFrom || null,
          ProjectDateTo:               fields.projectDateTo   || null,
          Introduction:                fields.introduction,
          ChurchesParticipated:        toNum(fields.churchesParticipated),
          Localities:                  toNum(fields.localities),
          NationalParticipants:        toNum(fields.nationalParticipants),
          USAParticipants:             toNum(fields.usaParticipants),
          OtherCountriesParticipants:  toNum(fields.otherCountriesParticipants),
          TotalVisits:                 toNum(fields.totalVisits),
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

// ────────────────────────────────────────────────────────────────────────────
// Photo upload
// ────────────────────────────────────────────────────────────────────────────

async function uploadPhotos(photoFiles, projectTitle, env, token) {
  const validFiles = photoFiles.filter(f => f?.size > 0 && f?.name);
  if (!validFiles.length) return { uploaded: [], folderPath: null, folderWebUrl: null };

  const headers  = { Authorization: `Bearer ${token}`, Accept: "application/json" };
  const hostname = new URL(env.SHAREPOINT_SITE_URL).hostname;
  const sitePath = new URL(env.SHAREPOINT_SITE_URL).pathname;

  const siteRes  = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`, { headers });
  const driveRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${siteRes.id}/drive`, { headers });
  const driveId  = driveRes.id;

  const safeTitle  = (projectTitle || "IC Report").replace(/[^a-zA-Z0-9 _-]/g, "").trim();
  const folderName = `${safeTitle} - ${new Date().toISOString().slice(0, 10)}`;

  // SHAREPOINT_FOLDER_PATH should be the path relative to the drive root
  // (i.e. relative to the Documents library), e.g. "Report Photos"
  // NOT the full server-relative path like /sites/reports/Documents/Report Photos
  const parentPath = decodeURIComponent(env.SHAREPOINT_FOLDER_PATH).replace(/^\//, "");

  console.log(`Creating folder "${folderName}" under "${parentPath}" on drive ${driveId}`);

  const folderRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${parentPath}:/children`,
    {
      method:  "POST",
      headers: { ...headers, "Content-Type": "application/json" },
      body: JSON.stringify({
        name:   folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }),
    }
  );

  if (!folderRes.ok) {
    const errText = await folderRes.text();
    console.error(`Folder creation failed (${folderRes.status}): ${errText}`);
    throw new Error(`Failed to create photo folder (${folderRes.status}) at "${parentPath}": ${errText}`);
  }

  const folderData    = await folderRes.json();
  const folderWebUrl  = folderData.webUrl;
  const folderDriveId = folderData.parentReference?.driveId || driveId;
  const folderItemId  = folderData.id;

  console.log(`Folder created: id=${folderItemId} url=${folderWebUrl}`);

  // Upload each file into the created folder
  const uploaded = [];
  for (const file of validFiles) {
    const safeName  = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
    const mimeType  = file.type || "application/octet-stream";
    const buffer    = await file.arrayBuffer();

    console.log(`Uploading ${safeName} (${buffer.byteLength} bytes, ${mimeType}) → folder item ${folderItemId}`);

    // Use a clean set of headers for the PUT — don't include Accept:json
    // as it can confuse the binary upload
    const uploadRes = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${folderDriveId}/items/${folderItemId}:/${encodeURIComponent(safeName)}:/content`,
      {
        method:  "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": mimeType,
          "Content-Length": String(buffer.byteLength),
        },
        body: buffer,
      }
    );

    const uploadText = await uploadRes.text();
    console.log(`Upload response ${uploadRes.status}: ${uploadText.slice(0, 200)}`);

    if (uploadRes.ok) {
      const data = JSON.parse(uploadText);
      uploaded.push({ name: safeName, webUrl: data.webUrl, id: data.id });
      console.log(`Uploaded ${safeName} → ${data.webUrl}`);
    } else {
      throw new Error(`Failed to upload "${safeName}" (${uploadRes.status}): ${uploadText}`);
    }
  }

  return { uploaded, folderWebUrl, driveId: folderDriveId, folderItemId };
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
    <p style="color:#555;margin-top:-8px;margin-bottom:20px;">${fields.location}${dateRange ? ` &nbsp;·&nbsp; ${dateRange}` : ""}</p>
    <h3 style="color:#1a3a5c;margin-top:24px;">Introduction</h3>
    <p style="white-space:pre-wrap;">${fields.introduction}</p>
    <h3 style="color:#1a3a5c;margin-top:28px;border-bottom:1px solid #ddd;padding-bottom:6px;">Statistics</h3>
    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      ${statRow("# of Churches Who Participated",         fields.churchesParticipated)}
      ${statRow("# of Localities",                        fields.localities)}
      ${statRow("# of National Project Participants",     fields.nationalParticipants)}
      ${statRow("# of USA Participants",                  fields.usaParticipants)}
      ${statRow("# of Participants From Other Countries", fields.otherCountriesParticipants)}
      ${statRow("# of Visits",                            fields.totalVisits)}
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
  const res = await fetch(
    `https://login.microsoftonline.com/${env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     env.AZURE_CLIENT_ID,
        client_secret: env.AZURE_CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
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