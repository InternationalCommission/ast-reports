const { app } = require('@azure/functions');

app.http('sharePointUpload', {
	methods: ['POST'],
	authLevel: 'function',
	handler: async (request, context) => {
		try {
			validateEnv(process.env, ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'SHAREPOINT_SITE_URL']);

			const formData = await request.formData();
			const projectTitle = String(formData.get('projectTitle') || 'IC Project Report').trim();
			const submittedAt = String(formData.get('submittedAt') || new Date().toISOString()).trim();
			const parentFolderPath = normalizeServerRelativePath(formData.get('sharePointFolderPath') || process.env.SHAREPOINT_FOLDER_PATH || '');
			const photos = formData.getAll('photos').filter(file => file?.size > 0 && file?.name);

			if (!parentFolderPath) {
				return jsonResponse({ error: 'sharePointFolderPath is required.' }, 400);
			}

			if (!photos.length) {
				return jsonResponse({ uploaded: [], errors: [], folderWebUrl: null, folderServerRelativePath: null });
			}

			const sharePointSiteUrl = process.env.SHAREPOINT_SITE_URL;
			const sharePointToken = await getAccessToken({
				tenantId: process.env.AZURE_TENANT_ID,
				clientId: process.env.AZURE_CLIENT_ID,
				clientSecret: process.env.AZURE_CLIENT_SECRET,
				scope: `${new URL(sharePointSiteUrl).origin}/.default`,
			});
			const requestDigest = await getRequestDigest(sharePointSiteUrl, sharePointToken);
			const folderName = buildFolderName(projectTitle, submittedAt);
			const folderServerRelativePath = await ensureFolder({
				sharePointSiteUrl,
				sharePointToken,
				requestDigest,
				parentFolderPath,
				folderName,
				context,
			});

			const uploaded = [];
			const errors = [];

			for (const file of photos) {
				const safeFileName = sanitizeFileName(file.name);
				try {
					const fileInfo = await uploadFile({
						sharePointSiteUrl,
						sharePointToken,
						requestDigest,
						folderServerRelativePath,
						fileName: safeFileName,
						contentType: file.type || 'application/octet-stream',
						buffer: Buffer.from(await file.arrayBuffer()),
					});
					uploaded.push({
						name: safeFileName,
						webUrl: buildAbsoluteUrl(sharePointSiteUrl, fileInfo.ServerRelativeUrl),
					});
				} catch (error) {
					context.error(`Upload failed for ${safeFileName}: ${error.message}`);
					errors.push({ file: safeFileName, error: error.message });
				}
			}

			return jsonResponse({
				uploaded,
				errors,
				folderWebUrl: buildAbsoluteUrl(sharePointSiteUrl, folderServerRelativePath),
				folderServerRelativePath,
			});
		} catch (error) {
			context.error('SharePoint upload function failed', error);
			return jsonResponse({ error: error.message || 'Upload failed.' }, 500);
		}
	},
});

function validateEnv(env, requiredKeys) {
	for (const key of requiredKeys) {
		if (!env[key]) throw new Error(`${key} is required.`);
	}
}

async function getAccessToken({ tenantId, clientId, clientSecret, scope }) {
	const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
		method: 'POST',
		headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
		body: new URLSearchParams({
			grant_type: 'client_credentials',
			client_id: clientId,
			client_secret: clientSecret,
			scope,
		}),
	});

	if (!response.ok) {
		throw new Error(`Token fetch failed (${response.status}): ${await response.text()}`);
	}

	const payload = await response.json();
	return payload.access_token;
}

async function getRequestDigest(sharePointSiteUrl, sharePointToken) {
	const response = await fetch(`${sharePointSiteUrl}/_api/contextinfo`, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
			Accept: 'application/json;odata=nometadata',
		},
	});

	if (!response.ok) {
		throw new Error(`Context info failed (${response.status}): ${await response.text()}`);
	}

	const payload = await response.json();
	const digest = payload?.FormDigestValue || payload?.d?.GetContextWebInformation?.FormDigestValue;
	if (!digest) throw new Error('SharePoint request digest was missing from contextinfo response.');
	return digest;
}

async function ensureFolder({ sharePointSiteUrl, sharePointToken, requestDigest, parentFolderPath, folderName, context }) {
	const folderServerRelativePath = normalizeServerRelativePath(`${parentFolderPath}/${folderName}`);
	const existingFolderResponse = await fetch(
		`${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(folderServerRelativePath)}')`,
		{
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				Accept: 'application/json;odata=nometadata',
			},
		}
	);

	if (existingFolderResponse.ok) {
		return folderServerRelativePath;
	}

	if (existingFolderResponse.status !== 404) {
		throw new Error(`Folder lookup failed (${existingFolderResponse.status}): ${await existingFolderResponse.text()}`);
	}

	context.log(`Creating SharePoint folder ${folderServerRelativePath}`);
	const createResponse = await fetch(
		`${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(parentFolderPath)}')/Folders/addUsingPath(decodedUrl='${escapeODataString(folderName)}')`,
		{
			method: 'POST',
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				Accept: 'application/json;odata=nometadata',
				'X-RequestDigest': requestDigest,
			},
		}
	);

	if (!createResponse.ok && createResponse.status !== 409) {
		throw new Error(`Folder create failed (${createResponse.status}): ${await createResponse.text()}`);
	}

	return folderServerRelativePath;
}

async function uploadFile({ sharePointSiteUrl, sharePointToken, requestDigest, folderServerRelativePath, fileName, contentType, buffer }) {
	const CHUNK_SIZE = 327680;
	const SMALL_FILE_THRESHOLD = 4194304;

	if (buffer.length <= SMALL_FILE_THRESHOLD) {
		return uploadFileSimple({ sharePointSiteUrl, sharePointToken, requestDigest, folderServerRelativePath, fileName, contentType, buffer });
	}

	return uploadFileChunked({ sharePointSiteUrl, sharePointToken, folderServerRelativePath, fileName, contentType, buffer, chunkSize: CHUNK_SIZE });
}

async function uploadFileSimple({ sharePointSiteUrl, sharePointToken, requestDigest, folderServerRelativePath, fileName, contentType, buffer }) {
	const response = await fetch(
		`${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(folderServerRelativePath)}')/Files/AddUsingPath(decodedurl='${escapeODataString(fileName)}',overwrite=true)`,
		{
			method: 'POST',
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				Accept: 'application/json;odata=nometadata',
				'Content-Type': contentType,
				'X-RequestDigest': requestDigest,
			},
			body: buffer,
		}
	);

	if (!response.ok) {
		throw new Error(`SharePoint upload failed (${response.status}): ${await response.text()}`);
	}

	return response.json();
}

async function uploadFileChunked({ sharePointSiteUrl, sharePointToken, folderServerRelativePath, fileName, contentType, buffer, chunkSize }) {
	const createSessionResponse = await fetch(
		`${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(folderServerRelativePath)}')/Files/CreateUploadSession(FileName='${escapeODataString(fileName)}')`,
		{
			method: 'POST',
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				Accept: 'application/json;odata=nometadata',
				'Content-Type': 'application/json',
			},
			body: JSON.stringify({
				deferCommit: false,
			}),
		}
	);

	if (!createSessionResponse.ok) {
		throw new Error(`CreateUploadSession failed (${createSessionResponse.status}): ${await createSessionResponse.text()}`);
	}

	const sessionData = await createSessionResponse.json();
	const uploadUrl = sessionData.uploadUrl;

	let offset = 0;
	const fileSize = buffer.length;

	while (offset < fileSize) {
		const end = Math.min(offset + chunkSize, fileSize) - 1;
		const chunk = buffer.slice(offset, end + 1);
		const contentRange = `bytes ${offset}-${end}/${fileSize}`;

		const putResponse = await fetch(uploadUrl, {
			method: 'PUT',
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				'Content-Length': chunk.length,
				'Content-Range': contentRange,
			},
			body: chunk,
		});

		if (putResponse.status !== 202 && putResponse.status !== 201 && putResponse.status !== 200) {
			throw new Error(`Chunk upload failed at offset ${offset} (${putResponse.status}): ${await putResponse.text()}`);
		}

		if (putResponse.status === 201 || putResponse.status === 200) {
			break;
		}

		const rangeHeader = putResponse.headers.get('nextExpectedRanges');
		if (rangeHeader) {
			const match = rangeHeader.match(/(\d+)-/);
			if (match) {
				offset = parseInt(match[1], 10);
			} else {
				offset += chunk.length;
			}
		} else {
			offset += chunk.length;
		}
	}

	const finalResponse = await fetch(uploadUrl, {
		method: 'GET',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
		},
	});

	if (!finalResponse.ok) {
		throw new Error(`Failed to get uploaded file info (${finalResponse.status})`);
	}

	return finalResponse.json();
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

function normalizeServerRelativePath(pathValue) {
	const normalized = String(pathValue || '').trim().replace(/\\/g, '/').replace(/\/+/g, '/');
	if (!normalized) return '';
	return normalized.startsWith('/') ? normalized : `/${normalized}`;
}

function escapeODataString(value) {
	return String(value || '').replace(/'/g, "''");
}

function buildAbsoluteUrl(siteUrl, serverRelativePath) {
	return `${new URL(siteUrl).origin}${serverRelativePath}`;
}

function jsonResponse(body, status = 200) {
	return {
		status,
		jsonBody: body,
	};
}
