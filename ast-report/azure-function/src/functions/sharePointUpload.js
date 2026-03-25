const { app } = require('@azure/functions');

app.http('sharePointUpload', {
	methods: ['POST'],
	authLevel: 'function',
	handler: async (request, context) => {
		const log = (msg, data) => {
			const timestamp = new Date().toISOString();
			const logEntry = data ? `${timestamp} | ${msg} | ${JSON.stringify(data)}` : `${timestamp} | ${msg}`;
			context.log(logEntry);
			console.log(logEntry);
		};

		try {
			log('=== SharePoint Upload Started ===');
			validateEnv(process.env, ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET', 'SHAREPOINT_SITE_URL']);

			const formData = await request.formData();
			const projectTitle = String(formData.get('projectTitle') || 'IC Project Report').trim();
			const submittedAt = String(formData.get('submittedAt') || new Date().toISOString()).trim();
			const parentFolderPath = normalizeServerRelativePath(formData.get('sharePointFolderPath') || process.env.SHAREPOINT_FOLDER_PATH || '');
			const photos = formData.getAll('photos').filter(file => file?.size > 0 && file?.name);

			log('Form data parsed', { projectTitle, submittedAt, parentFolderPath, photoCount: photos.length });

			if (!parentFolderPath) {
				return jsonResponse({ error: 'sharePointFolderPath is required.' }, 400);
			}

			if (!photos.length) {
				log('No photos to upload, returning empty result');
				return jsonResponse({ uploaded: [], errors: [], folderWebUrl: null, folderServerRelativePath: null });
			}

			const sharePointSiteUrl = process.env.SHAREPOINT_SITE_URL;
			log('Obtaining access token...');
			const sharePointToken = await getAccessToken({
				tenantId: process.env.AZURE_TENANT_ID,
				clientId: process.env.AZURE_CLIENT_ID,
				clientSecret: process.env.AZURE_CLIENT_SECRET,
				scope: `${new URL(sharePointSiteUrl).origin}/.default`,
			}, log);
			log('Access token obtained');

			log('Obtaining request digest...');
			const requestDigest = await getRequestDigest(sharePointSiteUrl, sharePointToken, log);
			log('Request digest obtained');

			const folderName = buildFolderName(projectTitle, submittedAt);
			log('Ensuring folder exists', { folderName, parentFolderPath });
			const folderServerRelativePath = await ensureFolder({
				sharePointSiteUrl,
				sharePointToken,
				requestDigest,
				parentFolderPath,
				folderName,
				context,
				log,
			});
			log('Folder ready', { folderServerRelativePath });

			const uploaded = [];
			const errors = [];

			for (let i = 0; i < photos.length; i++) {
				const file = photos[i];
				const safeFileName = sanitizeFileName(file.name);
				log(`Uploading file ${i + 1}/${photos.length}`, { fileName: safeFileName, size: file.size, type: file.type });

				try {
					const fileInfo = await uploadFile({
						sharePointSiteUrl,
						sharePointToken,
						requestDigest,
						folderServerRelativePath,
						fileName: safeFileName,
						contentType: file.type || 'application/octet-stream',
						buffer: Buffer.from(await file.arrayBuffer()),
						log,
					});
					log(`File uploaded successfully`, { fileName: safeFileName, serverRelativeUrl: fileInfo.ServerRelativeUrl });
					uploaded.push({
						name: safeFileName,
						webUrl: buildAbsoluteUrl(sharePointSiteUrl, fileInfo.ServerRelativeUrl),
					});
				} catch (error) {
					log(`Upload failed for ${safeFileName}`, { error: error.message });
					errors.push({ file: safeFileName, error: error.message });
				}
			}

			log('=== Upload Complete ===', { uploaded: uploaded.length, errors: errors.length });

			return jsonResponse({
				uploaded,
				errors,
				folderWebUrl: buildAbsoluteUrl(sharePointSiteUrl, folderServerRelativePath),
				folderServerRelativePath,
			});
		} catch (error) {
			log('SharePoint upload function failed', { error: error.message, stack: error.stack });
			return jsonResponse({ error: error.message || 'Upload failed.' }, 500);
		}
	},
});

function validateEnv(env, requiredKeys) {
	for (const key of requiredKeys) {
		if (!env[key]) throw new Error(`${key} is required.`);
	}
}

async function getAccessToken({ tenantId, clientId, clientSecret, scope }, log) {
	log('Fetching access token from Azure AD');
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

	log('Token response status', { status: response.status });

	if (!response.ok) {
		const errorText = await response.text();
		log('Token fetch failed', { status: response.status, error: errorText });
		throw new Error(`Token fetch failed (${response.status}): ${errorText}`);
	}

	const payload = await response.json();
	log('Token fetch successful', { hasToken: !!payload.access_token, expiresIn: payload.expires_in });
	return payload.access_token;
}

async function getRequestDigest(sharePointSiteUrl, sharePointToken, log) {
	log('Fetching request digest from SharePoint');
	const response = await fetch(`${sharePointSiteUrl}/_api/contextinfo`, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
			Accept: 'application/json;odata=nometadata',
		},
	});

	log('Context info response status', { status: response.status });

	if (!response.ok) {
		const errorText = await response.text();
		log('Context info failed', { status: response.status, error: errorText });
		throw new Error(`Context info failed (${response.status}): ${errorText}`);
	}

	const payload = await response.json();
	const digest = payload?.FormDigestValue || payload?.d?.GetContextWebInformation?.FormDigestValue;
	if (!digest) {
		log('Request digest missing from response', { payload });
		throw new Error('SharePoint request digest was missing from contextinfo response.');
	}
	log('Request digest obtained', { digestLength: digest.length });
	return digest;
}

async function ensureFolder({ sharePointSiteUrl, sharePointToken, requestDigest, parentFolderPath, folderName, context, log }) {
	const folderServerRelativePath = normalizeServerRelativePath(`${parentFolderPath}/${folderName}`);
	log('Checking if folder exists', { folderServerRelativePath });

	const existingFolderResponse = await fetch(
		`${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(folderServerRelativePath)}')`,
		{
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				Accept: 'application/json;odata=nometadata',
			},
		}
	);

	log('Folder check response', { status: existingFolderResponse.status, folderServerRelativePath });

	if (existingFolderResponse.ok) {
		log('Folder already exists');
		return folderServerRelativePath;
	}

	if (existingFolderResponse.status !== 404) {
		const errorText = await existingFolderResponse.text();
		log('Folder lookup failed with unexpected status', { status: existingFolderResponse.status, error: errorText });
		throw new Error(`Folder lookup failed (${existingFolderResponse.status}): ${errorText}`);
	}

	log('Creating folder', { folderServerRelativePath });
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

	log('Folder create response', { status: createResponse.status });

	if (!createResponse.ok && createResponse.status !== 409) {
		const errorText = await createResponse.text();
		log('Folder create failed', { status: createResponse.status, error: errorText });
		throw new Error(`Folder create failed (${createResponse.status}): ${errorText}`);
	}

	log('Folder ready', { folderServerRelativePath });
	return folderServerRelativePath;
}

async function uploadFile({ sharePointSiteUrl, sharePointToken, requestDigest, folderServerRelativePath, fileName, contentType, buffer, log }) {
	const CHUNK_SIZE = 327680;
	const SMALL_FILE_THRESHOLD = 4194304;

	log(`Upload method decision`, { fileName, fileSize: buffer.length, threshold: SMALL_FILE_THRESHOLD, willChunk: buffer.length > SMALL_FILE_THRESHOLD });

	if (buffer.length <= SMALL_FILE_THRESHOLD) {
		return uploadFileSimple({ sharePointSiteUrl, sharePointToken, requestDigest, folderServerRelativePath, fileName, contentType, buffer, log });
	}

	return uploadFileChunked({ sharePointSiteUrl, sharePointToken, folderServerRelativePath, fileName, contentType, buffer, chunkSize: CHUNK_SIZE, log });
}

async function uploadFileSimple({ sharePointSiteUrl, sharePointToken, requestDigest, folderServerRelativePath, fileName, contentType, buffer, log }) {
	const uploadUrl = `${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(folderServerRelativePath)}')/Files/AddUsingPath(decodedurl='${escapeODataString(fileName)}',overwrite=true)`;
	log('Simple upload starting', { fileName, url: uploadUrl, size: buffer.length, contentType });

	const response = await fetch(uploadUrl, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
			Accept: 'application/json;odata=nometadata',
			'Content-Type': contentType,
			'X-RequestDigest': requestDigest,
		},
		body: buffer,
	});

	log('Simple upload response', { status: response.status, statusText: response.statusText });

	if (!response.ok) {
		const errorText = await response.text();
		log('Simple upload failed', { status: response.status, error: errorText });
		throw new Error(`SharePoint upload failed (${response.status}): ${errorText}`);
	}

	const result = response.json();
	log('Simple upload successful', { fileName, result });
	return result;
}

async function uploadFileChunked({ sharePointSiteUrl, sharePointToken, folderServerRelativePath, fileName, contentType, buffer, chunkSize, log }) {
	const fileSize = buffer.length;
	log('Chunked upload starting', { fileName, fileSize, chunkSize, numChunks: Math.ceil(fileSize / chunkSize) });

	const sessionUrl = `${sharePointSiteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${escapeODataString(folderServerRelativePath)}')/Files/CreateUploadSession(FileName='${escapeODataString(fileName)}')`;
	log('Creating upload session', { url: sessionUrl });

	const createSessionResponse = await fetch(sessionUrl, {
		method: 'POST',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
			Accept: 'application/json;odata=nometadata',
			'Content-Type': 'application/json',
		},
		body: JSON.stringify({
			deferCommit: false,
		}),
	});

	log('Create session response', { status: createSessionResponse.status });

	if (!createSessionResponse.ok) {
		const errorText = await createSessionResponse.text();
		log('CreateUploadSession failed', { status: createSessionResponse.status, error: errorText });
		throw new Error(`CreateUploadSession failed (${createSessionResponse.status}): ${errorText}`);
	}

	const sessionData = await createSessionResponse.json();
	const uploadUrl = sessionData.uploadUrl;
	log('Upload session created', { uploadUrl, expirationDateTime: sessionData.expirationDateTime });

	let offset = 0;
	let chunkNumber = 0;

	while (offset < fileSize) {
		chunkNumber++;
		const end = Math.min(offset + chunkSize, fileSize) - 1;
		const chunk = buffer.slice(offset, end + 1);
		const contentRange = `bytes ${offset}-${end}/${fileSize}`;

		log(`Uploading chunk ${chunkNumber}`, { offset, end, chunkSize: chunk.length, contentRange });

		const putResponse = await fetch(uploadUrl, {
			method: 'PUT',
			headers: {
				Authorization: `Bearer ${sharePointToken}`,
				'Content-Length': chunk.length,
				'Content-Range': contentRange,
			},
			body: chunk,
		});

		log('Chunk upload response', { chunkNumber, status: putResponse.status, nextExpectedRanges: putResponse.headers.get('nextExpectedRanges') });

		if (putResponse.status !== 202 && putResponse.status !== 201 && putResponse.status !== 200) {
			const errorText = await putResponse.text();
			log('Chunk upload failed', { chunkNumber, offset, status: putResponse.status, error: errorText });
			throw new Error(`Chunk upload failed at offset ${offset} (${putResponse.status}): ${errorText}`);
		}

		if (putResponse.status === 201 || putResponse.status === 200) {
			log('Upload completed (final chunk)');
			break;
		}

		const rangeHeader = putResponse.headers.get('nextExpectedRanges');
		if (rangeHeader) {
			log('Received nextExpectedRanges', { rangeHeader });
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

	log('Fetching final file info', { uploadUrl });
	const finalResponse = await fetch(uploadUrl, {
		method: 'GET',
		headers: {
			Authorization: `Bearer ${sharePointToken}`,
		},
	});

	log('Final file info response', { status: finalResponse.status });

	if (!finalResponse.ok) {
		const errorText = await finalResponse.text();
		log('Failed to get uploaded file info', { status: finalResponse.status, error: errorText });
		throw new Error(`Failed to get uploaded file info (${finalResponse.status}): ${errorText}`);
	}

	const result = await finalResponse.json();
	log('Chunked upload successful', { fileName, result });
	return result;
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
