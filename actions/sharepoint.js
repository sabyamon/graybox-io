/* ************************************************************************
* ADOBE CONFIDENTIAL
* ___________________
*
* Copyright 2024 Adobe
* All Rights Reserved.
*
* NOTICE: All information contained herein is, and remains
* the property of Adobe and its suppliers, if any. The intellectual
* and technical concepts contained herein are proprietary to Adobe
* and its suppliers and are protected by all applicable intellectual
* property laws, including trade secret and copyright laws.
* Dissemination of this information or reproduction of this material
* is strictly forbidden unless prior written permission is obtained
* from Adobe.
************************************************************************* */

import { Headers } from 'node-fetch';
import fetch from 'node-fetch';
import { getAioLogger } from './utils.js';
import SharepointAuth from './sharepointAuth.js';
import util from 'util';
const SP_CONN_ERR_LST = ['ETIMEDOUT', 'ECONNRESET'];
const APP_USER_AGENT = 'NONISV|Adobe|MiloFloodgate/0.1.0';
const NUM_REQ_THRESHOLD = 5;
const RETRY_ON_CF = 3;
const TOO_MANY_REQUESTS = '429';
// Added for debugging rate limit headers
const LOG_RESP_HEADER = false;
let nextCallAfter = 0;
const itemIdMap = {};
const logger = getAioLogger();

class Sharepoint {
    constructor(appConfig) {
        this.appConfig = appConfig;
        this.sharepointAuth = new SharepointAuth(this.appConfig.getMsalConfig());
    }

    getSharepointAuth() {
        return this.sharepointAuth;
    }

    async getAuthorizedRequestOption({ body = null, json = true, method = 'GET' } = {}) {
        const appSpToken = await this.sharepointAuth.getAccessToken();
        const bearer = `Bearer ${appSpToken}`;

        const headers = new Headers();
        headers.append('Authorization', bearer);
        headers.append('User-Agent', APP_USER_AGENT);
        if (json) {
            headers.append('Accept', 'application/json');
            headers.append('Content-Type', 'application/json');
        }

        const options = {
            method,
            headers,
        };

        if (body) {
            options.body = typeof body === 'string' ? body : JSON.stringify(body);
        }

        return options;
    }

    async executeGQL(url, opts) {
        const options = await this.getAuthorizedRequestOption(opts);
        const res = await this.fetchWithRetry(url, options);
        if (!res.ok) {
            throw new Error(`Failed to execute ${url}`);
        }
        return res.json();
    }

    async getItemId(uri, path) {
        const key = `~${uri}~${path}~`;
        itemIdMap[key] = itemIdMap[key] || await this.executeGQL(`${uri}${path}?$select=id`);
        return itemIdMap[key]?.id;
    }

    async getFileData(filePath, isGraybox) {
        const sp = await this.appConfig.getSpConfig();
        console.log('sp in getFileData', util.inspect(sp, { depth: null, colors: true }));
        const options = await this.getAuthorizedRequestOption();
        const baseURI = isGraybox ? sp.api.directory.create.gbBaseURI : sp.api.directory.create.baseURI;
        logger.info(`Base URI in getFileData: ${baseURI}`);
        try {
            const resp = await this.fetchWithRetry(`${baseURI}${filePath}`, options);
            logger.info(`Response in getFileData: ${util.inspect(resp, { depth: null, colors: true })}`);
            logger.info(`Response status in getFileData: ${resp.ok}`);
            if (!resp.ok) {
                logger.error(`Failed to get file data for ${filePath}: ${resp.status}`);
                return { fileDownloadUrl: null, fileSize: 0 };
            }
            
            const json = await resp.json();
            logger.info(`JSON in getFileData: ${util.inspect(json, { depth: null, colors: true })}`);
            const fileDownloadUrl = json['@microsoft.graph.downloadUrl'];
            logger.info(`File download URL in getFileData: ${fileDownloadUrl}`);
            const fileSize = json.size;
            logger.info(`File size in getFileData: ${fileSize}`);
            
            if (!fileDownloadUrl) {
                logger.error(`No download URL found for ${filePath}`);
                return { fileDownloadUrl: null, fileSize: 0 };
            }
            
            return { fileDownloadUrl, fileSize };
        } catch (error) {
            logger.error(`Error getting file data for ${filePath}: ${error.message}`);
            return { fileDownloadUrl: null, fileSize: 0 };
        }
    }

    async getFileUsingDownloadUrl(downloadUrl) {
        if (!downloadUrl) {
            logger.error('No download URL provided');
            return null;
        }

        try {
            const options = await this.getAuthorizedRequestOption({ json: false });
            logger.info(`Options in getFileUsingDownloadUrl for ${downloadUrl}: ${util.inspect(options, { depth: null, colors: true })}`);
            const response = await this.fetchWithRetry(downloadUrl, options);
            // logger.info(`Response in getFileUsingDownloadUrl for ${downloadUrl}: ${util.inspect(response, { depth: null, colors: true })}`);
            /* if (!response.ok) {
                logger.error(`Failed to download file: ${response.status}`);
                return null;
            } */
            if (response) {
                // logger.info(`Response blob in getFileUsingDownloadUrl for ${downloadUrl}: ${util.inspect(response.blob(), { depth: null, colors: true })}`);
                // logger.error(`Failed to download file: ${response.status}`);
                return response.blob();
            }
            // logger.info(`Response blob in getFileUsingDownloadUrl: ${util.inspect(response.blob(), { depth: null, colors: true })}`);
            // return response.blob();
        } catch (error) {
            logger.error(`Error downloading file: ${error.message}`);
            return null;
        }
    }

    async createFolder(folder, isGraybox) {
        logger.info(`Creating folder: ${folder}`);
        const sp = await this.appConfig.getSpConfig();
        const options = await this.getAuthorizedRequestOption({ method: sp.api.directory.create.method });
        options.body = JSON.stringify(sp.api.directory.create.payload);

        const baseURI = isGraybox ? sp.api.directory.create.gbBaseURI : sp.api.directory.create.baseURI;
        logger.info(`Base URI in createFolder: ${baseURI}`); // https://graph.microsoft.com/v1.0/sites/adobe.sharepoint.com/root:/bacom-graybox
        const res = await this.fetchWithRetry(`${baseURI}${folder}`, options);
        logger.info(`Response in createFolder: ${util.inspect(res, { depth: null, colors: true })}`);
        if (res.ok) {
            return res.json();
        }
        throw new Error(`Could not create folder: ${folder}`);
    }

    getFolderFromPath(path) {
        // For paths with a file extension (containing a dot), return the directory path
        if (path.includes('.')) {
            return path.substring(0, path.lastIndexOf('/'));
        }
        // For paths without a file extension (like 'bacom-graybox/sabya'),
        // return the path as is since it's considered a folder path
        return path;
        // For the example 'bacom-graybox/sabya', this method will return 'bacom-graybox/sabya'
        // since it doesn't contain a dot and is treated as a folder path
    }

    getFileNameFromPath(path) {
        return path.split('/').pop().split('/').pop();
    }

    async createUploadSession(sp, file, dest, filename, isGraybox) {
        let fileSize = file.size;
        if (Buffer.isBuffer(file)) {
            fileSize = Buffer.byteLength(file);
        }

        const payload = {
            ...sp.api.file.createUploadSession.payload,
            description: 'Preview file',
            fileSize,
            name: filename,
        };
        const options = await this.getAuthorizedRequestOption({ method: sp.api.file.createUploadSession.method });
        options.body = JSON.stringify(payload);

        const baseURI = isGraybox ? sp.api.file.createUploadSession.gbBaseURI : sp.api.file.createUploadSession.baseURI;

        const createdUploadSession = await this.fetchWithRetry(`${baseURI}${dest}:/createUploadSession`, options);
        return createdUploadSession.ok ? createdUploadSession.json() : undefined;
    }

    async uploadFile(sp, uploadUrl, file) {
        const options = await this.getAuthorizedRequestOption({
            json: false,
            method: sp.api.file.upload.method,
        });
        let fileSize = file.size;
        if (Buffer.isBuffer(file)) {
            fileSize = Buffer.byteLength(file);
        }
        // TODO API is limited to 60Mb, for more, we need to batch the upload.
        options.headers.append('Content-Length', fileSize);
        options.headers.append('Content-Range', `bytes 0-${fileSize - 1}/${fileSize}`);
        options.headers.append('Prefer', 'bypass-shared-lock');
        options.body = file;
        return this.fetchWithRetry(`${uploadUrl}`, options);
    }

    async deleteFile(sp, filePath) {
        const options = await this.getAuthorizedRequestOption({
            json: false,
            method: sp.api.file.delete.method,
        });
        options.headers.append('Prefer', 'bypass-shared-lock');
        return fetch(filePath, options);
    }

    async createSessionAndUploadFile(sp, file, dest, filename, isGraybox) {
        const createdUploadSession = await this.createUploadSession(sp, file, dest, filename, isGraybox);
        const status = {};
        if (createdUploadSession) {
            const uploadSessionUrl = createdUploadSession.uploadUrl;
            if (!uploadSessionUrl) {
                return status;
            }
            status.sessionUrl = uploadSessionUrl;
            const uploadedFile = await this.uploadFile(sp, uploadSessionUrl, file);
            if (!uploadedFile) {
                return status;
            }
            if (uploadedFile.ok) {
                status.uploadedFile = await uploadedFile.json();
                status.success = true;
            } else if (uploadedFile.status === 423) {
                status.locked = true;
            }
        }
        return status;
    }

    async saveFileSimple(file, dest, isGraybox) {
        logger.info(`Saving file ${dest} to ${isGraybox}`);
        //dest = /demo-gb-bulk-copy/sabya/drafts/sabya-gb1-fragment.docx
        // actual path provided is - /sabya/drafts/fragments/sabya-gb1-fragment
        try {
            if (!file) {
                logger.error('No file content provided');
                return { success: false, path: dest, errorMsg: 'No file content provided' };
            }

            const folder = this.getFolderFromPath(dest);
            logger.info(`Folder in saveFileSimple: ${folder}`); // /demo-gb-bulk-copy/sabya/drafts
            const filename = this.getFileNameFromPath(dest);
            logger.info(`Filename in saveFileSimple: ${filename}`); // sabya-gb1-fragment.docx
            logger.info(`Saving file ${filename} to ${folder}`); // Saving file sabya-to-be-copied.docx to bacom-graybox/sabya

            // Ensure destination folder exists
            try {
                await this.createFolder(folder, isGraybox);
                logger.info(`Folder created: ${folder}`);
            } catch (error) {
                logger.error(`Error creating folder ${folder}: ${error.message}`);
                return { success: false, path: dest, errorMsg: `Failed to create destination folder: ${error.message}` };
            }

            const sp = await this.appConfig.getSpConfig();
            logger.info(`SP in saveFileSimple: ${util.inspect(sp, { depth: null, colors: true })}`);
            logger.info(`DEST in saveFileSimple: ${dest}`);
            logger.info(`filename in saveFileSimple: ${filename}`);
            logger.info(`isGraybox in saveFileSimple: ${isGraybox}`);
            const uploadFileStatus = await this.createSessionAndUploadFile(sp, file, dest, filename, isGraybox);
            logger.info(`Upload file status in saveFileSimple: ${util.inspect(uploadFileStatus, { depth: null, colors: true })}`);

            if (uploadFileStatus.locked) {
                logger.info(`Locked file detected: ${dest}`);
                return { success: false, path: dest, errorMsg: 'File is locked' };
            }

            const uploadedFileJson = uploadFileStatus.uploadedFile;
            if (uploadedFileJson) {
                return { 
                    success: true, 
                    uploadedFileJson, 
                    path: dest,
                    metadata: {
                        name: uploadedFileJson.name,
                        size: uploadedFileJson.size,
                        lastModifiedDateTime: uploadedFileJson.lastModifiedDateTime
                    }
                };
            }

            return { success: false, path: dest, errorMsg: 'Upload failed' };
        } catch (error) {
            logger.error(`Error while saving file: ${dest} ::: ${error.message}`);
            return { success: false, path: dest, errorMsg: error.message };
        }
    }

    async updateExcelTable(excelPath, tableName, values) {
        const sp = await this.appConfig.getSpConfig();
        // URI is set to the graybox sharepoint location where the promote project excel is created
        const itemId = await this.getItemId(sp.api.file.get.gbBaseURI, excelPath);
        if (itemId) {
            return this.executeGQL(`${sp.api.excel.update.baseItemsURI}/${itemId}/workbook/tables/${tableName}/rows`, {
                body: JSON.stringify({ values }),
                method: sp.api.excel.update.method,
            });
        }
        return {};
    }

    // fetch-with-retry added to check for Sharepoint RateLimit headers and 429 errors and to handle them accordingly.
    async fetchWithRetry(apiUrl, options, retryCounts) {
        let retryCount = retryCounts || 0;
        return new Promise((resolve, reject) => {
            const currentTime = Date.now();
            if (retryCount > NUM_REQ_THRESHOLD) {
                reject();
            } else if (nextCallAfter !== 0 && currentTime < nextCallAfter) {
                setTimeout(() => this.fetchWithRetry(apiUrl, options, retryCount)
                    .then((newResp) => resolve(newResp))
                    .catch((err) => reject(err)), nextCallAfter - currentTime);
            } else {
                retryCount += 1;
                logger.info(`Fetching with retry: ${apiUrl}`);
                logger.info(`Options in fetchWithRetry: ${util.inspect(options, { depth: null, colors: true })}`);
                fetch(apiUrl, options).then((resp) => {
                    this.logHeaders(resp);
                    const retryAfter = resp.headers.get('ratelimit-reset') || resp.headers.get('retry-after') || 0;
                    if ((resp.headers.get('test-retry-status') === TOO_MANY_REQUESTS) || (resp.status === TOO_MANY_REQUESTS)) {
                        nextCallAfter = Date.now() + retryAfter * 1000;
                        logger.info(`Retry ${nextCallAfter}`);
                        this.fetchWithRetry(apiUrl, options, retryCount)
                            .then((newResp) => resolve(newResp))
                            .catch((err) => reject(err));
                    } else {
                        nextCallAfter = retryAfter ? Math.max(Date.now() + retryAfter * 1000, nextCallAfter) : nextCallAfter;
                        resolve(resp);
                    }
                }).catch((err) => {
                    logger.warn(`Connection error ${apiUrl} with ${JSON.stringify(err)}`);
                    if (err && SP_CONN_ERR_LST.includes(err.code) && retryCount < NUM_REQ_THRESHOLD) {
                        logger.info(`Retry ${SP_CONN_ERR_LST}`);
                        nextCallAfter = Date.now() + RETRY_ON_CF * 1000;
                        return this.fetchWithRetry(apiUrl, options, retryCount)
                            .then((newResp) => resolve(newResp))
                            .catch((err2) => reject(err2));
                    }
                    return reject(err);
                });
            }
        });
    }

    getHeadersStr(response) {
        const headers = {};
        response?.headers?.forEach((value, name) => {
            headers[name] = value;
        });
        return JSON.stringify(headers);
    }

    getLogRespHeader = () => LOG_RESP_HEADER;

    logHeaders(response) {
        if (!this.getLogRespHeader()) return;
        const hdrStr = this.getHeadersStr(response);
        const logStr = `Status is ${response.status} with headers ${hdrStr}`;

        if (logStr.toUpperCase().indexOf('RATE') > 0 || logStr.toUpperCase().indexOf('RETRY') > 0) logger.info(logStr);
    }

    async checkFileExists(filePath, isGraybox = false) {
        const sp = await this.appConfig.getSpConfig();
        const options = await this.getAuthorizedRequestOption();
        const baseURI = isGraybox ? sp.api.file.get.gbBaseURI : sp.api.file.get.baseURI;
        const response = await this.fetchWithRetry(`${baseURI}${filePath}`, options);
        return response.ok;
    }

    async getFileMetadata(filePath, isGraybox = false) {
        const sp = await this.appConfig.getSpConfig();
        const options = await this.getAuthorizedRequestOption();
        const baseURI = isGraybox ? sp.api.file.get.gbBaseURI : sp.api.file.get.baseURI;
        const response = await this.fetchWithRetry(`${baseURI}${filePath}`, options);
        
        if (!response.ok) {
            logger.error(`Error getting file metadata for ${filePath}: ${response.status}`);
            return null;
        }

        const metadata = await response.json();
        return {
            createdDateTime: metadata.createdDateTime,
            lastModifiedDateTime: metadata.lastModifiedDateTime,
            path: `${metadata.parentReference.path.replace(/.*:.*:/, '')}/${metadata.name}`
        };
    }

    async bulkCopyFiles(sourcePaths, destinationPath, options = {}) {
        const results = {
            successful: [],
            failed: [],
            total: sourcePaths.length
        };

        for (const sourcePath of sourcePaths) {
            try {
                logger.info(`Processing file: ${sourcePath}`);
                
                // Get file data from source
                const { fileDownloadUrl, fileSize } = await this.getFileData(sourcePath, false);
                if (!fileDownloadUrl) {
                    throw new Error(`Failed to get file data for: ${sourcePath}`);
                }

                // Download the file
                const fileContent = await this.getFileUsingDownloadUrl(fileDownloadUrl);
                if (!fileContent) {
                    throw new Error(`Failed to download file: ${sourcePath}`);
                }

                // Prepare destination path
                const fileName = sourcePath.split('/').pop();
                const destPath = `${destinationPath}/${fileName}`;
                
                // Save the file to destination
                const saveResult = await this.saveFileSimple(fileContent, destPath, true);
                if (!saveResult.success) {
                    throw new Error(saveResult.errorMsg || `Failed to save file to: ${destPath}`);
                }

                results.successful.push({
                    sourcePath,
                    destinationPath: destPath,
                    fileSize,
                    metadata: saveResult.metadata
                });

                // Add a small delay between operations if specified
                if (options.delayBetweenOperations) {
                    await new Promise(resolve => setTimeout(resolve, options.delayBetweenOperations));
                }
            } catch (error) {
                logger.error(`Error processing ${sourcePath}: ${error.message}`);
                results.failed.push({
                    sourcePath,
                    error: error.message
                });
            }
        }

        return results;
    }
}

export default Sharepoint;
