/* ***********************************************************************
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

import crypto from 'crypto';
import { strToArray, getAioLogger } from './utils.js';
import UrlInfo from './urlInfo.js';

const GRAPH_API = 'https://graph.microsoft.com/v1.0';

/**
 * This stores the Graybox configs.
 */
class AppConfig {
    constructor(params) {
        this.configMap = { payload: {} };
        if (params) {
            this.setAppConfig(params);
        }
    }

    setAppConfig(params) {
        const payload = this.getPayload();

        // These are payload parameters
        // eslint-disable-next-line no-underscore-dangle
        const headers = params.__ow_headers;
        payload.spToken = headers?.['user-token'] || params.spToken;
        payload.adminPageUri = params.adminPageUri;
        payload.projectExcelPath = params.projectExcelPath;
        payload.rootFolder = params.rootFolder;
        payload.gbRootFolder = params.gbRootFolder;
        payload.promoteIgnorePaths = strToArray(params.promoteIgnorePaths) || [];
        payload.driveId = params.driveId;
        payload.draftsOnly = params.draftsOnly;
        payload.experienceName = params.experienceName;

        // Bulk copy specific configuration
        payload.sourcePaths = params.sourcePaths || [];
        payload.destinationPath = params.destinationPath || '';
        payload.bulkCopyOptions = params.options || {};

        // These are from params set in the github configs
        this.configMap.spSite = params.spSite;
        this.configMap.spClientId = params.spClientId;
        this.configMap.spAuthority = params.spAuthority;
        this.configMap.clientId = params.clientId;
        this.configMap.tenantId = params.tenantId;
        this.configMap.certPassword = params.certPassword;
        this.configMap.certKey = params.certKey;
        this.configMap.certThumbprint = params.certThumbprint;
        this.configMap.enablePreview = this.getJsonFromStr(params.enablePreview, []);
        this.configMap.helixAdminApiKeys = this.getJsonFromStr(params.helixAdminApiKeys);
        this.configMap.bulkPreviewCheckInterval = parseInt(params.bulkPreviewCheckInterval || '30', 10);
        this.configMap.maxBulkPreviewChecks = parseInt(params.maxBulkPreviewChecks || '30', 10);
        this.configMap.groupCheckUrl = params.groupCheckUrl || 'https://graph.microsoft.com/v1.0/groups/{groupOid}/members?$count=true';
        this.configMap.grayboxUserGroups = this.getJsonFromStr(params.grayboxUserGroups, []);
        this.configMap.ignoreUserCheck = (params.ignoreUserCheck || '').trim().toLowerCase() === 'true';

        this.extractPrivateKey();

        payload.ext = {
            urlInfo: payload.adminPageUri ? new UrlInfo(payload.adminPageUri) : null
        };
    }

    getPayload() {
        return this.configMap.payload;
    }

    // Configs related methods
    getConfig() {
        const { payload, ...configMap } = this.configMap;
        return { ...configMap, payload: this.getPayload() };
    }

    getJsonFromStr(str, def = {}) {
        try {
            return JSON.parse(str);
        } catch (err) {
            // Mostly bad string ignored
            getAioLogger().debug(`Error while parsing ${str}`);
        }
        return def;
    }

    getMsalConfig() {
        const {
            clientId, tenantId, certPassword, pvtKey, certThumbprint,
        } = this.configMap;
        return {
            clientId, tenantId, certPassword, pvtKey, certThumbprint,
        };
    }

    getSpSite() {
        return this.configMap.spSite;
    }

    getPromoteIgnorePaths() {
        const pips = this.getPayload().promoteIgnorePaths;
        return [...pips, '/.milo', '/.helix', '/metadata.xlsx', '*/query-index.xlsx'];
    }

    extractPrivateKey() {
        if (!this.configMap.certKey) return;
        const decodedKey = Buffer.from(
            this.configMap.certKey,
            'base64'
        ).toString('utf-8');
        this.configMap.pvtKey = crypto
            .createPrivateKey({
                key: decodedKey,
                passphrase: this.configMap.certPassword,
                format: 'pem',
            })
            .export({
                format: 'pem',
                type: 'pkcs8',
            });
    }

    getUrlInfo() {
        return this.getPayload().ext.urlInfo;
    }

    isDraftOnly() {
        const { draftsOnly } = this.getPayload();
        if (draftsOnly === undefined) {
            return true;
        }
        if (typeof draftsOnly === 'string') {
            return draftsOnly.trim().toLowerCase() !== 'false';
        }
        return draftsOnly;
    }

    ignoreUserCheck() {
        return true && this.configMap.ignoreUserCheck;
    }

    getUserToken() {
        return this.getPayload().spToken;
    }

    getSpConfig() {
        if (!this.getUrlInfo().isValid()) {
            logger.info('Invalid URL info');
            return undefined;
        }

        const config = this.getConfig();

        // get drive id if available
        const { driveId, rootFolder, gbRootFolder } = this.getPayload();
        const drive = driveId ? `/drives/${driveId}` : '/drive';

        const baseURI = `${config.spSite}${drive}/root:${rootFolder}`;
        const gbBaseURI = `${config.spSite}${drive}/root:${gbRootFolder}`;
        const baseItemsURI = `${config.spSite}${drive}/items`;
        return {
            api: {
                url: GRAPH_API,
                file: {
                    get: { baseURI, gbBaseURI },
                    download: { baseURI: `${config.spSite}${drive}/items` },
                    upload: {
                        baseURI,
                        gbBaseURI,
                        method: 'PUT',
                    },
                    delete: {
                        baseURI,
                        gbBaseURI,
                        method: 'DELETE',
                    },
                    update: {
                        baseURI,
                        gbBaseURI,
                        method: 'PATCH',
                    },
                    createUploadSession: {
                        baseURI,
                        gbBaseURI,
                        method: 'POST',
                        payload: { '@microsoft.graph.conflictBehavior': 'replace' },
                    },
                    copy: {
                        baseURI,
                        gbBaseURI,
                        method: 'POST',
                        payload: { '@microsoft.graph.conflictBehavior': 'replace' },
                    },
                },
                directory: {
                    create: {
                        baseURI,
                        gbBaseURI,
                        method: 'PATCH',
                        payload: { folder: {} },
                    },
                },
                excel: {
                    get: { baseItemsURI },
                    update: {
                        baseItemsURI,
                        method: 'POST',
                    },
                },
                batch: { uri: `${GRAPH_API}/$batch` },
            },
        };
    }

    // New methods for bulk copy support
    getSourcePaths() {
        return this.getPayload().sourcePaths;
    }

    getDestinationPath() {
        return this.getPayload().destinationPath;
    }

    getBulkCopyOptions() {
        return this.getPayload().bulkCopyOptions;
    }
}

export default AppConfig;
