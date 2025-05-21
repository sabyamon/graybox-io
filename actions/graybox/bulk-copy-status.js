/* ***********************************************************************
 * ADOBE CONFIDENTIAL
 * ___________________
 *
 * Copyright 2025 Adobe
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

import initFilesWrapper from './filesWrapper.js';
import { getAioLogger } from '../utils.js';

function determineOverallStatus(statuses) {
    if (!statuses || statuses.length === 0) {
        return 'unknown';
    }

    // Get the latest status
    const latestStatus = statuses[statuses.length - 1];

    // Check for error status
    if (latestStatus.status === 'error') {
        return 'error';
    }

    // Check for completion
    if (latestStatus.status === 'completed') {
        return 'completed';
    }

    // Check for in-progress statuses
    if (['processing', 'processing_file', 'saving_file'].includes(latestStatus.status)) {
        return 'in_progress';
    }

    // Check for file-specific statuses
    if (['file_copied', 'file_failed'].includes(latestStatus.status)) {
        return 'in_progress';
    }

    // Check for started status
    if (latestStatus.status === 'started') {
        return 'started';
    }

    return 'unknown';
}

async function main(params) {
    const logger = getAioLogger();
    const filesWrapper = await initFilesWrapper(logger);

    let responsePayload = 'Graybox Bulk Copy Status API invoked';
    const responseCode = 200;
    logger.info(responsePayload);

    const { bacomGraybox, sabyaBulkCopy } = params;

    if (!bacomGraybox || !sabyaBulkCopy) {
        responsePayload = {
            error: 'Missing required parameters: bacomGraybox and sabyaBulkCopy'
        };
        return {
            code: 400,
            payload: responsePayload
        };
    }

    const path = `graybox_promote/${bacomGraybox}/${sabyaBulkCopy}`;
    logger.info(`Reading bulk copy status from path: ${path}`);

    try {
        const statusFile = await filesWrapper.readFileIntoObject(`${path}/bulk-copy-status.json`);
        
        // Update the main status based on the current statuses
        const newStatus = determineOverallStatus(statusFile.statuses);
        statusFile.status = newStatus;
        
        // Write back the updated status file
        await filesWrapper.writeFile(`${path}/bulk-copy-status.json`, statusFile);
        
        responsePayload = {
            path,
            status: statusFile.status,
            details: statusFile
        };
        
        return {
            code: responseCode,
            payload: responsePayload
        };
    } catch (error) {
        logger.error(`Error reading bulk copy status: ${error.message}`);
        responsePayload = {
            error: `Failed to read bulk copy status: ${error.message}`
        };
        return {
            code: 500,
            payload: responsePayload
        };
    }
}

export { main }; 