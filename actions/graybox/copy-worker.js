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

import { getAioLogger, toUTCStr } from '../utils.js';
import AppConfig from '../appConfig.js';
import Sharepoint from '../sharepoint.js';
import initFilesWrapper from './filesWrapper.js';
import { checkAndCompareFileDates, updateExcelWithNewerFiles } from './fileComparisonUtils.js';

const logger = getAioLogger();

async function main(params) {
    logger.info('Graybox Copy Content Action triggered');

    const appConfig = new AppConfig(params);
    const { gbRootFolder, experienceName, projectExcelPath } = appConfig.getPayload();

    const sharepoint = new Sharepoint(appConfig);

    // process data in batches
    const filesWrapper = await initFilesWrapper(logger);
    let responsePayload;
    let promotes = [];
    const failedPromotes = [];

    logger.info('In Copy Worker, Processing Copy Content');

    const project = params.project || '';
    const batchName = params.batchName || '';

    // Read the Batch Status in the current project's "batch_status.json" file
    let batchStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/batch_status.json`);

    const promoteErrorsJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promote_errors.json`);

    let copyBatchesJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/copy_batches.json`);

    const copyBatchJson = copyBatchesJson[batchName] || {};

    logger.info(`In Copy Worker, Copy File Paths for project: ${project} for batchname ${batchName}:  ${JSON.stringify(copyBatchJson)}`);

    // Update & Write the Batch Status to in progress "batch_status.json" file
    // So that the scheduler doesn't pick the same batch again
    batchStatusJson[batchName] = 'copy_in_progress';
    await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/batch_status.json`, batchStatusJson);
    // Write the copy batches JSON file
    copyBatchesJson[batchName].status = 'promote_in_progress';
    await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/copy_batches.json`, copyBatchesJson);

    // Process the Copy Content
    const copyFilePathsJson = copyBatchJson.files || [];
    const newerDestinationFiles = [];

    for (let i = 0; i < copyFilePathsJson.length; i += 1) {
        const copyPathsEntry = copyFilePathsJson[i];
        
        // Check file existence and compare dates
        const { newerDestinationFiles: newFiles } = await checkAndCompareFileDates({
            sharepoint,
            filesWrapper,
            project,
            filePath: copyPathsEntry.copyDestFilePath
        });
        newerDestinationFiles.push(...newFiles);

        // Download the grayboxed file and save it to default content location
        // eslint-disable-next-line no-await-in-loop
        const { fileDownloadUrl } = await sharepoint.getFileData(copyPathsEntry.copySourceFilePath, true);
        // eslint-disable-next-line no-await-in-loop
        const file = await sharepoint.getFileUsingDownloadUrl(fileDownloadUrl);
        // eslint-disable-next-line no-await-in-loop
        const saveStatus = await sharepoint.saveFileSimple(file, copyPathsEntry.copyDestFilePath);

        if (saveStatus?.success) {
            promotes.push(copyPathsEntry.copyDestFilePath);
        } else if (saveStatus?.errorMsg?.includes('File is locked')) {
            failedPromotes.push(`${copyPathsEntry.copyDestFilePath} (locked file)`);
        } else {
            failedPromotes.push(copyPathsEntry.copyDestFilePath);
        }
    }

    // Update the Excel with newer destination files
    if (newerDestinationFiles.length > 0) {
        await updateExcelWithNewerFiles({
            sharepoint,
            projectExcelPath,
            newerDestinationFiles,
            workerType: 'copy',
            experienceName,
            filesWrapper
        });
    }

    logger.info(`In Copy Worker, Promotes for project: ${project} for batchname ${batchName} no.of files ${promotes.length}, files list: ${JSON.stringify(promotes)}`);
    // Update the Promoted Paths in the current project's "promoted_paths.json" file
    if (promotes.length > 0) {
        const promotedPathsJson = await filesWrapper.readFileIntoObject(`graybox_promote${gbRootFolder}/${experienceName}/promoted_paths.json`) || {};
        // Combined existing If any promotes already exist in promoted_paths.json for the current batch either from Copy action or Promote Action
        if (promotedPathsJson[batchName]) {
            promotes = promotes.concat(promotedPathsJson[batchName]);
        }
        promotedPathsJson[batchName] = promotes;
        await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/promoted_paths.json`, promotedPathsJson);
    }

    if (failedPromotes.length > 0) {
        await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/promote_errors.json`, promoteErrorsJson.concat(failedPromotes));
    }

    // Update the Copy Batch Status in the current project's "copy_batches.json" file
    copyBatchesJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/copy_batches.json`);
    copyBatchesJson[batchName].status = 'promoted';
    // Write the copy batches JSON file
    await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/copy_batches.json`, copyBatchesJson);

    // Check in parallel if the Same Batch Name Exists & is Promoted in the Promote Batches JSON
    const promoteBatchesJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promote_batches.json`);
    const promoteBatchJson = promoteBatchesJson[batchName];
    let markBatchAsPromoted = true;
    if (promoteBatchJson) {
        markBatchAsPromoted = promoteBatchJson.status === 'promoted';
    }

    batchStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/batch_status.json`);
    if (markBatchAsPromoted) {
        // Update the Batch Status in the current project's "batch_status.json" file
        if (batchStatusJson && batchStatusJson[batchName] && (promotes.length > 0 || failedPromotes.length > 0)) {
            batchStatusJson[batchName] = 'promoted';
            // Write the updated batch_status.json file
            await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/batch_status.json`, batchStatusJson);
        }

        // If all batches are promoted, then mark the project as 'promoted'
        const allBatchesPromoted = Object.keys(batchStatusJson).every((key) => batchStatusJson[key] === 'promoted');
        if (allBatchesPromoted) {
            // Update the Project Status in JSON files
            updateProjectStatus(gbRootFolder, experienceName, filesWrapper);
        }
    }

    // Update the Project Excel with the Promote Status
    try {
        const sFailedPromoteStatuses = failedPromotes.length > 0 ? `Failed Promotes: \n${failedPromotes.join('\n')}` : '';
        const promoteExcelValues = [[`Step 4 of 5: Promote Copy completed for Batch ${batchName}`, toUTCStr(new Date()), sFailedPromoteStatuses, JSON.stringify(promotes)]];
        await sharepoint.updateExcelTable(projectExcelPath, 'PROMOTE_STATUS', promoteExcelValues);

        // Write status to status.json
        const statusJsonPath = `graybox_promote/bacom-graybox/${experienceName}/status.json`;
        let statusJson = {};
        try {
            statusJson = await filesWrapper.readFileIntoObject(statusJsonPath);
        } catch (err) {
            // If file doesn't exist, create new object
            statusJson = { statuses: [] };
        }
        
        // Add new status entry
        statusJson.statuses.push({
            stepName: 'promote_copy_completed',
            step: `Step 4 of 5: Promote Copy completed for Batch ${batchName}`,
            timestamp: toUTCStr(new Date()),
            failures: sFailedPromoteStatuses,
            files: promotes
        });
        
        await filesWrapper.writeFile(statusJsonPath, statusJson);
    } catch (err) {
        logger.error(`Error Occured while updating Excel during Graybox Promote Copy: ${err}`);
    }

    responsePayload = `Copy Worker finished promoting content for batch ${batchName}`;
    logger.info(responsePayload);
    return exitAction({
        body: responsePayload,
        statusCode: 200
    });
}

/**
 * Update the Project Status in the current project's "status.json" file & the parent "project_queue.json" file
 * @param {*} gbRootFolder graybox root folder
 * @param {*} experienceName graybox experience name
 * @param {*} filesWrapper filesWrapper object
 * @returns updated project status
 */
async function updateProjectStatus(gbRootFolder, experienceName, filesWrapper) {
    const projectStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${gbRootFolder}/${experienceName}/status.json`);

    // Update the Project Status in the current project's "status.json" file
    projectStatusJson.status = 'promoted';
    await filesWrapper.writeFile(`graybox_promote${gbRootFolder}/${experienceName}/status.json`, projectStatusJson);

    // Update the Project Status in the parent "project_queue.json" file
    const projectQueue = await filesWrapper.readFileIntoObject('graybox_promote/project_queue.json');
    const index = projectQueue.findIndex((obj) => obj.projectPath === `${gbRootFolder}/${experienceName}`);
    if (index !== -1) {
        // Replace the object at the found index
        projectQueue[index].status = 'promoted';
        await filesWrapper.writeFile('graybox_promote/project_queue.json', projectQueue);
    }
}

function exitAction(resp) {
    return resp;
}

export { main };
