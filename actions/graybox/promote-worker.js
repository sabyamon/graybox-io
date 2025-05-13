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
    logger.info('Graybox Promote Content Action triggered');

    const appConfig = new AppConfig(params);
    const { gbRootFolder, experienceName, projectExcelPath } = appConfig.getPayload();

    const sharepoint = new Sharepoint(appConfig);

    // process data in batches
    const filesWrapper = await initFilesWrapper(logger);
    let responsePayload;
    let promotes = [];
    const failedPromotes = [];
    const newerDestinationFiles = [];

    const project = params.project || '';
    const batchName = params.batchName || '';

    logger.info(`In Promote Content Worker, Processing Promote Content for batch: ${batchName}`);

    // Read the Batch Status in the current project's "batch_status.json" file
    let batchStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/batch_status.json`);

    const promoteErrorsJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promote_errors.json`);

    let promoteBatchesJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promote_batches.json`);
    logger.info(`In Promote-worker, for project: ${project} Promote Batches Json: ${JSON.stringify(promoteBatchesJson)}`);

    const toBeStatus = 'promote_in_progress';
    // Update & Write the Batch Status to in progress "batch_status.json" file
    // So that the scheduler doesn't pick the same batch again
    batchStatusJson[batchName] = toBeStatus;
    await filesWrapper.writeFile(`graybox_promote${project}/batch_status.json`, batchStatusJson);

    await changeProjectStatusInQueue(filesWrapper, project, toBeStatus);

    if (!promoteBatchesJson || !promoteBatchesJson[batchName]) {
        responsePayload = `No batch found to promote in Promote Batches JSON for Batch Name: ${batchName} for project`;
        logger.info(responsePayload);
        return exitAction({
            body: responsePayload,
            statusCode: 200
        });
    }

    promoteBatchesJson[batchName].status = 'promote_in_progress';
    await filesWrapper.writeFile(`graybox_promote${project}/promote_batches.json`, promoteBatchesJson);

    const promoteFilePaths = promoteBatchesJson[batchName].files || [];

    logger.info(`In Promote Content Worker, for project: ${project} for Batch Name ${batchName} promoteFilePaths: ${JSON.stringify(promoteFilePaths)}`);
    // Process the Promote Content
    // Collect all promises from the forEach loop
    // eslint-disable-next-line no-restricted-syntax
    for (const promoteFilePath of promoteFilePaths) {
        // Check if the file is a docx or xlsx based on file extension
        const isExcelFile = promoteFilePath.toLowerCase().endsWith('.xlsx') || promoteFilePath.toLowerCase().endsWith('.xls');
        const folderType = isExcelFile ? 'excel' : 'docx';
        // eslint-disable-next-line no-await-in-loop
        const promoteFile = await filesWrapper.readFileIntoBuffer(`graybox_promote${project}/${folderType}${promoteFilePath}`);
        if (promoteFile) {
            // Check file existence and compare dates
            const { newerDestinationFiles: newFiles } = await checkAndCompareFileDates({
                sharepoint,
                filesWrapper,
                project,
                filePath: promoteFilePath
            });
            newerDestinationFiles.push(...newFiles);
            
            // If file doesn't exist or we're overwriting it anyway
            const saveStatus = await sharepoint.saveFileSimple(promoteFile, promoteFilePath);

            if (saveStatus?.success) {
                promotes.push(promoteFilePath);
            } else if (saveStatus?.errorMsg?.includes('File is locked')) {
                failedPromotes.push(`${promoteFilePath} (locked file)`);
            } else {
                failedPromotes.push(promoteFilePath);
            }
        }
    }

    // Update the Excel with newer destination files
    if (newerDestinationFiles.length > 0) {
        await updateExcelWithNewerFiles({
            sharepoint,
            projectExcelPath,
            newerDestinationFiles,
            workerType: 'promote',
            experienceName,
            filesWrapper
        });
    }

    // Wait for all the promises to resolve

    // Update the Promoted Paths in the current project's "promoted_paths.json" file
    if (promotes.length > 0) {
        const promotedPathsJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promoted_paths.json`) || {};
        // Combined existing If any promotes already exist in promoted_paths.json for the current batch either from Copy action or Promote Action
        if (promotedPathsJson[batchName]) {
            promotes = promotes.concat(promotedPathsJson[batchName]);
        }
        promotedPathsJson[batchName] = promotes;
        await filesWrapper.writeFile(`graybox_promote${project}/promoted_paths.json`, promotedPathsJson);
    }

    if (failedPromotes.length > 0) {
        await filesWrapper.writeFile(`graybox_promote${project}/promote_errors.json`, promoteErrorsJson.concat(failedPromotes));
    }

    // Update the Promote Batch Status in the current project's "promote_batches.json" file
    promoteBatchesJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promote_batches.json`);
    promoteBatchesJson[batchName].status = 'promoted';
    // Write the promote batches JSON file
    await filesWrapper.writeFile(`graybox_promote${project}/promote_batches.json`, promoteBatchesJson);

    // Check in parallel if the Same Batch Name Exists & is Promoted in the Copy Batches JSON
    const copyBatchesJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/copy_batches.json`);
    const copyBatchJson = copyBatchesJson[batchName];
    let markBatchAsPromoted = true;
    if (copyBatchJson) {
        markBatchAsPromoted = copyBatchJson.status === 'promoted';
    }
    batchStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/batch_status.json`);
    if (markBatchAsPromoted) {
        // Update the Batch Status in the current project's "batch_status.json" file
        if (batchStatusJson && batchStatusJson[batchName] && (promotes.length > 0 || failedPromotes.length > 0)) {
            batchStatusJson[batchName] = 'promoted';
            // Write the updated batch_status.json file
            await filesWrapper.writeFile(`graybox_promote${project}/batch_status.json`, batchStatusJson);
        }

        // Find if the current batch running is the Last Copy Batch Name, and then mark the project as 'promoted'
        const allBatchesPromoted = Object.keys(batchStatusJson).every((key) => batchStatusJson[key] === 'promoted');
        if (allBatchesPromoted) {
            // Update the Project Status in JSON files
            updateProjectStatus(project, filesWrapper);
        }
    }

    // Update the Project Excel with the Promote Status
    try {
        const sFailedPromoteStatuses = failedPromotes.length > 0 ? `Failed Promotes: \n${failedPromotes.join('\n')}` : '';
        const promoteExcelValues = [[`Step 3 of 5: Promote completed for Batch ${batchName}`, toUTCStr(new Date()), sFailedPromoteStatuses, JSON.stringify(promotes)]];
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
            step: `Step 3 of 5: Promote completed for Batch ${batchName}`,
            stepName: 'promoted',
            timestamp: toUTCStr(new Date()),
            failures: sFailedPromoteStatuses,
            files: promotes
        });
        
        await filesWrapper.writeFile(statusJsonPath, statusJson);
    } catch (err) {
        logger.error(`Error Occured while updating Excel during Graybox Promote: ${err}`);
    }

    logger.info(`In Promote Content Worker, for project: ${project} Promotes: ${JSON.stringify(promotes)}`);
    logger.info(`In Promote Content Worker, for project: ${project} Failed Promotes: ${JSON.stringify(failedPromotes)}`);

    responsePayload = `Promote Content Worker finished promoting content, for project: ${project} for batch ${batchName}`;
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
async function updateProjectStatus(project, filesWrapper) {
    const projectStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/status.json`);

    const toBeStatus = 'promoted';
    // Update the Project Status in the current project's "status.json" file
    projectStatusJson.status = toBeStatus;
    logger.info(`In Promote-content-worker After Processing Promote, Project Status Json: ${JSON.stringify(projectStatusJson)}`);
    await filesWrapper.writeFile(`graybox_promote${project}/status.json`, projectStatusJson);

    // Update the Project Status in the parent "project_queue.json" file
    const projectQueue = await changeProjectStatusInQueue(filesWrapper, project, toBeStatus);
    logger.info(`In Promote-content-worker After Processing Promote, Project Queue Json: ${JSON.stringify(projectQueue)}`);
    await filesWrapper.writeFile('graybox_promote/project_queue.json', projectQueue);
}

async function changeProjectStatusInQueue(filesWrapper, project, toBeStatus) {
    const projectQueue = await filesWrapper.readFileIntoObject('graybox_promote/project_queue.json');
    const index = projectQueue.findIndex((obj) => obj.projectPath === `${project}`);
    if (index !== -1) {
        // Replace the object at the found index
        projectQueue[index].status = toBeStatus;
    }
    return projectQueue;
}

function exitAction(resp) {
    return resp;
}

export { main };
