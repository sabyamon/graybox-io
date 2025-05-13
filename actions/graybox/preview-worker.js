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

import { getAioLogger, handleExtension, toUTCStr } from '../utils.js';
import AppConfig from '../appConfig.js';
import HelixUtils from '../helixUtils.js';
import Sharepoint from '../sharepoint.js';
import initFilesWrapper from './filesWrapper.js';

const logger = getAioLogger();

async function main(params) {
    logger.info('Graybox Preview Action triggered');

    const appConfig = new AppConfig(params);
    const { gbRootFolder, experienceName, projectExcelPath } = appConfig.getPayload();

    const sharepoint = new Sharepoint(appConfig);

    // process data in batches
    const helixUtils = new HelixUtils(appConfig);
    // Batch Name to Array of Batch Preview Statuses mapping
    const previewStatuses = {};
    const filesWrapper = await initFilesWrapper(logger);
    let responsePayload;

    const project = `${gbRootFolder}/${experienceName}`;

    // Read the Project Status in the current project's "status.json" file
    const projectStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/status.json`);

    if (helixUtils.canBulkPreview(true)) {
        logger.info('In Preview Worker, Bulk Previewing Graybox files');
        if ((projectStatusJson.status === 'initiated' || projectStatusJson.status === 'promoted')) {
            try {
                let excelValues = '';
                if (projectStatusJson.status === 'initiated') {
                    excelValues = [[`Initial Preview started for '${experienceName}' experience`, toUTCStr(new Date()), '', '']];
                } else if (projectStatusJson.status === 'promoted') {
                    excelValues = [[`Final Preview started for promoted content of '${experienceName}' experience`, toUTCStr(new Date()), '', '']];
                }
                // Update Preview Status
                await sharepoint.updateExcelTable(projectExcelPath, 'PROMOTE_STATUS', excelValues);
            } catch (err) {
                logger.error(`Error Occured while updating Excel before starting Graybox Preview: ${err}`);
            }
            const batchStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/batch_status.json`);

            logger.info(`In Preview-Worker, for project: ${project} batchStatusJson: ${JSON.stringify(batchStatusJson)}`);

            const noofbatches = batchStatusJson !== undefined ? Object.keys(batchStatusJson).length : 0;
            // iterate over batch_status.json file and process each batch
            if (projectStatusJson.status === 'initiated') {
                const toBeStatus = 'initial_preview_in_progress';

                projectStatusJson?.statuses?.push({
                    step: 'Initial Preview of Graybox started',
                    stepName: 'initial_preview_in_progress',
                    timestamp: toUTCStr(new Date()),
                    files: []
                });

                // Update the In Progress Status in the current project's "status.json" file
                projectStatusJson.status = toBeStatus;
                await filesWrapper.writeFile(`graybox_promote${project}/status.json`, projectStatusJson);

                // Update the Project Status in the parent "project_queue.json" file
                await changeProjectStatusInQueue(filesWrapper, gbRootFolder, experienceName, toBeStatus);

                // Perform Initial Preview
                const batchResults = {};
                // Read the Batch JSON file into an array
                const i = 0; // Start with counter as 0
                const isGraybox = true;
                await iterateAndPreviewBatchJson(i, batchResults, noofbatches, batchStatusJson, isGraybox);

                // Reattempt Preview for failed preview paths
                await retryFailedPreviews(isGraybox);
            } else if (projectStatusJson.status === 'promoted') {
                // Update the In Progress Status in the current project's "status.json" file
                projectStatusJson.status = 'final_preview_in_progress';
                projectStatusJson?.statuses?.push({
                    step: 'Final Preview of Promoted Content started',
                    stepName: 'final_preview_in_progress',
                    timestamp: toUTCStr(new Date()),
                    files: []
                });
                await filesWrapper.writeFile(`graybox_promote${project}/status.json`, projectStatusJson);

                // Perform Final Preview
                const promotedPathsJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/promoted_paths.json`);
                const i = 0; // Start with counter as 0
                const isGraybox = false;
                await iterateAndPreviewBatchJson(i, promotedPathsJson, noofbatches, batchStatusJson, isGraybox);
                // Reattempt Preview for failed preview paths
                await retryFailedPreviews(isGraybox);
            }

            // Write the updated batch_status.json file
            await filesWrapper.writeFile(`graybox_promote${project}/batch_status.json`, batchStatusJson);
            logger.info(`In Preview Worker, for project: ${project} Updated Batch Status Json: ${JSON.stringify(batchStatusJson)}`);
            logger.info(`In Preview Worker, for project: ${project} Preview Statuses: ${JSON.stringify(previewStatuses)}`);

            // PreviewStatuses is an object with keys(batchNames) mapping to arrays(previewStauses)
            const failedPreviews = Object.keys(previewStatuses).reduce((acc, key) => {
                const filteredStatuses = previewStatuses[key]
                    .filter((status) => !status.success) // Filter out failed statuses
                    .map((status) => status.path); // Map to get the path of the failed status
                return acc.concat(filteredStatuses); // Concatenate to the accumulator
            }, []);
            // Now failedPreviews contains all the paths from the filtered and mapped arrays

            const previewStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/preview_status.json`);
            const previewErrorsJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/preview_errors.json`);

            // Combine the Preview Statuses for each batch read from AIO Json with the Preview Statuses
            if (previewStatusJson) {
                Object.entries(previewStatusJson).forEach(([batchName, batchPreviewStatuses]) => {
                    if (previewStatuses[batchName]) {
                        previewStatuses[batchName] = previewStatuses[batchName].concat(batchPreviewStatuses);
                    } else {
                        previewStatuses[batchName] = batchPreviewStatuses;
                    }
                });
            }

            // Write the updated preview_errors.json file
            await filesWrapper.writeFile(`graybox_promote${project}/preview_status.json`, previewStatuses);

            // Write the updated preview_errors.json file
            await filesWrapper.writeFile(`graybox_promote${project}/preview_errors.json`, previewErrorsJson.concat(failedPreviews));

            // Update the Project Status in the current project's "status.json" file & the parent "project_queue.json" file
            await updateProjectStatus(project, filesWrapper, previewStatuses, failedPreviews);

            try {
                logger.info('Updating project excel file with status');
                let excelValues = '';
                if (projectStatusJson.status === 'initial_preview_in_progress') {
                    const sFailedPreviews = failedPreviews.length > 0 ?
                        `Failed Previews(Please preview these files individually or with Milo Bulk Preview tool, and trigger Promote): \n${failedPreviews.join('\n')}` : '';
                    excelValues = [['Step 1 of 5: Initial Preview of Graybox completed', toUTCStr(new Date()), sFailedPreviews, '']];
                } else if (projectStatusJson.status === 'final_preview_in_progress') {
                    const sFailedPreviews = failedPreviews.length > 0 ? `Failed Previews: \n${failedPreviews.join('\n')}` : '';
                    excelValues = [['Step 5 of 5: Final Preview of Promoted Content completed', toUTCStr(new Date()), sFailedPreviews, '']];
                }
                // Update Preview Status
                await sharepoint.updateExcelTable(projectExcelPath, 'PROMOTE_STATUS', excelValues);

                // Write status to status.json
                // TODO: Check again
                /* const statusJsonPath = `graybox_promote/bacom-graybox/${experienceName}/status.json`;
                let statusJson = {};
                try {
                    statusJson = await filesWrapper.readFileIntoObject(statusJsonPath);
                } catch (err) {
                    // If file doesn't exist, create new object
                    statusJson = { statuses: [] };
                }
                
                // Add new status entry
                if (projectStatusJson.status === 'initial_preview_in_progress') {
                    statusJson.statuses.push({
                        step: 'Initial Preview Of Graybox Content Completed',
                        stepName: 'initial_preview_done',
                        timestamp: toUTCStr(new Date()),
                        failures: failedPreviews.length > 0 ? 
                            `Failed Previews(Please preview these files individually or with Milo Bulk Preview tool, and trigger Promote): \n${failedPreviews.join('\n')}` : '',
                        files: previewStatuses.map(status => status.path)
                    });
                } else if (projectStatusJson.status === 'final_preview_in_progress') {
                    statusJson.statuses.push({
                        step: 'Final Preview Of Promoted Content Completed',
                        stepName: 'final_preview_done',
                        timestamp: toUTCStr(new Date()),
                        failures: failedPreviews.length > 0 ? `Failed Previews: \n${failedPreviews.join('\n')}` : '',
                        files: []
                    });
                } */
                
                await filesWrapper.writeFile(statusJsonPath, statusJson);
            } catch (err) {
                logger.error(`Error Occured while updating Excel during Graybox Preview: ${err}`);
            }
        }

        responsePayload = 'Graybox Preview Worker action completed.';
    } else {
        responsePayload = 'Bulk Preview not enabled for Graybox Content Tree';
        logger.error(responsePayload);
    }
    logger.info(responsePayload);
    return exitAction({
        body: responsePayload,
        statusCode: 200
    });

    async function retryFailedPreviews(isGraybox) {
        try {
            const failedPreviewPathToBatchMap = {};
            // Find the failed previews from the previewStatuses object and re-attempt preview for them
            // PreviewStatuses is an object with keys(batchNames) mapping to arrays(previewStauses)
            // Create a fresh object subset of previewStatuses with failed preview statuses
            Object.keys(previewStatuses).forEach((batch) => {
                previewStatuses[batch].forEach((status) => {
                    if (!status.success) {
                        failedPreviewPathToBatchMap[status.path] = batch;
                    }
                });
            });

            const failedPreviewPaths = Object.keys(failedPreviewPathToBatchMap);

            logger.info(`Failed paths from First Preview attempt for project: ${project} ::: ${failedPreviewPaths}`);
            // Re-attempt preview for failed previews
            const reAttemptPreviewStatuses = await helixUtils.bulkPreview(failedPreviewPaths, helixUtils.getOperations().PREVIEW, experienceName, isGraybox);

            logger.info(`Reattempt Preview Statuses for project: ${project} :::: ${JSON.stringify(reAttemptPreviewStatuses)}`);
            // Update the previewStatuses object with the re-attempted preview statuses
            reAttemptPreviewStatuses.forEach((reattemptedStatus) => {
                // Extract the batch name for a specific path
                const batch = failedPreviewPathToBatchMap[reattemptedStatus.path];

                // Find the batch that contains the reattempted status path
                previewStatuses[batch].forEach((status, index) => {
                    if (status.path === reattemptedStatus.path) {
                        // Update the status in the corresponding batch
                        previewStatuses[batch][index] = reattemptedStatus;
                    }
                });
            });
        } catch (err) {
            logger.error(`Error Occured while Reattempting Failed Previews: ${err}`);
        }
    }

    /**
     * Iterate over the Batch JSON files, read those into an array and perform Bulk Preview
     * @param {*} i counter
     * @param {*} batchResults batchResults array
     * @param {*} noofbatches total no of batches
     * @param {*} filesWrapper filesWrapper object
     * @param {*} gbRootFolder graybox root folder
     * @param {*} experienceName graybox experience name
     */
    async function iterateAndPreviewBatchJson(i, batchResults, noofbatches, batchStatusJson, isGraybox) {
        const batchName = `batch_${i + 1}`;
        if (i < noofbatches) {
            if (batchStatusJson[batchName] === 'initiated' || batchStatusJson[batchName] === 'promoted') {
                // Only for initial preview read the files from /batches/ folder,
                // Otherwise for final preview use the list passed as-is from copy-worker or promote-worker
                if (batchStatusJson[batchName] === 'initiated') {
                    // Read the Batch JSON file into an batchResults JSON object
                    const batchJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/batches/${batchName}.json`);
                    batchResults[`${batchName}`] = batchJson;
                }
                // Perform Bulk Preview of a Batch of Graybox files
                await previewBatch(batchName, batchResults, batchStatusJson, isGraybox);
            }

            // Recursively call the function to process the next batch
            await iterateAndPreviewBatchJson(i + 1, batchResults, noofbatches, batchStatusJson, isGraybox);
        }
    }

    /**
     * Perform a Bulk Preview on a Batch of Graybox files
     * @param {*} batchName batchName
     * @param {*} previewStatuses returned preview statuses
     * @param {*} helixUtils helixUtils object
     * @param {*} experienceName graybox experience name
     */
    async function previewBatch(batchName, batchResults, batchStatusJson, isGraybox = true) {
        const batchJson = batchResults[batchName];
        logger.info(`In Preview-worker, in previewBatch for Batch: ${batchName} Batch JSON: ${JSON.stringify(batchJson)}`);
        const paths = [];
        if (batchJson) {
            batchJson.forEach((gbFile) => paths.push(handleExtension(gbFile)));

            // Perform Bulk Preview of a Batch of Graybox files
            if (isGraybox) {
                previewStatuses[batchName] = await helixUtils.bulkPreview(paths, helixUtils.getOperations().PREVIEW, experienceName, isGraybox);
                batchStatusJson[batchName] = 'initial_preview_done';
            } else {
                // Don't pass experienceName for final preview
                previewStatuses[batchName] = await helixUtils.bulkPreview(paths, helixUtils.getOperations().PREVIEW);
                batchStatusJson[batchName] = 'final_preview_done';
            }
        }
    }
}

/**
 * Update the Project Status in the current project's "status.json" file & the parent "project_queue.json" file
 * @param {*} gbRootFolder graybox root folder
 * @param {*} experienceName graybox experience name
 * @param {*} filesWrapper filesWrapper object
 * @returns updated project status
 */
async function updateProjectStatus(project, filesWrapper, previewStatuses, failedPreviews) {
    const projectStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/status.json`);

    // Update the Project Status in the current project's "status.json" file
    // If the project status is 'initiated', set it to 'initial_preview_done', else if project status is 'promoted' set it to 'final_preview_done'
    let toBeStatus;
    if (projectStatusJson.status === 'initiated' || projectStatusJson.status === 'initial_preview_in_progress') {
        toBeStatus = 'initial_preview_done';
    } else if (projectStatusJson.status === 'promoted' || projectStatusJson.status === 'final_preview_in_progress') {
        toBeStatus = 'final_preview_done';
    }

    logger.info(`In Preview-sched Preview Statuses: ${JSON.stringify(previewStatuses)}`);
    if (toBeStatus) {
        projectStatusJson.status = toBeStatus;
        projectStatusJson.statuses.push({
            step: 'Preview Completed',
            stepName: toBeStatus,
            timestamp: toUTCStr(new Date()),
            files: [],
            failures: failedPreviews.length > 0 ? `Failed Previews: \n${failedPreviews.join('\n')}` : ''
        });
        logger.info(`In Preview-sched After Processing Preview, Project Status Json: ${JSON.stringify(projectStatusJson)}`);
        await filesWrapper.writeFile(`graybox_promote${project}/status.json`, projectStatusJson);

        // Update the Project Status in the parent "project_queue.json" file
        await changeProjectStatusInQueue(filesWrapper, project, toBeStatus);
    }
}

async function changeProjectStatusInQueue(filesWrapper, project, toBeStatus) {
    const projectQueue = await filesWrapper.readFileIntoObject('graybox_promote/project_queue.json');
    const index = projectQueue.findIndex((obj) => obj.projectPath === `${project}`);
    if (index !== -1) {
        // Replace the object at the found index
        projectQueue[index].status = toBeStatus;
        await filesWrapper.writeFile('graybox_promote/project_queue.json', projectQueue);
        logger.info(`In Preview-sched After Processing Preview, Project Queue Json: ${JSON.stringify(projectQueue)}`);
    }
}

function exitAction(resp) {
    return resp;
}

export { main };
