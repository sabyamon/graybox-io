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

// eslint-disable-next-line import/no-extraneous-dependencies
import initFilesWrapper from './filesWrapper.js';
import { getAioLogger } from '../utils.js';
/**
 * This Action Sets the project status to paused in Project Queue & the Project Status JSON of that project
 */
async function main(params) {
    const logger = getAioLogger();
    const filesWrapper = await initFilesWrapper(logger);
    let responsePayload = 'Graybox Pause Project in Project Queue action invoked';
    let responseCode = 200;
    logger.info(responsePayload);
    try {
        const { projectPath } = params;
        logger.info(`Project to be paused :: ${projectPath}`);
        const projectQueuePath = 'graybox_promote/project_queue.json';
        if (await filesWrapper.fileExists(projectQueuePath)) {
            const projectQueue = await filesWrapper.readFileIntoObject(projectQueuePath);
            if (projectQueue) {
                logger.info(`In Pause Project Action, Before pausing, Project Queue Json: ${JSON.stringify(projectQueue)}`);
                const index = projectQueue.findIndex((obj) => obj.projectPath === projectPath);
                if (index === -1) {
                    responsePayload = `No project with ${projectPath} path exists in the project queue`;
                    return {
                        code: responseCode,
                        payload: responsePayload,
                    };
                }
                projectQueue[index].status = 'paused';
                await filesWrapper.writeFile(projectQueuePath, projectQueue);
                const project = projectQueue[index].projectPath;
                logger.info(`In Pause Project Action, After pausing, Project Queue Json: ${JSON.stringify(projectQueue)}`);
                const projectStatusJson = await filesWrapper.readFileIntoObject(`graybox_promote${project}/status.json`);
                logger.info(`In Pause Graybox Project, Before Pausing Project Status Json: ${JSON.stringify(projectStatusJson)}`);
                projectStatusJson.status = 'paused';
                await filesWrapper.writeFile(`graybox_promote${project}/status.json`, projectStatusJson);

                // Write status to status.json
                const statusJsonPath = `graybox_promote/bacom-graybox/${project.split('/').pop()}/status.json`;
                let statusJson = {};
                try {
                    statusJson = await filesWrapper.readFileIntoObject(statusJsonPath);
                } catch (err) {
                    // If file doesn't exist, create new object
                    statusJson = { statuses: [] };
                }
                
                // Add new status entry
                statusJson.statuses.push({
                    step: 'Project paused',
                    timestamp: new Date().toISOString(),
                    projectPath: project
                });
                
                await filesWrapper.writeFile(statusJsonPath, statusJson);
            } else {
                responsePayload = `Project Queue empty. No project with ${projectPath} path exists`;
                return {
                    code: responseCode,
                    payload: responsePayload,
                };
            }
        } else {
            responsePayload = 'Project Queue file doesn\'t exist in AIO';
            return {
                code: responseCode,
                payload: responsePayload,
            };
        }
    } catch (err) {
        responsePayload = 'Unknown error occurred';
        logger.error(`${responsePayload}: ${err}`);
        responsePayload = err;
        responseCode = 500;
    }

    return {
        code: responseCode,
        payload: responsePayload,
    };
}

export { main };
