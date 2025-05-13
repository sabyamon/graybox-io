/* ************************************************************************
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

import { getAioLogger, toUTCStr } from '../utils.js';

const logger = getAioLogger();

/**
 * Checks if a file exists in destination and compares its date with source
 * @param {Object} params Parameters object
 * @param {Object} params.sharepoint Sharepoint instance
 * @param {Object} params.filesWrapper FilesWrapper instance
 * @param {string} params.project Project name
 * @param {string} params.filePath File path to check
 * @returns {Promise<Object>} Object containing newerDestinationFiles array and updated masterListMetadata
 */
async function checkAndCompareFileDates({ sharepoint, filesWrapper, project, filePath }) {
    const newerDestinationFiles = [];
    let masterListMetadata;

    // Check if the file already exists in the destination
    const fileExists = await sharepoint.checkFileExists(filePath);
    
    if (fileExists) {
        // Log metadata of the existing file
        logger.info(`File already exists at ${filePath}, checking metadata`);
        try {
            const fileMetadata = await sharepoint.getFileMetadata(filePath);
            // Get the source metadata to compare with destination
            masterListMetadata = await filesWrapper.readFileIntoObject(`graybox_promote${project}/master_list_metadata.json`);
            logger.info(`Master List Metadata: ${JSON.stringify(masterListMetadata)}`);
            
            if (masterListMetadata) {
                // Initialize destinationMetadata array if it doesn't exist
                if (!masterListMetadata.destinationMetadata) {
                    masterListMetadata.destinationMetadata = [];
                }
                
                // Add the destination file metadata to the array
                masterListMetadata.destinationMetadata.push({
                    createdDateTime: fileMetadata.createdDateTime,
                    lastModifiedDateTime: fileMetadata.lastModifiedDateTime,
                    path: fileMetadata.path
                });

                const sourceObjects = masterListMetadata.sourceMetadata || [];
                // Find the source object where the path is included in fileMetadata.path
                const matchingSourceObject = sourceObjects.find(sourceObj => {
                    return fileMetadata.path.includes(sourceObj.path);
                });
                
                if (matchingSourceObject) {
                    logger.info(`Found matching source metadata for ${fileMetadata.path}: ${JSON.stringify(matchingSourceObject)}`);
                    const sourceCreatedDate = new Date(matchingSourceObject.createdDateTime);
                    const destLastModifiedDate = new Date(fileMetadata.lastModifiedDateTime);
                    
                    // Compare dates including time, minutes and seconds
                    if (destLastModifiedDate.getTime() > sourceCreatedDate.getTime()) { 
                        logger.info(`Destination file is newer than source file: 
                            Source created: ${matchingSourceObject.createdDateTime}, 
                            Destination last modified: ${fileMetadata.lastModifiedDateTime}, 
                            Path: ${fileMetadata.path}`);
                            
                        // Add to the array of newer destination files
                        newerDestinationFiles.push({
                            path: fileMetadata.path.replace(/^\/drives\/.*\/root:/, ''),
                            sourceCreatedDateTime: matchingSourceObject.createdDateTime,
                            destinationLastModifiedDateTime: fileMetadata.lastModifiedDateTime
                        });
                    } else {
                        logger.info(`Source file is newer than destination file: 
                            Source created: ${matchingSourceObject.createdDateTime}, 
                            Destination last modified: ${fileMetadata.lastModifiedDateTime}, 
                            Path: ${fileMetadata.path}`);
                    }
                } else {
                    logger.warn(`No matching source metadata found for ${fileMetadata.path}`);
                }
                
                // Write the updated metadata back to the file
                await filesWrapper.writeFile(`graybox_promote${project}/master_list_metadata.json`, masterListMetadata);
                logger.info(`Updated master_list_metadata.json with destination metadata for ${filePath}`);
            } else {
                logger.warn(`Could not find master_list_metadata.json for project ${project}`);
            }
            logger.info(`Existing file metadata: ${JSON.stringify(fileMetadata)}`);
        } catch (error) {
            logger.warn(`Failed to get metadata for existing file ${filePath}: ${error.message}`);
        }
    }

    return { newerDestinationFiles, masterListMetadata };
}

/**
 * Updates Excel with newer destination files information
 * @param {Object} params Parameters object
 * @param {Object} params.sharepoint Sharepoint instance
 * @param {string} params.projectExcelPath Project Excel path
 * @param {Array} params.newerDestinationFiles Array of newer destination files
 * @param {string} params.workerType Type of worker ('copy' or 'promote')
 * @param {string} params.experienceName Experience name
 * @param {Object} params.filesWrapper Files wrapper instance
 */
async function updateExcelWithNewerFiles({ sharepoint, projectExcelPath, newerDestinationFiles, workerType, experienceName, filesWrapper }) {
    const message = workerType === 'copy' ? 'Copying' : 'Promoting';
    if (newerDestinationFiles.length > 0) {
        try {
            const newerFilesExcelValues = [
                [`Newer destination files detected while ${message}`, toUTCStr(new Date()), 
                 `${newerDestinationFiles.length} files in destination are newer than source`, 
                 JSON.stringify(newerDestinationFiles.map(file => file.path))]
            ];
            await sharepoint.updateExcelTable(projectExcelPath, 'PROMOTE_STATUS', newerFilesExcelValues);

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
                step: `Newer destination files detected while ${message}`,
                timestamp: toUTCStr(new Date()),
                newerFiles: {
                    count: newerDestinationFiles.length,
                    files: newerDestinationFiles.map(file => file.path)
                }
            });
            
            await filesWrapper.writeFile(statusJsonPath, statusJson);
            logger.info(`${workerType.charAt(0).toUpperCase() + workerType.slice(1)} Worker: Updated project Excel with newer destination files information`);
        } catch (err) {
            logger.error(`Error occurred while updating Excel with newer destination files: ${err}`);
        }
    }
}

export { checkAndCompareFileDates, updateExcelWithNewerFiles };
