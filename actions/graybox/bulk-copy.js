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

import { getAioLogger, strToArray } from '../utils.js';
import openwhisk from 'openwhisk';

async function main(params) {
    // create a Logger
    const logger = getAioLogger('bulk-copy', params.LOG_LEVEL || 'info');
    const ow = openwhisk();
    try {
        logger.info('Starting bulk copy operation');
        // check for missing request input parameters
        const requiredParams = ['sourcePaths'];

        const missingParams = requiredParams.filter(param => !params[param]);
        if (missingParams.length > 0) {
            return {
                statusCode: 400,
                body: {
                    error: `Missing required parameters: ${missingParams.join(', ')}`
                }
            };
        }

        const sourcePaths = strToArray(params.sourcePaths);
        if (!Array.isArray(sourcePaths) || sourcePaths.length === 0) {
            return {
                statusCode: 400,
                body: {
                    error: 'sourcePaths must be a non-empty array or comma-separated string'
                }
            };
        }

        try {
            // Process sourcePaths to extract the actual path from AEM URLs
            const processedSourcePaths = sourcePaths.map((path) => {
                // Check if the path is an AEM URL
                if (path.includes('aem.page')) {
                    // Extract the path after aem.page
                    const match = path.match(/aem\.page(\/.*?)(?:$|\s)/);
                    if (match && match[1]) {
                        // Add .docx extension if not present
                        if (!match[1].includes('.')) {
                            return {
                                sourcePath: `${match[1]}.docx`,
                                destinationPath: `/${params?.experienceName}${match[1]}.docx`
                            };
                        }
                        return {
                            sourcePath: match[1],
                            destinationPath: `/${params?.experienceName}${match[1]}`
                        };
                    }
                }
                return {
                    sourcePath: path,
                    destinationPath: `/${params?.experienceName}${path}`
                };
            });

            // Extract the destination folder structure from the first source path
            let destinationSubPath = '';
            if (processedSourcePaths.length > 0) {
                const firstSourcePath = sourcePaths[0];
                // Check if the path is an AEM URL
                if (firstSourcePath.includes('aem.page')) {
                    // Extract the path structure after the domain but before the filename
                    const match = firstSourcePath.match(/aem\.page\/(.*?)(?:\/[^\/]+(?:\.\w+)?)?$/);
                    if (match && match[1]) {
                        destinationSubPath = `/${match[1]}`;
                    }
                }
            }

            // Form the complete destination path by combining gbRootFolder with the extracted subpath
            const formattedDestinationPath = `/${params?.experienceName}${destinationSubPath}`;
            const workerResponse = await ow.actions.invoke({
                name: 'graybox/bulk-copy-worker',
                blocking: false,
                result: false,
                params: {
                    ...params,
                    sourcePaths: processedSourcePaths,
                    destinationPath: formattedDestinationPath
                }
            });

            return {
                statusCode: 200,
                body: {
                    pathDetails: processedSourcePaths,
                    destinationFolder: formattedDestinationPath,
                    message: 'Bulk copy operation started',
                    activationId: workerResponse.activationId,
                }
            };
        } catch (err) {
            const errorMessage = 'Failed to invoke graybox bulk-copy-worker action';
            logger.error(`${errorMessage}: ${err}`);
            return {
                statusCode: 500,
                body: {
                    error: errorMessage,
                    message: err.message
                }
            };
        }
    } catch (error) {
        logger.error(error);
        return {
            statusCode: 500,
            body: {
                error: 'Internal server error',
                message: error.message
            }
        };
    }
}

export { main };
