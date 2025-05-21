import { getAioLogger, strToArray } from '../utils.js';
import openwhisk from 'openwhisk';

async function main(params) {
    // create a Logger
    const logger = getAioLogger('bulk-copy', params.LOG_LEVEL || 'info');
    const ow = openwhisk();
    logger.info(`Params in Bulk Copy: ${JSON.stringify(params)}`);
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

        // Convert sourcePaths to array if it's a string
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
            const processedSourcePaths = sourcePaths.map(path => {
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
            
            logger.info(`Formed destination path: ${formattedDestinationPath}`);
            
            // Update the destination path in params
            // destinationPath = /graybox-test/sabya/drafts (creates a folder inside /bacom-graybox)
            // gbRootFolder = /bacom-graybox
            // rootFolder = /bacom

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
            
            logger.info(workerResponse);
            return {
                statusCode: 200,
                body: {
                    processedSourcePaths,
                    formattedDestinationPath,
                    message: 'Bulk copy operation started',
                    activationId: workerResponse.activationId,
                    sourcePaths,
                    destinationPath: params.destinationPath
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